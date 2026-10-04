"""SSRF-hardened download of remote images.

``fetch_image`` only talks to globally routable hosts. Every hop (the first
request and each redirect) is re-validated:

* the URL must be ``http``/``https`` without user info;
* *every* address the host name resolves to must be public -- loopback,
  private, link-local (incl. the 169.254.169.254 metadata service), CGNAT,
  multicast, reserved and unspecified ranges are refused, also when wrapped in
  IPv4-mapped, NAT64, 6to4 or Teredo IPv6 forms;
* the connection goes to a vetted IP itself (no second DNS lookup, so DNS
  rebinding cannot swap in an internal address), with the original ``Host``
  header and, for HTTPS, SNI plus certificate verification against the
  original host name.

One wall-clock budget (``total_timeout``) covers connecting, the TLS
handshake, sending the request, the status line, the headers and the body of
every hop: each socket operation gets ``min(timeout, time left)``, so a host
trickling bytes cannot hold a fetch open. Redirects are followed manually,
response headers and the body are size-capped, and network and protocol
failures surface only as :class:`RemoteFetchError`.
"""

from __future__ import annotations

from collections.abc import Callable
import functools
import http.client
import io
import ipaddress
import socket
import ssl
import time
from dataclasses import dataclass
from typing import Any
from urllib.parse import urljoin, urlsplit

__all__ = ["RemoteFetchError", "fetch_image", "is_public_address"]

_REDIRECT_STATUSES = frozenset({301, 302, 303, 307, 308})
_DEFAULT_PORTS = {"http": 80, "https": 443}
_CHUNK_SIZE = 64 * 1024
# Status line plus all header lines; http.client separately caps the count at 100.
_MAX_HEADER_BYTES = 64 * 1024
_USER_AGENT = "MDtoWord/1.0 (image fetch)"
_NAT64_PREFIX = ipaddress.IPv6Network("64:ff9b::/96")

IPAddress = ipaddress.IPv4Address | ipaddress.IPv6Address
Budget = Callable[[], float]


class RemoteFetchError(Exception):
    """A remote image could not be fetched; the message is a short reason."""


@dataclass(frozen=True)
class _Target:
    """One validated request target."""

    scheme: str
    host: str  # ASCII host name or IP literal, without brackets
    port: int
    request_path: str  # path + query, never empty
    host_header: str


# ---------------------------------------------------------------------------
# Address policy
# ---------------------------------------------------------------------------


def _embedded_ipv4(address: ipaddress.IPv6Address) -> list[ipaddress.IPv4Address]:
    """IPv4 addresses a 6to4 / Teredo address tunnels to."""
    embedded: list[ipaddress.IPv4Address] = []
    if address.sixtofour is not None:
        embedded.append(address.sixtofour)
    if address.teredo is not None:
        embedded.extend(address.teredo)
    return embedded


def _is_public(address: IPAddress) -> bool:
    if isinstance(address, ipaddress.IPv6Address):
        # IPv4-mapped and NAT64 addresses reach exactly the embedded IPv4 host.
        translated = address.ipv4_mapped
        if translated is None and address in _NAT64_PREFIX:
            translated = ipaddress.IPv4Address(int(address) & 0xFFFFFFFF)
        if translated is not None:
            return _is_public(translated)
        if any(not _is_public(inner) for inner in _embedded_ipv4(address)):
            return False
        if address.is_site_local:
            return False
    # ``is_global`` alone lets multicast (224/4) and IPv4-compatible (::a.b.c.d)
    # addresses through, hence the explicit exclusions.
    return address.is_global and not (
        address.is_multicast
        or address.is_reserved
        or address.is_unspecified
        or address.is_loopback
        or address.is_link_local
        or address.is_private
    )


def is_public_address(address: str) -> bool:
    """True when ``address`` (an IP literal) is globally routable and allowed."""
    try:
        parsed = ipaddress.ip_address(address)
    except ValueError:
        return False
    return _is_public(parsed)


# ---------------------------------------------------------------------------
# URL and DNS validation
# ---------------------------------------------------------------------------


def _ascii_host(hostname: str) -> str:
    try:
        ipaddress.ip_address(hostname)
    except ValueError:
        pass
    else:
        return hostname
    try:
        ascii_host = hostname.encode("idna").decode("ascii")
    except UnicodeError as exc:
        raise RemoteFetchError("invalid host name") from exc
    if not ascii_host or any(char in ascii_host for char in "%/\\ "):
        raise RemoteFetchError("invalid host name")
    return ascii_host


def _parse_target(url: str) -> _Target:
    try:
        parts = urlsplit(url.strip())
        port = parts.port
    except ValueError as exc:
        raise RemoteFetchError("invalid URL") from exc
    scheme = parts.scheme.lower()
    if scheme not in _DEFAULT_PORTS:
        raise RemoteFetchError(f"unsupported URL scheme {scheme or '(none)'!r}")
    if "@" in parts.netloc or parts.username is not None or parts.password is not None:
        raise RemoteFetchError("URLs with credentials are not allowed")
    if not parts.hostname:
        raise RemoteFetchError("URL has no host")
    host = _ascii_host(parts.hostname)
    port = port or _DEFAULT_PORTS[scheme]
    path = parts.path or "/"
    if parts.query:
        path = f"{path}?{parts.query}"
    host_header = f"[{host}]" if ":" in host else host
    if port != _DEFAULT_PORTS[scheme]:
        host_header = f"{host_header}:{port}"
    return _Target(scheme, host, port, path, host_header)


def _vetted_addresses(target: _Target) -> list[str]:
    """Resolve ``target.host`` once; refuse it unless every address is public.

    Returns the addresses in resolver order, without duplicates.
    """
    try:
        infos = socket.getaddrinfo(
            target.host, target.port, type=socket.SOCK_STREAM, proto=socket.IPPROTO_TCP
        )
    except (OSError, UnicodeError) as exc:
        raise RemoteFetchError(f"cannot resolve host {target.host!r}") from exc
    addresses = list(dict.fromkeys(
        str(sockaddr[0])
        for family, _type, _proto, _canonname, sockaddr in infos
        if family in (socket.AF_INET, socket.AF_INET6)
    ))
    if not addresses:
        raise RemoteFetchError(f"host {target.host!r} has no usable address")
    for address in addresses:
        if not is_public_address(address):
            raise RemoteFetchError(f"host {target.host!r} resolves to a non-public address")
    return addresses


# ---------------------------------------------------------------------------
# Deadline-bounded I/O
# ---------------------------------------------------------------------------


class _Deadline:
    """Callable time budget: ``min(timeout, seconds left)``, TimeoutError once spent."""

    def __init__(self, timeout: float, total_timeout: float) -> None:
        self._timeout = timeout
        self._deadline = time.monotonic() + total_timeout

    def __call__(self) -> float:
        left = self._deadline - time.monotonic()
        if left <= 0:
            raise TimeoutError("timed out")
        return min(self._timeout, left)


class _DeadlineSocketReader(io.RawIOBase):
    """Raw reader that re-arms the socket timeout from the budget on every recv.

    ``keepalive`` is http.client's own ``makefile()`` reader: holding it open
    keeps the socket's io refcount up, so ``HTTPConnection.close()`` (called
    by ``getresponse`` for ``Connection: close`` replies) defers the real close
    until this response is closed -- exactly as with the stock reader.
    """

    def __init__(self, sock: socket.socket, budget: Budget, keepalive: Any) -> None:
        super().__init__()
        self._sock = sock
        self._budget = budget
        self._keepalive = keepalive

    def readable(self) -> bool:
        return True

    def close(self) -> None:
        if not self.closed:
            try:
                self._keepalive.close()
            finally:
                super().close()

    def readinto(self, buffer: Any) -> int:
        self._sock.settimeout(self._budget())
        return self._sock.recv_into(buffer)


class _HeaderCappedReader(io.BufferedReader):
    """Buffered reader whose ``readline`` enforces a byte budget while armed."""

    def __init__(self, raw: io.RawIOBase) -> None:
        super().__init__(raw)
        self.header_budget: int | None = None

    def readline(self, size: int | None = -1) -> bytes:
        budget = self.header_budget
        if budget is None:
            return super().readline(size)
        limit = budget + 1
        line = super().readline(limit if size is None or size < 0 else min(size, limit))
        self.header_budget = budget - len(line)
        if self.header_budget < 0:
            raise RemoteFetchError("response headers are too large")
        return line


class _DeadlineResponse(http.client.HTTPResponse):
    """HTTPResponse reading through the deadline-aware, header-capped reader."""

    def __init__(self, sock: socket.socket, *args: Any, budget: Budget, **kwargs: Any) -> None:
        super().__init__(sock, *args, **kwargs)
        self.fp = _HeaderCappedReader(_DeadlineSocketReader(sock, budget, keepalive=self.fp))

    def begin(self) -> None:
        self.fp.header_budget = _MAX_HEADER_BYTES
        try:
            super().begin()
        finally:
            self.fp.header_budget = None


class _PinnedHTTPSConnection(http.client.HTTPSConnection):
    """HTTPS to a pre-resolved IP, with SNI and certificate checks for ``server_hostname``."""

    def __init__(
        self,
        address: str,
        port: int,
        *,
        server_hostname: str,
        budget: Budget,
        context: ssl.SSLContext,
    ) -> None:
        super().__init__(address, port, timeout=budget(), context=context)
        self._server_hostname = server_hostname
        self._tls_context = context
        self._budget = budget

    def connect(self) -> None:
        http.client.HTTPConnection.connect(self)  # plain TCP to the vetted IP
        # The whole handshake is bounded by the socket timeout set here.
        self.sock.settimeout(self._budget())
        self.sock = self._tls_context.wrap_socket(
            self.sock, server_hostname=self._server_hostname
        )


_tls_context_cache: ssl.SSLContext | None = None


def _tls_context() -> ssl.SSLContext:
    """Default context: system CAs, CERT_REQUIRED, host name checking on."""
    global _tls_context_cache
    if _tls_context_cache is None:
        _tls_context_cache = ssl.create_default_context()
    return _tls_context_cache


def _open_connection(
    scheme: str, address: str, port: int, server_hostname: str, budget: Budget
) -> http.client.HTTPConnection:
    """Connection object for one hop; ``connect()`` is called by the caller."""
    connection: http.client.HTTPConnection
    if scheme == "https":
        connection = _PinnedHTTPSConnection(
            address, port, server_hostname=server_hostname, budget=budget, context=_tls_context()
        )
    else:
        connection = http.client.HTTPConnection(address, port, timeout=budget())
    connection.response_class = functools.partial(_DeadlineResponse, budget=budget)
    return connection


def _connect_any(target: _Target, addresses: list[str], budget: Budget) -> Any:
    """Connect to the first vetted address that accepts, within the budget.

    TCP-level failures move on to the next address; TLS failures do not (a
    bad certificate is not an address problem).
    """
    last_error: OSError | None = None
    for address in addresses:
        connection = None
        try:
            budget()
            connection = _open_connection(target.scheme, address, target.port, target.host, budget)
            connection.timeout = budget()
            connection.connect()
            return connection
        except ssl.SSLError:
            if connection is not None:
                _close_quietly(connection)
            raise
        except OSError as exc:
            if connection is not None:
                _close_quietly(connection)
            last_error = exc
    assert last_error is not None
    raise last_error


def _apply_timeout(connection: Any, seconds: float) -> None:
    sock = getattr(connection, "sock", None)
    if sock is not None:
        sock.settimeout(seconds)


def _close_quietly(connection: Any) -> None:
    try:
        connection.close()
    except Exception:  # closing must never mask the real outcome
        pass


def _describe(exc: BaseException) -> str:
    if isinstance(exc, TimeoutError):
        return "timed out"
    if isinstance(exc, ssl.SSLCertVerificationError):
        return "TLS certificate verification failed"
    if isinstance(exc, ssl.SSLError):
        return "TLS error"
    if isinstance(exc, ConnectionRefusedError):
        return "connection refused"
    if isinstance(exc, http.client.HTTPException):
        return "invalid HTTP response"
    if isinstance(exc, OSError):
        return f"network error: {exc.strerror or type(exc).__name__}"
    return f"request failed: {type(exc).__name__}"


# ---------------------------------------------------------------------------
# Fetch
# ---------------------------------------------------------------------------

_WRAPPED_ERRORS = (OSError, http.client.HTTPException, ValueError, UnicodeError)


def _read_body(response: Any, max_bytes: int, budget: Budget) -> bytes:
    declared = response.getheader("Content-Length")
    if declared is not None and declared.strip().isdigit() and int(declared) > max_bytes:
        raise RemoteFetchError(f"image is larger than {max_bytes} bytes")
    chunks: list[bytes] = []
    received = 0
    while True:
        budget()  # each recv is bounded too; this also stops between chunks
        chunk = response.read1(min(_CHUNK_SIZE, max_bytes - received + 1))
        if not chunk:
            return b"".join(chunks)
        received += len(chunk)
        if received > max_bytes:
            raise RemoteFetchError(f"image is larger than {max_bytes} bytes")
        chunks.append(chunk)


def _request_headers(target: _Target) -> dict[str, str]:
    return {
        "Host": target.host_header,
        "User-Agent": _USER_AGENT,
        "Accept": "image/*, */*;q=0.5",
        "Accept-Encoding": "identity",
        "Connection": "close",
    }


def fetch_image(
    url: str,
    *,
    max_bytes: int,
    timeout: float = 10.0,
    max_redirects: int = 5,
    total_timeout: float = 30.0,
) -> bytes:
    """Download ``url`` and return its body, refusing anything but public hosts.

    ``timeout`` bounds each socket operation and ``total_timeout`` the whole
    exchange including redirects, connects, TLS handshakes, headers and body.
    Only the DNS lookups themselves cannot be interrupted. Raises
    :class:`RemoteFetchError` for every network, protocol or policy failure,
    ``ValueError`` only for invalid arguments.
    """
    if max_bytes < 0 or max_redirects < 0 or timeout <= 0 or total_timeout < 0:
        raise ValueError("max_bytes/max_redirects must be >= 0 and timeouts positive")
    budget = _Deadline(timeout, total_timeout)
    current_url = url
    for _hop in range(max_redirects + 1):
        target = _parse_target(current_url)
        addresses = _vetted_addresses(target)
        connection = None
        try:
            connection = _connect_any(target, addresses, budget)
            _apply_timeout(connection, budget())
            connection.request("GET", target.request_path, headers=_request_headers(target))
            response = connection.getresponse()
            status = response.status
            if 200 <= status < 300:
                return _read_body(response, max_bytes, budget)
            if status not in _REDIRECT_STATUSES:
                raise RemoteFetchError(f"HTTP {status}")
            location = (response.getheader("Location") or "").strip()
            if not location:
                raise RemoteFetchError(f"HTTP {status} redirect without a Location header")
        except RemoteFetchError:
            raise
        except _WRAPPED_ERRORS as exc:
            raise RemoteFetchError(_describe(exc)) from exc
        finally:
            if connection is not None:
                _close_quietly(connection)
        current_url = urljoin(current_url, location)
    raise RemoteFetchError(f"too many redirects (limit {max_redirects})")
