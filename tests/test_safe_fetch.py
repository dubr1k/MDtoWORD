from __future__ import annotations

import http.client
import http.server
import socket
import ssl
import threading
import time
import unittest
from unittest.mock import MagicMock, patch

from mdtoword import safe_fetch
from mdtoword.safe_fetch import RemoteFetchError, fetch_image, is_public_address


_PUBLIC_V4 = "93.184.216.34"
_PUBLIC_V6 = "2606:2800:220:1:248:1893:25c8:1946"


def _addrinfo(*addresses: str):
    """Fake ``socket.getaddrinfo`` returning ``addresses`` for every host."""

    def resolver(host, port, *args, **kwargs):
        infos = []
        for address in addresses:
            if ":" in address:
                infos.append((socket.AF_INET6, socket.SOCK_STREAM, 6, "", (address, port, 0, 0)))
            else:
                infos.append((socket.AF_INET, socket.SOCK_STREAM, 6, "", (address, port)))
        return infos

    return resolver


def _resolver_by_host(mapping: dict[str, str]):
    def resolver(host, port, *args, **kwargs):
        return _addrinfo(mapping[host])(host, port)

    return resolver


class _FakeResponse:
    def __init__(self, status: int = 200, body: bytes = b"", headers: dict[str, str] | None = None):
        self.status = status
        self._body = body
        self._headers = {k.lower(): v for k, v in (headers or {}).items()}
        self.bytes_read = 0
        self.on_read = lambda: None

    def getheader(self, name: str, default=None):
        return self._headers.get(name.lower(), default)

    def read1(self, amount: int) -> bytes:
        self.on_read()
        chunk = self._body[self.bytes_read : self.bytes_read + min(amount, 7)]
        self.bytes_read += len(chunk)
        return chunk


class _FakeConnection:
    def __init__(self, response: _FakeResponse, connect_error: BaseException | None = None):
        self.response = response
        self.connect_error = connect_error
        self.requests: list[tuple[str, str, dict[str, str]]] = []
        self.closed = False
        self.sock = MagicMock()
        self.timeout = None

    def connect(self) -> None:
        if self.connect_error is not None:
            raise self.connect_error

    def request(self, method: str, path: str, headers: dict[str, str]) -> None:
        self.requests.append((method, path, headers))

    def getresponse(self) -> _FakeResponse:
        return self.response

    def close(self) -> None:
        self.closed = True


class _ConnectionFactory:
    """Stand-in for ``_open_connection`` serving canned responses in order."""

    def __init__(self, *responses: _FakeResponse, connect_errors: dict[str, BaseException] | None = None):
        self._responses = list(responses)
        self._connect_errors = connect_errors or {}
        self.calls: list[tuple[str, str, int, str]] = []
        self.connections: list[_FakeConnection] = []

    def __call__(self, scheme, address, port, server_hostname, budget):
        self.calls.append((scheme, address, port, server_hostname))
        error = self._connect_errors.get(address)
        if error is not None:
            connection = _FakeConnection(self._responses[0], connect_error=error)
            self.connections.append(connection)
            return connection
        response = self._responses.pop(0) if len(self._responses) > 1 else self._responses[0]
        connection = _FakeConnection(response)
        self.connections.append(connection)
        return connection


class AddressPolicyTests(unittest.TestCase):
    def test_blocked_addresses(self) -> None:
        for address in (
            "127.0.0.1", "10.1.2.3", "172.16.0.1", "192.168.1.1", "169.254.169.254",
            "100.64.0.1", "0.0.0.0", "224.0.0.1", "240.0.0.1", "255.255.255.255",
            "198.18.0.1", "::1", "::", "fd00::1", "fe80::1", "fe80::1%en0", "ff02::1",
            "::ffff:127.0.0.1", "::ffff:10.0.0.1", "::ffff:169.254.169.254", "::7f00:1",
            "64:ff9b::a00:1", "2002:7f00:1::", "2001:db8::1", "fec0::1", "not-an-ip",
        ):
            with self.subTest(address=address):
                self.assertFalse(is_public_address(address))

    def test_public_addresses(self) -> None:
        for address in (_PUBLIC_V4, "8.8.8.8", _PUBLIC_V6, "::ffff:8.8.8.8", "64:ff9b::808:808"):
            with self.subTest(address=address):
                self.assertTrue(is_public_address(address))


class FetchImageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.factory = _ConnectionFactory(_FakeResponse(200, b"PNGDATA" * 3))
        patcher = patch.object(safe_fetch, "_open_connection", self.factory)
        patcher.start()
        self.addCleanup(patcher.stop)

    def _resolve(self, *addresses: str):
        patcher = patch.object(safe_fetch.socket, "getaddrinfo", side_effect=_addrinfo(*addresses))
        mock = patcher.start()
        self.addCleanup(patcher.stop)
        return mock

    def test_happy_path_pins_vetted_ip_and_keeps_host(self) -> None:
        self._resolve(_PUBLIC_V4)
        data = fetch_image("https://Example.com/img.png?x=1#frag", max_bytes=1000)
        self.assertEqual(data, b"PNGDATA" * 3)
        self.assertEqual(self.factory.calls, [("https", _PUBLIC_V4, 443, "example.com")])
        method, path, headers = self.factory.connections[0].requests[0]
        self.assertEqual((method, path), ("GET", "/img.png?x=1"))
        self.assertEqual(headers["Host"], "example.com")
        self.assertEqual(headers["Accept-Encoding"], "identity")
        self.assertTrue(self.factory.connections[0].closed)

    def test_non_default_port_and_ipv6_host_header(self) -> None:
        self._resolve(_PUBLIC_V6)
        fetch_image(f"http://[{_PUBLIC_V6}]:8080/a.png", max_bytes=1000)
        self.assertEqual(self.factory.calls, [("http", _PUBLIC_V6, 8080, _PUBLIC_V6)])
        self.assertEqual(self.factory.connections[0].requests[0][2]["Host"], f"[{_PUBLIC_V6}]:8080")

    def test_idn_host_is_encoded(self) -> None:
        resolver = self._resolve(_PUBLIC_V4)
        fetch_image("http://пример.рф/a.png", max_bytes=1000)
        self.assertEqual(resolver.call_args[0][0], "xn--e1afmkfd.xn--p1ai")
        self.assertEqual(self.factory.connections[0].requests[0][2]["Host"], "xn--e1afmkfd.xn--p1ai")

    def test_private_addresses_rejected_before_connecting(self) -> None:
        for address in (
            "127.0.0.1", "10.0.0.7", "192.168.0.10", "169.254.169.254", "::1", "fd00::5",
            "::ffff:127.0.0.1", "100.64.0.1",
        ):
            with self.subTest(address=address), patch.object(
                safe_fetch.socket, "getaddrinfo", side_effect=_addrinfo(address)
            ):
                with self.assertRaisesRegex(RemoteFetchError, "non-public"):
                    fetch_image("http://images.example/a.png", max_bytes=1000)
        self.assertEqual(self.factory.calls, [])

    def test_rejected_when_any_resolved_address_is_private(self) -> None:
        self._resolve(_PUBLIC_V4, "10.0.0.1")
        with self.assertRaisesRegex(RemoteFetchError, "non-public"):
            fetch_image("http://images.example/a.png", max_bytes=1000)
        self.assertEqual(self.factory.calls, [])

    def test_literal_loopback_urls_rejected(self) -> None:
        for url in ("http://127.0.0.1/a.png", "http://[::1]/a.png", "http://2130706433/a.png"):
            with self.subTest(url=url):
                with self.assertRaises(RemoteFetchError):
                    fetch_image(url, max_bytes=1000)
        self.assertEqual(self.factory.calls, [])

    def test_scheme_and_userinfo_rejected(self) -> None:
        resolver = self._resolve(_PUBLIC_V4)
        for url in (
            "file:///etc/passwd", "ftp://example.com/a.png", "data:image/png;base64,AAAA",
            "javascript:alert(1)", "//example.com/a.png", "example.com/a.png",
        ):
            with self.subTest(url=url):
                with self.assertRaisesRegex(RemoteFetchError, "scheme"):
                    fetch_image(url, max_bytes=1000)
        for url in ("http://user:pw@example.com/a.png", "https://user@example.com/a.png",
                    "http://example.com@10.0.0.1/a.png"):
            with self.subTest(url=url):
                with self.assertRaisesRegex(RemoteFetchError, "credentials"):
                    fetch_image(url, max_bytes=1000)
        with self.assertRaisesRegex(RemoteFetchError, "invalid URL"):
            fetch_image("http://example.com:99999/a.png", max_bytes=1000)
        resolver.assert_not_called()

    def test_relative_redirect_is_followed(self) -> None:
        self._resolve(_PUBLIC_V4)
        self.factory = _ConnectionFactory(
            _FakeResponse(302, headers={"Location": "../img/b.png"}),
            _FakeResponse(200, b"OK"),
        )
        with patch.object(safe_fetch, "_open_connection", self.factory):
            self.assertEqual(fetch_image("https://example.com/a/b/c.png", max_bytes=10), b"OK")
        self.assertEqual(self.factory.connections[1].requests[0][1], "/a/img/b.png")

    def test_redirect_to_private_host_rejected(self) -> None:
        resolver = _resolver_by_host({"cdn.example": _PUBLIC_V4, "internal.example": "10.0.0.5"})
        patcher = patch.object(safe_fetch.socket, "getaddrinfo", side_effect=resolver)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.factory = _ConnectionFactory(
            _FakeResponse(301, headers={"Location": "http://internal.example/secret"}),
            _FakeResponse(200, b"SECRET"),
        )
        with patch.object(safe_fetch, "_open_connection", self.factory):
            with self.assertRaisesRegex(RemoteFetchError, "non-public"):
                fetch_image("https://cdn.example/a.png", max_bytes=1000)
        self.assertEqual(len(self.factory.calls), 1)

    def test_redirect_to_other_scheme_rejected(self) -> None:
        self._resolve(_PUBLIC_V4)
        self.factory = _ConnectionFactory(_FakeResponse(302, headers={"Location": "file:///etc/passwd"}))
        with patch.object(safe_fetch, "_open_connection", self.factory):
            with self.assertRaisesRegex(RemoteFetchError, "scheme"):
                fetch_image("https://example.com/a.png", max_bytes=1000)

    def test_redirect_limit(self) -> None:
        self._resolve(_PUBLIC_V4)
        self.factory = _ConnectionFactory(_FakeResponse(307, headers={"Location": "/loop"}))
        with patch.object(safe_fetch, "_open_connection", self.factory):
            with self.assertRaisesRegex(RemoteFetchError, "too many redirects"):
                fetch_image("https://example.com/loop", max_bytes=1000, max_redirects=3)
        self.assertEqual(len(self.factory.calls), 4)
        self.assertTrue(all(connection.closed for connection in self.factory.connections))

    def test_redirect_without_location(self) -> None:
        self._resolve(_PUBLIC_V4)
        self.factory = _ConnectionFactory(_FakeResponse(302))
        with patch.object(safe_fetch, "_open_connection", self.factory):
            with self.assertRaisesRegex(RemoteFetchError, "Location"):
                fetch_image("https://example.com/a.png", max_bytes=1000)

    def test_oversize_content_length_rejected_without_reading(self) -> None:
        self._resolve(_PUBLIC_V4)
        response = _FakeResponse(200, b"x" * 50, headers={"Content-Length": "5000"})
        with patch.object(safe_fetch, "_open_connection", _ConnectionFactory(response)):
            with self.assertRaisesRegex(RemoteFetchError, "larger than 100 bytes"):
                fetch_image("https://example.com/a.png", max_bytes=100)
        self.assertEqual(response.bytes_read, 0)

    def test_oversize_stream_stops_at_limit(self) -> None:
        self._resolve(_PUBLIC_V4)
        response = _FakeResponse(200, b"x" * 10_000)
        with patch.object(safe_fetch, "_open_connection", _ConnectionFactory(response)):
            with self.assertRaisesRegex(RemoteFetchError, "larger than 100 bytes"):
                fetch_image("https://example.com/a.png", max_bytes=100)
        self.assertLessEqual(response.bytes_read, 101)

    def test_body_exactly_at_limit_is_accepted(self) -> None:
        self._resolve(_PUBLIC_V4)
        response = _FakeResponse(200, b"x" * 100, headers={"Content-Length": "100"})
        with patch.object(safe_fetch, "_open_connection", _ConnectionFactory(response)):
            self.assertEqual(len(fetch_image("https://example.com/a.png", max_bytes=100)), 100)

    def test_http_error_status(self) -> None:
        self._resolve(_PUBLIC_V4)
        with patch.object(safe_fetch, "_open_connection", _ConnectionFactory(_FakeResponse(404))):
            with self.assertRaisesRegex(RemoteFetchError, "^HTTP 404$"):
                fetch_image("https://example.com/a.png", max_bytes=100)

    def test_network_errors_are_wrapped(self) -> None:
        self._resolve(_PUBLIC_V4)
        cases = (
            (ConnectionRefusedError(61, "Connection refused"), "connection refused"),
            (socket.timeout("timed out"), "timed out"),
            (ssl.SSLCertVerificationError("bad cert"), "certificate"),
            (ssl.SSLError("handshake"), "TLS error"),
            (http.client.RemoteDisconnected("closed"), "invalid HTTP response"),
            (OSError(65, "No route to host"), "network error"),
        )
        for error, message in cases:
            with self.subTest(error=type(error).__name__):
                with patch.object(safe_fetch, "_open_connection", side_effect=error):
                    with self.assertRaisesRegex(RemoteFetchError, message):
                        fetch_image("https://example.com/a.png", max_bytes=100)

    def test_resolution_failure_is_wrapped(self) -> None:
        with patch.object(safe_fetch.socket, "getaddrinfo", side_effect=socket.gaierror(8, "nodename")):
            with self.assertRaisesRegex(RemoteFetchError, "cannot resolve"):
                fetch_image("https://nope.example/a.png", max_bytes=100)

    def test_total_timeout_is_enforced(self) -> None:
        self._resolve(_PUBLIC_V4)
        with self.assertRaisesRegex(RemoteFetchError, "timed out"):
            fetch_image("https://example.com/a.png", max_bytes=100, total_timeout=0)
        self.assertEqual(self.factory.calls, [])
        now = [0.0]
        response = _FakeResponse(200, b"PNGDATA" * 3)

        def slow_read() -> None:
            now[0] += 20.0

        response.on_read = slow_read
        with patch.object(safe_fetch.time, "monotonic", side_effect=lambda: now[0]), patch.object(
            safe_fetch, "_open_connection", _ConnectionFactory(response)
        ):
            with self.assertRaisesRegex(RemoteFetchError, "timed out"):
                fetch_image("https://example.com/a.png", max_bytes=100, total_timeout=30)
        self.assertEqual(response.bytes_read, 14)  # two chunks, then the budget ran out

    def test_falls_back_to_next_vetted_address(self) -> None:
        self._resolve(_PUBLIC_V6, _PUBLIC_V4)
        factory = _ConnectionFactory(
            _FakeResponse(200, b"OK"), connect_errors={_PUBLIC_V6: OSError(65, "No route to host")}
        )
        with patch.object(safe_fetch, "_open_connection", factory):
            self.assertEqual(fetch_image("https://example.com/a.png", max_bytes=10), b"OK")
        self.assertEqual([call[1] for call in factory.calls], [_PUBLIC_V6, _PUBLIC_V4])
        self.assertTrue(factory.connections[0].closed)

    def test_all_addresses_failing_reports_the_last_error(self) -> None:
        self._resolve(_PUBLIC_V4, "8.8.8.8")
        refused = ConnectionRefusedError(61, "Connection refused")
        factory = _ConnectionFactory(
            _FakeResponse(200, b"OK"), connect_errors={_PUBLIC_V4: refused, "8.8.8.8": refused}
        )
        with patch.object(safe_fetch, "_open_connection", factory):
            with self.assertRaisesRegex(RemoteFetchError, "connection refused"):
                fetch_image("https://example.com/a.png", max_bytes=10)
        self.assertEqual(len(factory.calls), 2)

    def test_tls_failure_does_not_try_other_addresses(self) -> None:
        self._resolve(_PUBLIC_V4, "8.8.8.8")
        factory = _ConnectionFactory(
            _FakeResponse(200, b"OK"),
            connect_errors={_PUBLIC_V4: ssl.SSLCertVerificationError("bad cert")},
        )
        with patch.object(safe_fetch, "_open_connection", factory):
            with self.assertRaisesRegex(RemoteFetchError, "certificate"):
                fetch_image("https://example.com/a.png", max_bytes=10)
        self.assertEqual(len(factory.calls), 1)

    def test_invalid_arguments(self) -> None:
        with self.assertRaises(ValueError):
            fetch_image("https://example.com/a.png", max_bytes=-1)
        with self.assertRaises(ValueError):
            fetch_image("https://example.com/a.png", max_bytes=1, timeout=0)


class PinnedConnectionTests(unittest.TestCase):
    def test_https_connects_to_ip_but_verifies_original_host(self) -> None:
        context = MagicMock(spec=ssl.SSLContext)
        raw_socket = MagicMock()
        budgets = iter([5.0, 4.0])
        with patch.object(socket, "create_connection", return_value=raw_socket) as create:
            connection = safe_fetch._PinnedHTTPSConnection(
                _PUBLIC_V4, 443, server_hostname="example.com",
                budget=lambda: next(budgets), context=context,
            )
            connection.connect()
        self.assertEqual(create.call_args[0][:2], ((_PUBLIC_V4, 443), 5.0))
        raw_socket.settimeout.assert_called_with(4.0)  # handshake bounded by the budget
        context.wrap_socket.assert_called_once_with(raw_socket, server_hostname="example.com")
        self.assertIs(connection.sock, context.wrap_socket.return_value)

    def test_open_connection_uses_verifying_tls_context(self) -> None:
        connection = safe_fetch._open_connection("https", _PUBLIC_V4, 443, "example.com", lambda: 5.0)
        self.assertIsInstance(connection, safe_fetch._PinnedHTTPSConnection)
        context = connection._tls_context
        self.assertEqual(context.verify_mode, ssl.CERT_REQUIRED)
        self.assertTrue(context.check_hostname)
        self.assertIs(connection.response_class.func, safe_fetch._DeadlineResponse)
        plain = safe_fetch._open_connection("http", _PUBLIC_V4, 80, "example.com", lambda: 5.0)
        self.assertEqual((plain.host, plain.port, plain.timeout), (_PUBLIC_V4, 80, 5.0))
        self.assertIs(plain.response_class.func, safe_fetch._DeadlineResponse)


class _ImageHandler(http.server.BaseHTTPRequestHandler):
    hosts_seen: list[str] = []

    def do_GET(self) -> None:  # noqa: N802 - http.server API
        type(self).hosts_seen.append(self.headers.get("Host", ""))
        if self.path == "/redirect":
            self.send_response(302)
            self.send_header("Location", "/image.png")
            self.end_headers()
            return
        body = b"\x89PNG-fake" if self.path == "/image.png" else b"y" * 5000
        self.send_response(200)
        if self.path == "/image.png":
            self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args) -> None:
        pass


class LocalServerTests(unittest.TestCase):
    """Real sockets on 127.0.0.1 -- no internet needed."""

    def setUp(self) -> None:
        _ImageHandler.hosts_seen = []
        self.server = http.server.HTTPServer(("127.0.0.1", 0), _ImageHandler)
        self.port = self.server.server_address[1]
        thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(self.server.server_close)
        self.addCleanup(self.server.shutdown)

    def test_loopback_server_is_refused(self) -> None:
        with self.assertRaisesRegex(RemoteFetchError, "non-public"):
            fetch_image(f"http://127.0.0.1:{self.port}/image.png", max_bytes=1000)
        with self.assertRaisesRegex(RemoteFetchError, "non-public"):
            fetch_image(f"http://localhost:{self.port}/image.png", max_bytes=1000)
        self.assertEqual(_ImageHandler.hosts_seen, [])

    def test_real_http_stack_with_policy_relaxed_for_the_test_server(self) -> None:
        real_getaddrinfo = socket.getaddrinfo

        def resolver(host, port, *args, **kwargs):
            if host == "images.example":
                return [(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("127.0.0.1", port))]
            return real_getaddrinfo(host, port, *args, **kwargs)

        with patch.object(safe_fetch.socket, "getaddrinfo", side_effect=resolver), patch.object(
            safe_fetch, "is_public_address", return_value=True
        ):
            data = fetch_image(f"http://images.example:{self.port}/redirect", max_bytes=1000)
            with self.assertRaisesRegex(RemoteFetchError, "larger than 1000 bytes"):
                fetch_image(f"http://images.example:{self.port}/big", max_bytes=1000)
        self.assertEqual(data, b"\x89PNG-fake")
        self.assertEqual(_ImageHandler.hosts_seen[:2], [f"images.example:{self.port}"] * 2)


class _ScriptedServer:
    """Raw TCP server on 127.0.0.1 running ``script(conn, stop_event)`` per client."""

    def __init__(self, script) -> None:
        self._script = script
        self.stop = threading.Event()
        self._listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._listener.bind(("127.0.0.1", 0))
        self._listener.listen(8)
        self._listener.settimeout(0.1)
        self.port = self._listener.getsockname()[1]
        threading.Thread(target=self._accept_loop, daemon=True).start()

    def _accept_loop(self) -> None:
        while not self.stop.is_set():
            try:
                conn, _ = self._listener.accept()
            except OSError:
                continue
            threading.Thread(target=self._serve, args=(conn,), daemon=True).start()

    def _serve(self, conn: socket.socket) -> None:
        try:
            self._script(conn, self.stop)
        except OSError:
            pass
        finally:
            conn.close()

    def close(self) -> None:
        self.stop.set()
        self._listener.close()


def _drip(conn: socket.socket, stop: threading.Event, data: bytes, interval: float) -> None:
    for index in range(len(data)):
        if stop.wait(interval):
            return
        conn.sendall(data[index : index + 1])


def _trickle_status_line(conn, stop) -> None:
    conn.recv(65536)
    _drip(conn, stop, b"HTTP/1.1 200 OK\r\n" * 50, 0.1)


def _trickle_headers(conn, stop) -> None:
    conn.recv(65536)
    conn.sendall(b"HTTP/1.1 200 OK\r\n")
    while not stop.wait(0.1):
        conn.sendall(b"X-Slow: 1\r\n")


def _trickle_body(conn, stop) -> None:
    conn.recv(65536)
    conn.sendall(b"HTTP/1.1 200 OK\r\nConnection: close\r\n\r\n")
    while not stop.wait(0.1):
        conn.sendall(b"x")


def _silent(conn, stop) -> None:
    stop.wait(30)


def _huge_headers(conn, stop) -> None:
    conn.recv(65536)
    conn.sendall(b"HTTP/1.1 200 OK\r\n" + (b"X-Big: " + b"a" * 8000 + b"\r\n") * 10 + b"\r\n")
    stop.wait(5)


def _many_headers(conn, stop) -> None:
    conn.recv(65536)
    conn.sendall(b"HTTP/1.1 200 OK\r\n" + b"X-A: b\r\n" * 150 + b"\r\n")
    stop.wait(5)


class DeadlineTests(unittest.TestCase):
    """Hostile local servers must not hold a fetch beyond ``total_timeout``."""

    def _fetch(self, script, scheme: str = "http", **kwargs) -> tuple[RemoteFetchError, float]:
        server = _ScriptedServer(script)
        self.addCleanup(server.close)
        real_getaddrinfo = socket.getaddrinfo

        def resolver(host, port, *args, **kw):
            if host == "slow.example":
                return [(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("127.0.0.1", port))]
            return real_getaddrinfo(host, port, *args, **kw)

        started = time.monotonic()
        with patch.object(safe_fetch.socket, "getaddrinfo", side_effect=resolver), patch.object(
            safe_fetch, "is_public_address", return_value=True
        ):
            with self.assertRaises(RemoteFetchError) as caught:
                fetch_image(f"{scheme}://slow.example:{server.port}/a.png", max_bytes=10_000, **kwargs)
        return caught.exception, time.monotonic() - started

    def assert_timed_out_within(self, script, total: float, scheme: str = "http") -> None:
        error, elapsed = self._fetch(script, scheme, timeout=1.0, total_timeout=total)
        self.assertEqual(str(error), "timed out")
        self.assertLess(elapsed, total + 1.0)

    def test_trickled_status_line(self) -> None:
        self.assert_timed_out_within(_trickle_status_line, 1.5)

    def test_trickled_headers(self) -> None:
        self.assert_timed_out_within(_trickle_headers, 1.5)

    def test_trickled_body(self) -> None:
        self.assert_timed_out_within(_trickle_body, 1.5)

    def test_silent_server(self) -> None:
        error, elapsed = self._fetch(_silent, timeout=5.0, total_timeout=0.8)
        self.assertEqual(str(error), "timed out")
        self.assertLess(elapsed, 1.8)

    def test_stalled_tls_handshake(self) -> None:
        error, elapsed = self._fetch(_silent, scheme="https", timeout=5.0, total_timeout=0.8)
        self.assertEqual(str(error), "timed out")
        self.assertLess(elapsed, 1.8)

    def test_header_size_and_count_are_capped(self) -> None:
        error, _ = self._fetch(_huge_headers, timeout=2.0, total_timeout=3.0)
        self.assertEqual(str(error), "response headers are too large")
        error, _ = self._fetch(_many_headers, timeout=2.0, total_timeout=3.0)
        self.assertEqual(str(error), "invalid HTTP response")


if __name__ == "__main__":
    unittest.main()
