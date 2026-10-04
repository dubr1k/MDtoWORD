"""Батч: файлы по одному в рабочем потоке, прогресс, отчёт с диагностикой."""

from __future__ import annotations

from collections.abc import Callable, Sequence
import logging
from pathlib import Path

import anyio
from mcp.server.fastmcp import Context

from ..converters import ConversionError
from .models import ConversionReport, ConvertedFile, Diagnostic, FailedFile

# Логи — только в stderr: обработчики FastMCP пишут туда, а без них
# logging.lastResort тоже пишет в stderr. stdout занят протоколом.
# Имя логгера — имя пакета, как до разбиения сервера на модули.
logger = logging.getLogger(__package__)


def diagnostics(warnings: Sequence[str]) -> list[Diagnostic]:
    """Перевести предупреждения конвертера в структурированный вид.

    ``ConversionWarning`` несёт ``code`` и ``line``; обычная строка (старый
    код рендерера) получает ``general`` и ``None``.
    """
    result: list[Diagnostic] = []
    for warning in warnings:
        line = getattr(warning, "line", None)
        result.append(
            Diagnostic(
                message=str(warning),
                code=getattr(warning, "code", None) or "general",
                line=line if isinstance(line, int) and line >= 1 else None,
            )
        )
    return result


# (исходник, предупреждения, ошибка): ровно одно из двух последних — None.
Outcome = tuple[Path, list[str] | None, str | None]


async def for_each_file(
    sources: list[Path],
    work: Callable[[Path], list[str]],
    ctx: Context | None,
) -> list[Outcome]:
    """Обработать файлы по одному в рабочем потоке, сообщая о прогрессе.

    Отказ одного файла батч не прерывает. Исключение не из
    ``ConversionError`` — это баг, но и он не должен терять результаты уже
    сконвертированных файлов, поэтому тоже попадает в отчёт, с именем типа
    для диагностики.
    """
    outcomes: list[Outcome] = []
    total = len(sources)
    for done, source in enumerate(sources, start=1):
        warnings: list[str] | None = None
        error_text: str | None = None
        try:
            warnings = await anyio.to_thread.run_sync(work, source)
        except ConversionError as error:
            error_text = str(error)
        except Exception as error:  # noqa: BLE001 — см. докстринг
            error_text = f"{type(error).__name__}: {error}"
        outcomes.append((source, warnings, error_text))
        if ctx is not None:
            status = "failed" if error_text is not None else "done"
            await _report_progress(ctx, done, total, f"{source.name}: {status}")
    return outcomes


async def _report_progress(ctx: Context, done: int, total: int, message: str) -> None:
    """Уведомить о прогрессе, не рискуя отчётом батча.

    Уведомление — вспомогательное: если клиент уже отключился или канал
    закрыт, файлы всё равно записаны, и терять из-за этого отчёт нельзя.
    Отмена (BaseException) сюда не попадает и прерывает батч как положено.
    """
    try:
        await ctx.report_progress(done, total, message=message)
    except Exception:  # noqa: BLE001 — см. докстринг
        logger.warning("Could not send a progress notification", exc_info=True)


def report_conversions(
    sources: list[Path], outputs: dict[Path, Path], outcomes: list[Outcome]
) -> ConversionReport:
    report = ConversionReport(sources_found=len(sources))
    for source, warnings, error_text in outcomes:
        if error_text is not None:
            report.failed.append(FailedFile(source=str(source), error=error_text))
        else:
            report.converted.append(
                ConvertedFile(
                    source=str(source),
                    output=str(outputs[source]),
                    warnings=diagnostics(warnings or []),
                )
            )
    return report
