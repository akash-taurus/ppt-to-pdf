#!/usr/bin/env python3
"""
Convert PowerPoint presentations to PDF.

Conversion backends (tried in order by ``method="auto"``):
  - Microsoft PowerPoint COM automation (Windows, requires ``comtypes``).
  - LibreOffice in headless mode (soffice/libreoffice), cross-platform.

The module can be used as a library::

    from ppt_to_pdf import convert, convert_many

    convert("deck.pptx")                      # -> deck.pdf next to it
    convert("deck.pptx", "out/deck.pdf")      # explicit destination
    convert_many(["a.pptx", "b.pptx"], "out") # batch
"""

from __future__ import annotations

import argparse
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Sequence

#: File types the converters can reasonably handle.
SUPPORTED_SUFFIXES = frozenset(
    {
        ".ppt",
        ".pptx",
        ".pptm",
        ".pps",
        ".ppsx",
        ".ppsm",
        ".pot",
        ".potx",
        ".potm",
        ".odp",
    }
)

DEFAULT_TIMEOUT = 300.0


class ConverterError(RuntimeError):
    """Raised when a conversion backend fails."""


# --------------------------------------------------------------------------
# Backend discovery
# --------------------------------------------------------------------------


def _libreoffice_candidates() -> list[str]:
    candidates: list[str] = []

    env_path = os.environ.get("LIBREOFFICE_PATH")
    if env_path:
        candidates.append(env_path)

    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            candidates.append(path)

    system = platform.system()
    if system == "Windows":
        for env_var in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
            root = os.environ.get(env_var)
            if root:
                candidates.append(str(Path(root) / "LibreOffice" / "program" / "soffice.exe"))
    elif system == "Darwin":
        candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    else:
        candidates.extend(
            ["/usr/bin/soffice", "/usr/local/bin/soffice", "/snap/bin/libreoffice"]
        )

    return candidates


def find_libreoffice() -> str | None:
    """Return the path/name of a LibreOffice executable if available."""
    for candidate in _libreoffice_candidates():
        path = Path(candidate)
        if path.is_file():
            return str(path)
        resolved = shutil.which(candidate)
        if resolved:
            return resolved
    return None


def powerpoint_available() -> bool:
    """Best-effort check for a usable PowerPoint COM backend."""
    if platform.system() != "Windows":
        return False
    try:
        import comtypes.client  # noqa: F401  # type: ignore
    except Exception:
        return False
    return True


def describe_backend(method: str = "auto") -> str:
    """Human readable description of the backend ``method`` will use."""
    if method == "powerpoint":
        return "Microsoft PowerPoint"
    if method == "libreoffice":
        return "LibreOffice"
    if powerpoint_available():
        return "Microsoft PowerPoint"
    if find_libreoffice():
        return "LibreOffice"
    return "no backend found"


# --------------------------------------------------------------------------
# LibreOffice
# --------------------------------------------------------------------------


def convert_with_libreoffice(
    input_path: Path,
    output_path: Path,
    timeout: float = DEFAULT_TIMEOUT,
) -> None:
    soffice = find_libreoffice()
    if not soffice:
        raise ConverterError(
            "LibreOffice not found. Install LibreOffice or put 'soffice' on PATH "
            "(you can also set the LIBREOFFICE_PATH environment variable)."
        )

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="ppt2pdf_") as tmp:
        tmpdir = Path(tmp)
        outdir = tmpdir / "out"
        outdir.mkdir()
        # A private profile lets multiple conversions run at the same time
        # instead of silently reusing (and colliding with) a running instance.
        profile = tmpdir / "profile"

        cmd = [
            soffice,
            f"-env:UserInstallation={profile.as_uri()}",
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--nodefault",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            str(outdir),
            str(input_path),
        ]

        try:
            result = subprocess.run(
                cmd,
                check=False,
                capture_output=True,
                text=True,
                errors="replace",
                timeout=timeout,
            )
        except subprocess.TimeoutExpired as e:
            raise ConverterError(
                f"LibreOffice conversion timed out after {timeout:.0f}s."
            ) from e
        except OSError as e:
            raise ConverterError(f"Could not start LibreOffice: {e}") from e

        produced = outdir / f"{input_path.stem}.pdf"
        if not produced.is_file():
            # Some builds write a differently-named file; take any PDF we got.
            pdfs = sorted(outdir.glob("*.pdf"))
            produced = pdfs[0] if pdfs else produced

        if not produced.is_file():
            log = ((result.stderr or "") + (result.stdout or "")).strip()
            details = log or f"exit code {result.returncode}"
            raise ConverterError(
                f"LibreOffice did not produce a PDF ({details}). "
                "The input file may be corrupt or password protected."
            )

        if output_path.exists():
            output_path.unlink()
        shutil.move(str(produced), str(output_path))


# --------------------------------------------------------------------------
# Microsoft PowerPoint (COM)
# --------------------------------------------------------------------------


def convert_with_powerpoint(
    input_path: Path,
    output_path: Path,
    timeout: float = DEFAULT_TIMEOUT,
) -> None:
    if platform.system() != "Windows":
        raise ConverterError("PowerPoint automation is only available on Windows.")

    try:
        import comtypes  # type: ignore
        import comtypes.client  # type: ignore
    except Exception as e:
        raise ConverterError(
            "PowerPoint conversion requires the 'comtypes' package. "
            "Install it with: pip install comtypes"
        ) from e

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    powerpoint = None
    presentation = None
    com_initialized = False

    try:
        try:
            comtypes.CoInitialize()
            com_initialized = True
        except Exception:
            pass

        try:
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        except Exception as e:
            raise ConverterError(
                "Could not start Microsoft PowerPoint. Is it installed?"
            ) from e

        # Best effort: keep PowerPoint out of the way and quiet.
        try:
            powerpoint.Visible = 0
        except Exception:
            pass
        try:
            powerpoint.DisplayAlerts = 1  # ppAlertsNone
        except Exception:
            pass

        try:
            # Open(FileName, ReadOnly, Untitled, WithWindow)
            presentation = powerpoint.Presentations.Open(
                str(input_path), True, False, False
            )
        except Exception as e:
            raise ConverterError(f"PowerPoint could not open the file: {e}") from e

        try:
            # 32 = ppSaveAsPDF
            presentation.SaveAs(str(output_path), 32)
        except Exception as e:
            raise ConverterError(f"PowerPoint could not export the PDF: {e}") from e
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass
        if com_initialized:
            try:
                comtypes.CoUninitialize()
            except Exception:
                pass


# --------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------


def convert(
    input_path: str | os.PathLike[str],
    output_path: str | os.PathLike[str] | None = None,
    *,
    method: str = "auto",
    timeout: float = DEFAULT_TIMEOUT,
) -> Path:
    """Convert a presentation to PDF and return the PDF path.

    ``method`` is one of ``"auto"``, ``"powerpoint"`` or ``"libreoffice"``.
    Raises :class:`ConverterError` (or a ``ValueError``/``FileNotFoundError``
    for bad input) on failure.
    """
    input_path = Path(input_path).expanduser().resolve()
    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if input_path.suffix.lower() not in SUPPORTED_SUFFIXES:
        supported = ", ".join(sorted(SUPPORTED_SUFFIXES))
        raise ValueError(f"Unsupported input type '{input_path.suffix}'. Supported: {supported}")

    if output_path is None:
        output_path = input_path.with_suffix(".pdf")
    else:
        output_path = Path(output_path).expanduser()
        if output_path.is_dir():
            output_path = output_path / f"{input_path.stem}.pdf"
        elif not output_path.suffix:
            output_path = output_path.with_suffix(".pdf")
        elif output_path.suffix.lower() != ".pdf":
            raise ValueError("Output path must end in .pdf")
        output_path = output_path.resolve()

    if output_path == input_path:
        raise ValueError("Output path must differ from the input path.")

    attempts: list[tuple[str, Callable[[], None]]] = []
    if method == "powerpoint":
        attempts.append(("powerpoint", lambda: convert_with_powerpoint(input_path, output_path, timeout)))
    elif method == "libreoffice":
        attempts.append(("libreoffice", lambda: convert_with_libreoffice(input_path, output_path, timeout)))
    elif method == "auto":
        if platform.system() == "Windows":
            attempts.append(("powerpoint", lambda: convert_with_powerpoint(input_path, output_path, timeout)))
        attempts.append(("libreoffice", lambda: convert_with_libreoffice(input_path, output_path, timeout)))
    else:
        raise ValueError(f"Unknown method: {method}")

    errors: list[str] = []
    for name, attempt in attempts:
        try:
            attempt()
        except Exception as e:  # noqa: BLE001 - report every backend's failure
            errors.append(f"{name}: {e}")
            continue
        if output_path.is_file() and output_path.stat().st_size > 0:
            return output_path
        errors.append(f"{name}: no PDF was produced")

    raise ConverterError(
        "Could not convert the presentation.\n- " + "\n- ".join(errors)
    )


@dataclass
class ConversionResult:
    """Outcome of a single file in :func:`convert_many`."""

    source: Path
    output: Path | None
    ok: bool
    seconds: float = 0.0
    error: str | None = None


def convert_many(
    paths: Iterable[str | os.PathLike[str]],
    output_dir: str | os.PathLike[str] | None = None,
    *,
    method: str = "auto",
    timeout: float = DEFAULT_TIMEOUT,
    on_result: Callable[[ConversionResult], None] | None = None,
) -> list[ConversionResult]:
    """Convert several presentations, optionally into ``output_dir``.

    ``on_result`` is called after each file with its :class:`ConversionResult`.
    Failures are recorded, not raised, so one bad deck does not stop the batch.
    """
    out_dir = Path(output_dir).expanduser().resolve() if output_dir else None
    if out_dir is not None:
        out_dir.mkdir(parents=True, exist_ok=True)

    results: list[ConversionResult] = []
    for raw in paths:
        source = Path(raw).expanduser().resolve()
        destination = None if out_dir is None else out_dir / f"{source.stem}.pdf"
        started = time.monotonic()
        try:
            output = convert(source, destination, method=method, timeout=timeout)
            result = ConversionResult(source, output, True, time.monotonic() - started)
        except Exception as e:  # noqa: BLE001 - collected per file
            result = ConversionResult(
                source, None, False, time.monotonic() - started, str(e)
            )
        results.append(result)
        if on_result is not None:
            on_result(result)
    return results


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert PowerPoint presentations (.ppt, .pptx, ...) to PDF."
    )
    parser.add_argument(
        "input",
        nargs="+",
        help="One or more .ppt/.pptx files (or folders of them)",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Output .pdf path (only valid with a single input file)",
    )
    parser.add_argument(
        "--outdir",
        help="Directory to write PDFs into (default: next to each input)",
    )
    parser.add_argument(
        "-m",
        "--method",
        choices=["auto", "powerpoint", "libreoffice"],
        default="auto",
        help="Conversion backend to use (default: auto)",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=DEFAULT_TIMEOUT,
        help=f"Per-file timeout in seconds (default: {DEFAULT_TIMEOUT:.0f})",
    )
    parser.add_argument("-q", "--quiet", action="store_true", help="Only print errors")
    return parser.parse_args(argv)


def _expand_inputs(raw_inputs: Sequence[str]) -> list[Path]:
    files: list[Path] = []
    seen: set[Path] = set()
    for raw in raw_inputs:
        path = Path(raw).expanduser()
        if path.is_dir():
            candidates = sorted(
                p for p in path.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_SUFFIXES
            )
        else:
            candidates = [path]
        for candidate in candidates:
            resolved = candidate.resolve()
            if resolved not in seen:
                seen.add(resolved)
                files.append(resolved)
    return files


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    files = _expand_inputs(args.input)
    if not files:
        print("Error: no convertible files found.", file=sys.stderr)
        return 1
    if args.output and len(files) > 1:
        print("Error: --output can only be used with a single input file.", file=sys.stderr)
        return 1

    failures = 0
    for index, source in enumerate(files, start=1):
        destination = Path(args.output) if args.output else (
            Path(args.outdir) / f"{source.stem}.pdf" if args.outdir else None
        )
        if not args.quiet:
            prefix = f"[{index}/{len(files)}] " if len(files) > 1 else ""
            print(f"{prefix}Converting {source.name} ...", flush=True)
        try:
            result = convert(
                source, destination, method=args.method, timeout=args.timeout
            )
        except Exception as e:  # noqa: BLE001 - CLI reports and continues
            failures += 1
            print(f"Error: {source.name}: {e}", file=sys.stderr)
            continue
        if not args.quiet:
            print(f"Created: {result}")

    if len(files) > 1 and not args.quiet:
        ok = len(files) - failures
        print(f"Done: {ok}/{len(files)} converted.")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())