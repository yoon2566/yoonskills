#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path


AUTH_ERROR_EXIT = 22


def configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")


def fail(message: str, code: int = 1) -> None:
    print(message, file=sys.stderr)
    raise SystemExit(code)


def resolve_native_gws() -> Path:
    candidates: list[Path] = []

    direct = shutil.which("gws.exe")
    if direct:
        candidates.append(Path(direct))

    for path_entry in os.environ.get("PATH", "").split(os.pathsep):
        if not path_entry:
            continue
        base = Path(path_entry)
        candidates.extend(
            [
                base / "gws.exe",
                base / "node_modules" / "@googleworkspace" / "cli" / "bin" / "gws.exe",
            ]
        )

    candidates.append(
        Path.home()
        / "AppData"
        / "Roaming"
        / "npm"
        / "node_modules"
        / "@googleworkspace"
        / "cli"
        / "bin"
        / "gws.exe"
    )

    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()

    fail("Could not locate the native gws executable. Install Google Workspace CLI and try again.")


def run_gws(executable: Path, args: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [str(executable), *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )


def load_json(path: str) -> str:
    try:
        with Path(path).open("r", encoding="utf-8-sig") as handle:
            value = json.load(handle)
    except OSError as exc:
        fail(f"Could not read JSON file '{path}': {exc}")
    except json.JSONDecodeError as exc:
        fail(f"Invalid JSON file '{path}': {exc}")
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def check_auth(executable: Path) -> None:
    result = run_gws(executable, ["auth", "status"])
    if result.returncode != 0:
        fail(f"gws auth status failed.\n{result.stdout}{result.stderr}", AUTH_ERROR_EXIT)

    try:
        status = json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        fail(f"Could not parse gws auth status output: {exc}", AUTH_ERROR_EXIT)

    if status.get("token_valid") is False:
        fail("gws authentication is not ready: token_valid is false.", AUTH_ERROR_EXIT)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run native gws safely with JSON loaded from UTF-8 files."
    )
    parser.add_argument("--params-file", help="JSON file to pass as --params")
    parser.add_argument("--json-file", help="JSON file to pass as --json")
    parser.add_argument("--skip-auth-check", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("command", nargs=argparse.REMAINDER)
    return parser


def main() -> None:
    configure_stdio()
    args = build_parser().parse_args()

    command = list(args.command)
    if command and command[0] == "--":
        command = command[1:]
    if not command:
        fail("Provide the gws command after --.")
    if args.params_file and "--params" in command:
        fail("Do not combine --params-file with a command-level --params.")
    if args.json_file and "--json" in command:
        fail("Do not combine --json-file with a command-level --json.")

    executable = resolve_native_gws()
    if not args.skip_auth_check:
        check_auth(executable)

    if args.params_file:
        command.extend(["--params", load_json(args.params_file)])
    if args.json_file:
        command.extend(["--json", load_json(args.json_file)])

    if args.dry_run:
        print(
            json.dumps(
                {
                    "executable": str(executable),
                    "command": command,
                    "auth_checked": not args.skip_auth_check,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return

    result = run_gws(executable, command)
    if result.stdout:
        sys.stdout.write(result.stdout)
    if result.stderr:
        sys.stderr.write(result.stderr)
    raise SystemExit(result.returncode)


if __name__ == "__main__":
    main()
