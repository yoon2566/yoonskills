# Environment And Safety

## Windows PowerShell

- Use PowerShell syntax and Windows paths.
- Do not assume Bash, WSL, or Linux utilities.
- Use `Get-Command <name>` and `<command> --help` to inspect tools.
- If quoting or JSON arguments fail, change the invocation strategy instead of repeating the same command.

## Python

- Use the workspace `.venv`.
- Create it when needed:

```powershell
if (-not (Test-Path -LiteralPath ".\.venv\Scripts\python.exe")) {
    python -m venv .venv
}
```

- Invoke Python explicitly:

```powershell
& ".\.venv\Scripts\python.exe" --version
& ".\.venv\Scripts\python.exe" -m pip install <package>
```

- Do not use global `pip`, `pip install --user`, or an unrelated virtual environment.

## Tool Installation

1. Check whether the tool already exists.
2. Inspect the repository's preferred setup method.
3. Explain meaningful system-wide effects before a global install.
4. Verify the installed command and version.

Use `winget`, npm global installs, or another package manager only when appropriate for the tool and environment. Never present a package ID as universal without checking it.

## File Operations

Before recursive moves, bulk renames, copies, extraction, or overwrite:

1. Resolve absolute source and destination paths.
2. Confirm they remain within the intended location.
3. Preview the file list and collisions.
4. Perform the operation with native PowerShell cmdlets.
5. Verify the result.

## Codex And Plugins

Codex CLI commands are version-sensitive. Run `codex --version` and the relevant `--help` before documenting or executing plugin, MCP, cloud, or experimental commands.

Do not assume `codex plugin list` or `codex plugin install` exists. Use the subcommands exposed by the installed CLI.

## Portability

- Use placeholders such as `<workspace>`, `<skill-dir>`, `<file-id>`, and `<output-path>`.
- Do not publish personal emails, account identifiers, session paths, or machine-specific absolute paths.
- Keep credentials and generated authentication files outside the repository.
