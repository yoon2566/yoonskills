# Google Workspace

Use this reference for Google Drive, Sheets, Docs, Slides, Gmail, and Calendar stages.

## Rules

1. Run `gws auth status` before non-trivial work.
2. Stop and report the issue when `token_valid` is false or authentication fails.
3. Run `gws schema <service.resource.method>` before using an unfamiliar route.
4. On Windows PowerShell, use `scripts/run_gws.py` with UTF-8 JSON files for calls that need `--params` or `--json`.
5. Re-read the affected resource after every write.

Do not assume `gws auth status` contains a particular account-name or scope-count field. Inspect the actual output.

## Safe PowerShell Pattern

Use the workspace Python executable:

```powershell
$py = ".\.venv\Scripts\python.exe"
$runner = "<edukit-dir>\scripts\run_gws.py"
```

Create JSON with PowerShell objects rather than hand-escaped strings:

```powershell
$paramsPath = Join-Path $env:TEMP "gws-params.json"
@{
    pageSize = 20
    orderBy = "modifiedTime desc"
    fields = "files(name,id,mimeType,modifiedTime,size)"
} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $paramsPath -Encoding UTF8

& $py $runner --params-file $paramsPath -- drive files list
```

For a request body:

```powershell
$bodyPath = Join-Path $env:TEMP "gws-body.json"
@{
    properties = @{ title = "수업 자료" }
} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $bodyPath -Encoding UTF8

& $py $runner --json-file $bodyPath -- sheets spreadsheets create
```

## Verified Route Shapes

Inspect them again with `gws schema` because CLI discovery surfaces can change.

```text
drive.files.list
drive.files.create
drive.files.get
drive.files.export
drive.files.update
sheets.spreadsheets.create
sheets.spreadsheets.get
sheets.spreadsheets.values.get
sheets.spreadsheets.values.update
sheets.spreadsheets.values.append
slides.presentations.create
slides.presentations.get
gmail.users.messages.list
gmail.users.messages.get
gmail.users.messages.send
calendar.events.list
calendar.events.insert
```

There is no `slides.presentations.pages.list` route in the checked CLI surface. Retrieve a presentation with `slides.presentations.get` and inspect its `slides` field.

## Common Operations

### Drive Search

Put queries such as `'<folder-id>' in parents`, `name contains '수업'`, or MIME-type filters in the params JSON file. Do not inline them in PowerShell.

### Sheets Append

Use `sheets spreadsheets values append`, not `sheets values append`.

### Download And Export

- Use `drive files get` with `alt=media` for binary files stored in Drive.
- Use `drive files export` for native Google Docs, Sheets, and Slides.
- Specify the desired export MIME type and output path.

### Writes

For create, update, append, send, trash, or upload operations:

1. Confirm the target ID or destination.
2. Use file-backed JSON.
3. Run the command.
4. Re-read or re-list the target.
