# Copilot Instructions — app-updater

## Project overview

A PowerShell script (`main.ps1`) that reads a CSV file of Intune app names and logo paths, fetches the corresponding apps from Microsoft Intune via the Microsoft Graph API, and bulk-updates their large icons using the Graph `$batch` endpoint.

No external modules are required. The script runs on PowerShell 5.1+.

## How to run

```powershell
# Minimum — uses config.json and applist.csv from the same folder
.\main.ps1

# Custom config / CSV paths (relative to $PSScriptRoot)
.\main.ps1 -configFile myConfig.json -Applist myApps.csv

# Enable verbose diagnostics
.\main.ps1 -Verbose
```

### Required files

| File | Purpose |
|---|---|
| `config.json` | Auth credentials (see below). Gitignored — never commit. |
| `applist.csv` | Two columns: `AppName`, `Logo`. `Logo` is a filename relative to `$PSScriptRoot`. Gitignored. |
| `<logo files>` | Image files referenced in `applist.csv`. PNG preferred; JPG, JPEG, GIF also accepted. |

### `config.json` schema

```json
{
  "tenantId":  "<Azure AD tenant ID>",
  "appId":     "<App registration client ID>",
  "AppSecret": "<Client secret>",
  "thumbprint": "<Certificate thumbprint>"
}
```

- Either `AppSecret` or `thumbprint` is required (both can be provided; certificate is tried first with fallback to secret).
- If `config.json` is absent the script continues with all values `$null`, resulting in 401 errors from Graph.

### `applist.csv` schema

```
AppName,Logo
7-Zip,icons\7zip.png
Google Chrome,icons\chrome.png
```

## Architecture

### Key functions

| Function | Purpose |
|---|---|
| `Invoke-GraphAPI` | Universal Graph wrapper. String `ResourcePath` → single request with auto-pagination. Array → `$batch` (max 20/sub-batch). Returns response object, integer HTTP status on error, or a batch-result hashtable. |
| `Get-GraphAccessToken` | Client-credentials token with in-memory caching (5-min renewal lead time). Supports certificate JWT assertion or client secret. |

### Batch update flow

1. Fetch all Intune apps (`GET deviceAppManagement/mobileApps`).
2. Filter to the 9 supported `@odata.type` values listed in `$appTypes`.
3. Match each CSV row to an app by case-insensitive `displayName`.
4. Base64-encode each logo and build a PATCH body (`microsoft.graph.mimeContent`).
5. Pass parallel `$updatePaths[]` + `$updateBodies[]` arrays to `Invoke-GraphAPI` → triggers native `$batch`.
6. Print per-app success/fail summary.

### `Invoke-GraphAPI` return contract

| Scenario | Return value |
|---|---|
| Single request success | Response `PSCustomObject` (or `$null` for 204 No Content) |
| Single request failure | `[int]` HTTP status code |
| Batch request | `[hashtable]` with keys `value`, `batchProcessed`, `successCount`, `failureCount`, `totalCount` |

Always check `$result -is [int]` before treating a response as a success. Never assume a non-null return means success.

## Conventions

- **No logging to disk** — this is the `-noLogs` variant. All diagnostics use `Write-Verbose`; user-facing output uses `Write-Host` with color (`Green` = success, `Yellow` = warning/skip, `Red` = error).
- **Never throw** inside `Invoke-GraphAPI` — all exceptions are caught and surfaced as integer status codes. Callers must check the return type.
- **`$body` can be a string or a string array** in `Invoke-GraphAPI`. When passing batch PATCH calls supply a parallel array of JSON strings, one per path.
- **API version defaults to `beta`**. Use `v1.0` explicitly when the endpoint is stable and the difference matters.
- **Do not add `ProcessFilterCondition` calls** unless that function is dot-sourced into scope first — it is an undeclared external dependency and will fail at runtime if the filter path is exercised without it.
- **`$logWarn` is not defined** in this file — it is a residual artefact. Do not add new references to it.

## Known gotchas

- A `$null` access token does not cause an immediate hard exit — the script continues and Graph calls silently return 401. Add an explicit null-check after `Get-GraphAccessToken` if you need early termination.
- Single-item arrays are silently collapsed to a scalar inside `Invoke-GraphAPI`, so a one-app CSV update goes through the single-request path, not the batch path. The result-evaluation code handles both cases.
- MIME type defaults to `image/png` for unrecognised extensions — caller is responsible for naming files correctly.
- `config.json` and `applist.csv` are both gitignored. They must be created manually on each machine.

## Security

- `config.json` is gitignored — keep it that way. Never hardcode credentials in the script.
- Certificate-based auth is preferred over client secrets for production use.
- Images are read as bytes and base64-encoded; no shell expansion or command injection is possible.
- The `Invoke-RestMethod` calls use `-UseBasicParsing` and do not invoke Internet Explorer's engine.
