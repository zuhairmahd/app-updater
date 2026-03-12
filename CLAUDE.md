# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A PowerShell script (`main.ps1`) that bulk-updates Intune application logos via the Microsoft Graph `$batch` endpoint. It reads app names and logo paths from a CSV, matches them to apps in Intune, and uploads base64-encoded images.

**Requirements:** PowerShell 5.1+, no external modules.

## How to Run

```powershell
# Default (uses config.json and applist.csv from script folder)
.\main.ps1

# Custom config/CSV paths
.\main.ps1 -configFile myConfig.json -Applist myApps.csv

# Verbose diagnostics
.\main.ps1 -Verbose
```

## Required Files (gitignored, create manually)

**config.json:**
```json
{
  "tenantId": "<Azure AD tenant ID>",
  "appId": "<App registration client ID>",
  "AppSecret": "<Client secret>",
  "thumbprint": "<Certificate thumbprint>"
}
```
Either `AppSecret` or `thumbprint` required (certificate tried first).

**applist.csv:**
```
AppName,Logo
7-Zip,icons\7zip.png
Google Chrome,icons\chrome.png
```

## Architecture

Single-file architecture in `main.ps1` (~1,590 lines).

### Key Functions

| Function | Lines | Purpose |
|----------|-------|---------|
| `Invoke-GraphAPI` | 8-896 | Graph API wrapper; string path = single request with auto-pagination; array = `$batch` (max 20/sub-batch) |
| `Get-GraphAccessToken` | 897-1271 | OAuth token with in-memory caching (5-min renewal buffer); supports certificate JWT or client secret |

### Main Flow (lines 1274-1591)

1. Load `config.json` and validate `applist.csv` exists
2. Acquire access token via `Get-GraphAccessToken`
3. Fetch all Intune apps, filter to 9 supported `@odata.type` values
4. Match CSV app names to Intune apps (case-insensitive)
5. Base64-encode logos and build PATCH bodies
6. Call `Invoke-GraphAPI` with parallel path/body arrays (triggers `$batch`)
7. Display color-coded summary (Green=success, Yellow=warning, Red=error)

### Return Contract for `Invoke-GraphAPI`

| Scenario | Return Value |
|----------|--------------|
| Single request success | `PSCustomObject` (or `$null` for 204) |
| Single request failure | `[int]` HTTP status code |
| Batch request | `[hashtable]` with keys: `value`, `batchProcessed`, `successCount`, `failureCount`, `totalCount` |

**Always check `$result -is [int]` before treating as success.**

## Critical Conventions

- **No logging to disk** — diagnostics via `Write-Verbose`, user output via `Write-Host` with colors
- **Never throw inside helper functions** — return integer status codes instead
- **API version defaults to `beta`** — use `v1.0` explicitly when needed
- **Single-item arrays collapse to scalar** — one-app CSV uses single-request path, not batch
- **`ProcessFilterCondition` is undeclared** — will fail if filter path exercised without dot-sourcing
- **`$logWarn` is undefined** — do not reference it
- **`$null` token doesn't exit** — Graph calls silently return 401; add explicit null-check if early termination needed
- **MIME type defaults to `image/png`** for unrecognized extensions
