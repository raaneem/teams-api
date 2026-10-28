# Security

## Reporting a vulnerability

If you discover a security vulnerability in this project, please open a GitHub issue with the label `security`. Do not include exploit code or sensitive credentials in the report.

---

## Windows Defender false positive

### What triggers it

On Windows, the previous token storage implementation spawned a PowerShell process with an inline `-Command` script that loaded `System.Security.Cryptography.ProtectedData` and called `Protect`/`Unprotect` on binary data:

```
powershell -Command "Add-Type -AssemblyName System.Security; ... ProtectedData::Protect(...) ..."
```

This is one of the most reliable behavioral signatures in Windows Defender's heuristics for ransomware and credential-theft malware — spawning PowerShell, loading a crypto assembly, and encrypting arbitrary binary data inline. Defender flags the _behavior_, not the intent.

### What we do instead

Tokens are now stored using [keytar](https://github.com/atom/node-keytar), which calls the native **Windows Credential Manager** (`wincred`) API directly from C++. No PowerShell is spawned, no inline scripts are executed, no crypto assembly is loaded. Stored credentials appear in **Control Panel → Credential Manager → Windows Credentials** under the service name `teams-api`.

### Token storage by platform

| Platform | Storage                                                              |
| -------- | -------------------------------------------------------------------- |
| macOS    | Keychain (via `security` CLI)                                        |
| Windows  | Windows Credential Manager (via keytar / wincred)                    |
| Linux    | `secret-tool` (libsecret) or `~/.config/teams-api/` with 0o600 perms |
| Fallback | `~/.teams-api/tokens/` (directory 0700, files 0600) on all platforms |

When the platform store refuses a write — for example the locked login Keychain in an SSH session — tokens are saved to the private file cache instead, and `TEAMS_TOKEN_STORE=file` uses only that cache. The cache relies on file permissions, like the browser profile in `~/.teams-api/browser-profile` (Playwright runs Chrome with a mock keychain, so the profile's session cookies are protected by file permissions only). Permissions protect against other users, not against programs running as you; to give a bot access without a shell, see "Locking a bot's SSH key to read-only commands" in the README. Consider excluding `~/.teams-api` from backups (`tmutil addexclusion ~/.teams-api` on macOS).

## Token handling

- Tokens are never printed. Errors, logs, MCP results, and telemetry pass through `src/redact.ts`, which removes every token `teams-api` holds and anything shaped like a credential. `teams-api auth` reports which tokens are cached; only `--print-token` prints them, and read-only mode refuses it.
- `TEAMS_READ_ONLY=true` refuses every request that could change Teams, at the action, client, and HTTP layers.
- Avoid `--token`, `--bearer-token`, and `--substrate-token` on shared machines: command-line arguments are visible to other processes. Prefer the cache or environment variables.

### Migrating from the old DPAPI storage

If you used an older version of `teams-api` on Windows, your tokens were stored as DPAPI-encrypted `.dat` files in `%APPDATA%\teams-api\`. These are no longer read. Run `teams-api auth --login` once to re-authenticate and store tokens in the new location. You can safely delete the old `%APPDATA%\teams-api\` directory.
