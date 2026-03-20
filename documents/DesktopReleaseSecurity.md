# Desktop Release Security Checklist

This checklist is for Windows `.exe` releases of **Customs Billing Portal**.

## 1. Build Artifacts

1. Run a fresh build:
   - `npm run desktop:dist`
2. Verify generated files:
   - `release/Customs Billing Portal Setup 1.0.0.exe` (installer)
   - `release/win-unpacked/Customs Billing Portal.exe` (app exe)

## 2. Integrity and Corruption Checks

1. Run verification:
   - `npm run desktop:verify`
2. Confirm output includes:
   - SHA256 for installer and app exe
   - File sizes (must not be tiny; script fails if `< 5 MB`)
   - Authenticode status

## 3. Production Signing (Required for Public Distribution)

Unsigned EXEs trigger SmartScreen warnings and have weaker trust.

1. Configure signing certificate environment variables before build:
   - `CSC_LINK` (path/URL to `.pfx`)
   - `CSC_KEY_PASSWORD` (certificate password)
2. Build with signing enforcement:
   - `npm run desktop:dist:signed`
3. Verify signature is valid:
   - `npm run desktop:verify:signed`

If `desktop:verify:signed` fails, do not distribute the installer.

## 4. Runtime Hardening (Already Implemented)

The desktop app is configured to:

1. Keep `contextIsolation: true` and `nodeIntegration: false`
2. Enable renderer sandboxing
3. Deny permission prompts by default
4. Block webview attach
5. Prevent navigation/redirect to unknown pages
6. Restrict remote requests to allowlisted origins only
7. Disable devtools in packaged mode

## 5. Final Pre-Release Gate

Before sharing installer outside your team:

1. `npm run desktop:release:check`
2. Manual smoke test:
   - Open app
   - Run Candata conversion with known sample files
   - Confirm both outputs download and open
3. Attach SHA256 hash in release notes for installer verification
