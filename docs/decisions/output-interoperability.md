# Output interoperability matrix

Date: 2026-07-23

The durable probe is:

```bash
npm run test:output-interop
```

It creates an isolated temporary XLSX, verifies typed strings, formulas,
worksheet visibility, exact print area, custom paper geometry, and conditional
formatting, round-trips it, then deletes the temporary directory. It never
opens or modifies a user workbook.

The script reports the implementation gate separately from the external
release gate. Exit status `2` means the generated package passed the local
implementation probes while a required external consumer remains pending; it
must not be interpreted as a fully green release gate. Excel Desktop automation
is opt-in with `TEGO_SHEET_EXCEL_INTEROP=1`.

| Consumer         | Status       | Evidence                                                                                                                                                              |
| ---------------- | ------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| openpyxl 3.1.5   | Passed       | Opened and verified the generated source and the LibreOffice round-trip. No Excel Desktop round-trip claim is made.                                                   |
| LibreOffice 24.8 | Passed       | Headless open/save completed and the round-trip passed the openpyxl probe. It normalized `veryHidden` to `hidden` and converted 4in × 2in to equivalent 102mm × 51mm. |
| Excel Desktop    | Blocked      | The opt-in AppleScript probe hit its 60-second hard timeout (`ETIMEDOUT`, `SIGTERM`) with no stdout/stderr, consistent with an Automation or Excel GUI prompt.        |
| Excel Web        | Not executed | No authorized Excel Web account/upload channel is available locally.                                                                                                  |

The timed-out run removed its isolated temporary directory and the
probe-launched Excel process was closed; no user workbook was opened or
modified. Excel Desktop and Excel Web remain explicit external release checks.
This repository does not claim either passed.
