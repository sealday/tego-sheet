# Migration to Workbook 2.0

Workbook 2.0 replaces mutable sheet JSON with a versioned `SpreadsheetDocument`. Perform the
conversion once at the storage boundary, retain stable IDs, and let every editor, command,
template, and output adapter consume the same snapshot.

## Convert legacy storage

```ts
import { migrateLegacyWorkbook, serializeSpreadsheetDocument } from 'tego-sheet';

const migrated = migrateLegacyWorkbook(storedWorkbook);
if (!migrated.ok) {
  throw new Error(
    migrated.diagnostics.map(({ code, message }) => `${code}: ${message}`).join('\n'),
  );
}

await saveJson(serializeSpreadsheetDocument(migrated.document));
```

`migrateLegacyWorkbook` accepts the legacy single-sheet object or ordered sheet array. It returns
the document only after full validation. Unsupported fields are explicit
`LEGACY_FIELD_DEGRADED` or `LEGACY_FIELD_DROPPED` diagnostics; cached formula results are never
promoted to document truth.

Use deterministic ID factories when fixtures, signatures, or idempotent batch migrations require
byte-stable output. Never regenerate `document.id` or `workbook.sheets[].id` during ordinary saves.

## Replace mutations with commands

Do not mutate cells, rows, styles, or sheet arrays. UI edits, imperative handle calls, undo/redo,
and host integrations all cross the validated command/transaction boundary. A rejected command
leaves the prior snapshot and revision unchanged.

Use `defaultDocument` for an uncontrolled editor. Use `document` plus `onDocumentChange` for a
controlled editor, and keep the current object reference across unrelated parent renders. Do not
switch ownership modes after mount.

## Persistence rollout

1. Read legacy JSON and migrate it in memory.
2. Store the schema 2 snapshot beside the legacy record with a format/version marker.
3. Read it back and validate it before changing the active pointer.
4. Retain the old record for rollback until application-level verification completes.
5. Persist later changes from `onDocumentChange`; serialize the supplied snapshot, not UI state.

## Compatibility checklist

- Stable document, sheet, table, object, and column IDs survive renames and structural edits.
- Unknown valid extension data survives unrelated edits and serialization.
- Formula source is persisted; cached values are recalculated.
- Controlled replacements are atomic and revision-aware.
- Print, PDF, image, interchange, and SDK adapters receive immutable snapshots.
- Migration diagnostics are recorded before the legacy copy is retired.

Workbook 2.0 has no public mutable controller or renderer internals. Import only documented
package entry points; internal paths are intentionally unavailable.
