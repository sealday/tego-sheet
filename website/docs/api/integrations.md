---
sidebar_label: Host integrations
---

# Host integration protocols

Import host-owned coordination protocols from `tego-sheet/integrations`. The package supplies
safe state machines and validation boundaries, not storage, collaboration, comment, AI, or
identity services.

```ts
import {
  createPermissionSnapshot,
  createPermissionStore,
} from 'tego-sheet/integrations';

const permissions = createPermissionStore();
permissions.replace(
  createPermissionSnapshot({
    revision: 'permission-42',
    actorId: 'user-7',
    grants: [
      {
        action: 'range:edit',
        target: {
          type: 'range',
          range: {
            sheetId: 'sheet-1',
            start: { row: 0, column: 0 },
            end: { row: 99, column: 9 },
          },
        },
      },
    ],
  }),
);

const allowed = permissions.can('range:edit', {
  type: 'range',
  range: {
    sheetId: 'sheet-1',
    start: { row: 2, column: 1 },
    end: { row: 2, column: 1 },
  },
});
```

Permissions default deny and multi-target checks are all-or-none. Persistence batches retain one
idempotency key across retry. Comment anchors transform with structural edits and use a durable
outbox. Remote transactions are revision-checked, deduplicated, bounded, and applied atomically;
presence is ephemeral and sanitized.

AI integration projects only explicitly selected context, applies redaction and byte/cell limits,
validates commands against a whitelist, performs a dry run, and requires explicit acceptance
against the same revision and permission snapshot. API keys, prompts, cell values, model output,
and command payloads must not enter diagnostics or telemetry.

`createControllerAiProposalSession` binds that protocol to the public document controller. It
shows a value-free `contextSummary`, previews the exact transaction, and commits it only when the
user explicitly calls `accept()`. `AiProposalPanel` is the optional React review surface:

```tsx
import { AiProposalPanel } from 'tego-sheet';
import { createControllerAiProposalSession } from 'tego-sheet/integrations';

const session = await createControllerAiProposalSession({
  controller,
  permissions,
  adapter: hostOwnedModelAdapter,
  signal,
  transactionId: 'ai-change-42',
  request: {
    instruction: 'Normalize the selected amounts',
    locale: 'en-US',
    allowedCommandTypes: ['set-cell-input'],
  },
  context: {
    ranges: [{ sheetId: 'sheet-1', start: { row: 1, column: 1 }, end: { row: 20, column: 1 } }],
    include: ['values'],
    redactions: [{ kind: 'mask-strings', replacement: '[redacted]' }],
  },
});

root.render(<AiProposalPanel session={session} />);
```
