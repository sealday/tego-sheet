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
  evaluatePermission,
} from 'tego-sheet/integrations';

const permissions = createPermissionStore();
permissions.replace(
  createPermissionSnapshot({
    revision: 1,
    grants: [{ action: 'cell:write', target: { kind: 'document' }, effect: 'allow' }],
  }),
);

const decision = evaluatePermission(permissions.current(), {
  action: 'cell:write',
  targets: [{ kind: 'document' }],
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
