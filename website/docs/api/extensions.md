---
sidebar_label: Extension SDK
---

# Extension SDK

Import extension contracts from `tego-sheet/sdk`.

```ts
import { createAdapterRegistry, createCapabilityGrant } from 'tego-sheet/sdk';

const registry = createAdapterRegistry({
  apiVersion: '1.0',
  environment: 'browser',
});

const grant = createCapabilityGrant(['workbook:read']);
```

The adapter registry validates version, environment, kind, execution mode, capabilities, and
duplicate IDs before publication. Resolution is deterministic and ambiguity is an error.
`trusted-main` adapters receive only an abort signal, optional document identity, and granted
capabilities. `isolated-worker` adapters expose a transport descriptor, never same-realm
implementation code.

The same entry point provides versioned custom cell plugins and template modules. Unknown or
invalid custom cells fall back to canonical accessible text. Editor sessions settle exactly once.
Registries own initialization, cancellation, unregistration, and asynchronous disposal; hosts
must dispose them when their integration scope ends.
