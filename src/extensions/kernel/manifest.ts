import type { JsonValue } from '../../core/types/json';
import type { Diagnostic } from '../../document/diagnostics';
import type { KernelExtensionKind } from './capabilities';

/** Kernel API compatibility version. */
export type ApiVersion = `${number}.${number}`;

/** Runtime environments understood by the internal adapter kernel. */
export type KernelEnvironment = 'browser' | 'worker' | 'node';

/** Minimal internal extension registration manifest. */
export interface ExtensionManifest {
  readonly id: string;
  readonly apiVersion: ApiVersion;
  readonly kind: KernelExtensionKind;
  readonly environments: readonly KernelEnvironment[];
}

export type ExtensionDiagnosticCode =
  | 'EXTENSION_MANIFEST_INVALID'
  | 'EXTENSION_API_INCOMPATIBLE'
  | 'EXTENSION_DUPLICATE_ID'
  | 'EXTENSION_NOT_FOUND'
  | 'EXTENSION_AMBIGUOUS'
  | 'EXTENSION_ENVIRONMENT_UNSUPPORTED'
  | 'EXTENSION_INITIALIZE_FAILED'
  | 'EXTENSION_REGISTRY_DISPOSED'
  | 'CELL_TYPE_VALUE_INVALID'
  | 'EXTENSION_DISPOSE_FAILED';

/** Diagnostic emitted by extension validation, resolution, execution, or cleanup. */
export interface ExtensionDiagnostic extends Diagnostic {
  readonly code: ExtensionDiagnosticCode;
}

interface ParsedVersion {
  readonly major: number;
  readonly minor: number;
}

const environmentSet = new Set<KernelEnvironment>(['browser', 'worker', 'node']);
const idPattern = /^[a-z][a-z0-9]*(?:[.-][a-z0-9]+)*$/;
const kindPattern = /^[a-z][a-z0-9]*(?:-[a-z0-9]+)*$/;
const versionPattern = /^(0|[1-9]\d*)\.(0|[1-9]\d*)$/;

function details(value: Record<string, JsonValue>): JsonValue {
  return value;
}

export function extensionDiagnostic(
  code: ExtensionDiagnosticCode,
  stage: ExtensionDiagnostic['stage'],
  message: string,
  manifest?: Pick<ExtensionManifest, 'id' | 'kind'>,
  diagnosticDetails?: Record<string, JsonValue>,
  cause?: unknown,
): ExtensionDiagnostic {
  return Object.freeze({
    code,
    severity: 'error' as const,
    domain: 'extension' as const,
    stage,
    message,
    ...(manifest === undefined ? {} : { location: Object.freeze({ adapterId: manifest.id }) }),
    ...(diagnosticDetails === undefined ? {} : { details: details(diagnosticDetails) }),
    ...(cause === undefined ? {} : { cause }),
  });
}

export class ExtensionKernelError extends Error {
  readonly code: ExtensionDiagnosticCode;
  readonly diagnostic: ExtensionDiagnostic;
  readonly diagnostics: readonly ExtensionDiagnostic[];

  constructor(diagnostics: readonly ExtensionDiagnostic[]) {
    const primary = diagnostics[0];
    if (primary === undefined) throw new Error('ExtensionKernelError requires a diagnostic');
    super(primary.message, { cause: primary.cause });
    this.name = 'ExtensionKernelError';
    this.code = primary.code;
    this.diagnostic = primary;
    this.diagnostics = Object.freeze([...diagnostics]);
  }
}

export function parseApiVersion(version: string): ParsedVersion | undefined {
  const match = versionPattern.exec(version);
  if (match === null) return undefined;
  return { major: Number(match[1]), minor: Number(match[2]) };
}

export function assertSupportedApiVersion(
  manifest: ExtensionManifest,
  supportedVersion: ApiVersion,
): void {
  const requested = parseApiVersion(manifest.apiVersion);
  const supported = parseApiVersion(supportedVersion);
  if (
    requested === undefined ||
    supported === undefined ||
    requested.major !== supported.major ||
    requested.minor > supported.minor
  ) {
    throw new ExtensionKernelError([
      extensionDiagnostic(
        'EXTENSION_API_INCOMPATIBLE',
        'validate',
        `Extension ${manifest.kind}/${manifest.id} requires API ${manifest.apiVersion}; this package supports ${supportedVersion}`,
        manifest,
        { requested: manifest.apiVersion, supported: supportedVersion },
      ),
    ]);
  }
}

export function validateAndSnapshotManifest(
  manifest: ExtensionManifest,
): Readonly<ExtensionManifest> {
  const candidate = manifest as unknown as Partial<ExtensionManifest>;
  const environments = Array.isArray(candidate.environments) ? [...candidate.environments] : [];
  const valid =
    typeof candidate.id === 'string' &&
    idPattern.test(candidate.id) &&
    typeof candidate.kind === 'string' &&
    kindPattern.test(candidate.kind) &&
    typeof candidate.apiVersion === 'string' &&
    parseApiVersion(candidate.apiVersion) !== undefined &&
    environments.length > 0 &&
    environments.every(
      (environment): environment is KernelEnvironment =>
        typeof environment === 'string' && environmentSet.has(environment as KernelEnvironment),
    ) &&
    new Set(environments).size === environments.length;

  if (!valid) {
    throw new ExtensionKernelError([
      extensionDiagnostic(
        'EXTENSION_MANIFEST_INVALID',
        'validate',
        'Extension manifest must have a stable lowercase ID, kind, API version, and unique supported environments',
        typeof candidate.id === 'string' && typeof candidate.kind === 'string'
          ? { id: candidate.id, kind: candidate.kind as KernelExtensionKind }
          : undefined,
      ),
    ]);
  }

  return Object.freeze({
    id: candidate.id!,
    apiVersion: candidate.apiVersion!,
    kind: candidate.kind!,
    environments: Object.freeze(environments),
  });
}
