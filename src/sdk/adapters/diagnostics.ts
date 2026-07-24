import type { Diagnostic } from '../../document/diagnostics';
import type { ExtensionKernelError } from '../../extensions/kernel/manifest';
import type {
  AdapterDiagnostic,
  AdapterDiagnosticCode,
  AdapterDiagnosticDetails,
  AdapterManifest,
} from './types';

/** Public adapter failure carrying one stable diagnostic. */
export class AdapterSdkError extends Error {
  readonly code: AdapterDiagnosticCode;
  readonly diagnostic: AdapterDiagnostic;
  readonly diagnostics: readonly AdapterDiagnostic[];

  constructor(diagnostics: readonly AdapterDiagnostic[]) {
    const primary = diagnostics[0];
    if (primary === undefined) throw new Error('AdapterSdkError requires a diagnostic');
    super(primary.message, { cause: primary.cause });
    this.name = 'AdapterSdkError';
    this.code = primary.code;
    this.diagnostic = primary;
    this.diagnostics = Object.freeze([...diagnostics]);
  }
}

export function adapterDiagnostic(
  code: AdapterDiagnosticCode,
  stage: Diagnostic['stage'],
  message: string,
  options: {
    readonly manifest?: Pick<AdapterManifest, 'id'>;
    readonly details?: AdapterDiagnosticDetails;
    readonly cause?: unknown;
  } = {},
): AdapterDiagnostic {
  return Object.freeze({
    code,
    severity: 'error',
    domain: 'extension',
    stage,
    message,
    ...(options.manifest === undefined
      ? {}
      : { location: Object.freeze({ adapterId: options.manifest.id }) }),
    ...(options.details === undefined ? {} : { details: options.details }),
    ...(options.cause === undefined ? {} : { cause: options.cause }),
  });
}

const kernelCodeMap = {
  EXTENSION_MANIFEST_INVALID: 'ADAPTER_MANIFEST_INVALID',
  EXTENSION_API_INCOMPATIBLE: 'ADAPTER_VERSION_UNSUPPORTED',
  EXTENSION_DUPLICATE_ID: 'ADAPTER_DUPLICATE_ID',
  EXTENSION_NOT_FOUND: 'ADAPTER_NOT_FOUND',
  EXTENSION_AMBIGUOUS: 'ADAPTER_AMBIGUOUS',
  EXTENSION_ENVIRONMENT_UNSUPPORTED: 'ADAPTER_ENVIRONMENT_UNSUPPORTED',
  EXTENSION_INITIALIZE_FAILED: 'ADAPTER_INITIALIZATION_FAILED',
  EXTENSION_REGISTRY_DISPOSED: 'ADAPTER_REGISTRY_DISPOSED',
  EXTENSION_DISPOSE_FAILED: 'ADAPTER_DISPOSE_FAILED',
  CELL_TYPE_VALUE_INVALID: 'ADAPTER_INVOCATION_FAILED',
} as const satisfies Readonly<Record<string, AdapterDiagnosticCode>>;

export function mapKernelError(error: ExtensionKernelError): AdapterSdkError {
  return new AdapterSdkError(
    error.diagnostics.map((diagnostic) =>
      adapterDiagnostic(
        kernelCodeMap[diagnostic.code] ?? 'ADAPTER_INVOCATION_FAILED',
        diagnostic.stage,
        diagnostic.message.replaceAll('Extension ', 'Adapter '),
        {
          ...(diagnostic.location?.adapterId === undefined
            ? {}
            : { manifest: { id: diagnostic.location.adapterId } }),
          ...(diagnostic.details === undefined
            ? {}
            : { details: diagnostic.details as AdapterDiagnosticDetails }),
          ...(diagnostic.cause === undefined ? {} : { cause: diagnostic.cause }),
        },
      ),
    ),
  );
}
