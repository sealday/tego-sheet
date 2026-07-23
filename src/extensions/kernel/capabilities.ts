import type { JsonValue } from '../../core/types/json';
import type { Diagnostic } from '../../document/diagnostics';
import type { FormulaFunctionDefinition } from '../../formula/function-registry';
import type { ResourceResolver } from '../../template/resources';
import type { ExtensionManifest, KernelEnvironment } from './manifest';

/** Scalar value exposed to the formula evaluator by a custom cell type. */
export type CellTypeScalar = null | string | number | boolean;

/** Presentation-independent semantics owned by a cell type definition. */
export interface CellTypeSemantics {
  readonly formattedText: string;
  readonly accessibilityLabel: string;
  readonly role: 'text' | 'checkbox' | 'combobox';
  readonly checked?: boolean;
}

/** Environment inputs that may affect localized cell text. */
export interface CellTypeDescriptionEnvironment {
  readonly locale: string;
  readonly timeZone: string;
}

/** Internal contract implemented by built-in custom cell types. */
export interface BuiltInCellTypeDefinition<Value extends JsonValue> {
  readonly id: string;
  readonly schemaVersion: number;
  validate(value: JsonValue): value is Value;
  migrate?(value: JsonValue, fromVersion: number): Value;
  describe(value: Value, environment: CellTypeDescriptionEnvironment): CellTypeSemantics;
  toFormulaScalar(value: Value): CellTypeScalar;
}

/**
 * Open capability map. Internal feature modules extend it with declaration merging.
 *
 * @internal
 */
export interface KernelCapabilities {
  'cell-type': BuiltInCellTypeDefinition<JsonValue>;
  /** One explicitly registered formula function definition. */
  'formula-function': FormulaFunctionDefinition;
  /** Explicit host-owned template resource resolver. */
  'resource-resolver': ResourceResolver;
}

/** Registered capability discriminator. */
export type KernelExtensionKind = keyof KernelCapabilities & string;

/** Restricted initialization context for trusted package extensions. */
export interface KernelContext {
  readonly environment: KernelEnvironment;
  readonly signal: AbortSignal;
  readonly diagnostics: (diagnostic: Diagnostic) => void;
}

/** One typed implementation and its lifecycle hooks. */
export type KernelRegistration<K extends KernelExtensionKind = KernelExtensionKind> = {
  [CapabilityKind in K]: {
    readonly manifest: ExtensionManifest & { readonly kind: CapabilityKind };
    readonly implementation: KernelCapabilities[CapabilityKind];
    readonly initialize?: (context: KernelContext) => void | Promise<void>;
    readonly dispose?: () => void | Promise<void>;
  };
}[K];
