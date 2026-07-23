import type { JsonValue } from '../../core/types/json';
import type { Diagnostic } from '../../document/diagnostics';
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
export interface KernelRegistration<K extends KernelExtensionKind> {
  readonly manifest: ExtensionManifest & { readonly kind: K };
  readonly implementation: KernelCapabilities[K];
  readonly initialize?: (context: KernelContext) => void | Promise<void>;
  readonly dispose?: () => void | Promise<void>;
}
