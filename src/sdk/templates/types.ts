import type { JsonValue } from '../../core/types/json';
import type { Diagnostic } from '../../document';
import type { PrintDisplayCommand } from '../../print';
import type { AdapterApiVersion, AdapterEnvironment } from '../adapters';

/** Immutable, renderer-neutral source node offered to template modules. */
export interface TemplateModuleNode {
  readonly id: string;
  readonly type: string;
  readonly data: JsonValue;
}

/** Explicit module dependencies and source nodes for one pipeline execution. */
export interface TemplateModuleDocument {
  readonly requiredModules: readonly string[];
  readonly nodes: readonly TemplateModuleNode[];
}

/** Compatibility and deterministic ordering metadata for one template module. */
export interface TemplateModuleManifest {
  readonly id: string;
  readonly apiVersion: AdapterApiVersion;
  readonly environments: readonly AdapterEnvironment[];
  readonly priority: number;
}

/** Restricted context shared by every finite module stage. */
export interface TemplateModuleStageContext {
  readonly moduleId: string;
  readonly nodeId: string;
  readonly signal: AbortSignal;
}

/** Versioned recognize-to-paint template extension protocol. */
export interface TemplateModule {
  readonly manifest: TemplateModuleManifest;
  recognize(
    node: TemplateModuleNode,
    context: TemplateModuleStageContext,
  ): boolean | Promise<boolean>;
  transform(
    node: TemplateModuleNode,
    context: TemplateModuleStageContext,
  ): JsonValue | Promise<JsonValue>;
  resolve(ir: JsonValue, context: TemplateModuleStageContext): JsonValue | Promise<JsonValue>;
  layout(ir: JsonValue, context: TemplateModuleStageContext): JsonValue | Promise<JsonValue>;
  paint(
    ir: JsonValue,
    context: TemplateModuleStageContext,
  ): readonly PrintDisplayCommand[] | Promise<readonly PrintDisplayCommand[]>;
  dispose?: () => void | Promise<void>;
}

export interface TemplateModuleRegistryOptions {
  readonly apiVersion: AdapterApiVersion;
  readonly environment: AdapterEnvironment;
}

/** Isolated registration and lifecycle owner for public template modules. */
export interface TemplateModuleRegistry {
  register(module: TemplateModule): () => Promise<readonly TemplateModuleDiagnostic[]>;
  list(): readonly TemplateModule[];
  dispose(): Promise<readonly TemplateModuleDiagnostic[]>;
}

/** Hard limits applied across one complete module pipeline. */
export interface TemplateModulePipelineLimits {
  readonly maximumNodes: number;
  readonly maximumCommands: number;
  readonly maximumOutputBytes: number;
  readonly maximumMilliseconds: number;
}

export interface ExecuteTemplateModulePipelineOptions {
  readonly document: TemplateModuleDocument;
  readonly signal: AbortSignal;
  readonly limits?: Partial<TemplateModulePipelineLimits>;
  /** Monotonic clock injection for deterministic hosts and tests. */
  readonly clock?: () => number;
}

/** One owned node after all deterministic stages have completed. */
export interface TemplateModuleNodeOutput {
  readonly nodeId: string;
  readonly moduleId: string;
  readonly ir: JsonValue;
  readonly commands: readonly PrintDisplayCommand[];
}

/** Atomic pipeline output; no partial result is returned after any failure. */
export interface TemplateModulePipelineOutput {
  readonly nodes: readonly TemplateModuleNodeOutput[];
  readonly commands: readonly PrintDisplayCommand[];
}

export type TemplateModuleDiagnosticCode =
  | 'TEMPLATE_MODULE_OPTIONS_INVALID'
  | 'TEMPLATE_MODULE_MANIFEST_INVALID'
  | 'TEMPLATE_MODULE_DUPLICATE'
  | 'TEMPLATE_MODULE_VERSION_UNSUPPORTED'
  | 'TEMPLATE_MODULE_ENVIRONMENT_UNSUPPORTED'
  | 'TEMPLATE_MODULE_MISSING'
  | 'TEMPLATE_NODE_INVALID'
  | 'TEMPLATE_NODE_OWNERSHIP_CONFLICT'
  | 'TEMPLATE_NODE_LIMIT_EXCEEDED'
  | 'TEMPLATE_OUTPUT_LIMIT_EXCEEDED'
  | 'TEMPLATE_COMMAND_LIMIT_EXCEEDED'
  | 'TEMPLATE_TIME_LIMIT_EXCEEDED'
  | 'TEMPLATE_PIPELINE_ABORTED'
  | 'TEMPLATE_STAGE_FAILED'
  | 'TEMPLATE_DRAW_COMMAND_INVALID'
  | 'TEMPLATE_MODULE_DISPOSE_FAILED'
  | 'TEMPLATE_REGISTRY_DISPOSED';

export interface TemplateModuleDiagnostic extends Diagnostic {
  readonly code: TemplateModuleDiagnosticCode;
  readonly domain: 'extension';
}
