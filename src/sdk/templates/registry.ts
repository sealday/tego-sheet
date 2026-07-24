import type { JsonValue } from '../../core/types/json';
import type { PrintDisplayCommand } from '../../print';
import { jsonSnapshotBytes, snapshotJsonValue } from '../adapters/json-safe';
import type {
  ExecuteTemplateModulePipelineOptions,
  TemplateModule,
  TemplateModuleDiagnostic,
  TemplateModuleDiagnosticCode,
  TemplateModuleManifest,
  TemplateModuleNode,
  TemplateModuleNodeOutput,
  TemplateModulePipelineLimits,
  TemplateModulePipelineOutput,
  TemplateModuleRegistry,
  TemplateModuleRegistryOptions,
  TemplateModuleStageContext,
} from './types';

interface RegistryState {
  readonly apiVersion: string;
  readonly environment: TemplateModuleRegistryOptions['environment'];
  readonly records: Map<string, TemplateModule>;
  readonly idleWaiters: Set<() => void>;
  activeExecutions: number;
  disposed: boolean;
  disposePromise?: Promise<readonly TemplateModuleDiagnostic[]>;
}

const states = new WeakMap<TemplateModuleRegistry, RegistryState>();
const idPattern = /^[a-z][a-z0-9]*(?:[.-][a-z0-9]+)*$/;
const typePattern = /^[a-z][a-z0-9]*(?:[._-][a-z0-9]+)*$/;
const apiVersionPattern = /^(0|[1-9]\d*)\.(0|[1-9]\d*)$/;
const environments = new Set(['browser', 'worker', 'node']);
const defaults: Readonly<TemplateModulePipelineLimits> = Object.freeze({
  maximumNodes: 10_000,
  maximumCommands: 100_000,
  maximumOutputBytes: 16 * 1024 * 1024,
  maximumMilliseconds: 5_000,
});

function diagnostic(
  code: TemplateModuleDiagnosticCode,
  stage: TemplateModuleDiagnostic['stage'],
  message: string,
  details?: JsonValue,
  cause?: unknown,
): TemplateModuleDiagnostic {
  return Object.freeze({
    code,
    severity: 'error',
    domain: 'extension',
    stage,
    message,
    ...(details === undefined ? {} : { details }),
    ...(cause === undefined ? {} : { cause }),
  });
}

/** Error carrying stable diagnostics for template module boundary failures. */
export class TemplateModuleSdkError extends Error {
  readonly diagnostics: readonly TemplateModuleDiagnostic[];

  constructor(diagnostics: readonly TemplateModuleDiagnostic[]) {
    super(diagnostics.map(({ message }) => message).join('; '));
    this.name = 'TemplateModuleSdkError';
    this.diagnostics = Object.freeze([...diagnostics]);
  }
}

function fail(
  code: TemplateModuleDiagnosticCode,
  stage: TemplateModuleDiagnostic['stage'],
  message: string,
  details?: JsonValue,
  cause?: unknown,
): never {
  throw new TemplateModuleSdkError([diagnostic(code, stage, message, details, cause)]);
}

function major(version: string): string {
  return version.split('.')[0] as string;
}

function compareAscii(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

function compareModules(left: TemplateModule, right: TemplateModule): number {
  return (
    right.manifest.priority - left.manifest.priority ||
    compareAscii(left.manifest.id, right.manifest.id)
  );
}

function dataProperty(object: object, key: PropertyKey, required = true): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(object, key);
  if (descriptor === undefined) {
    if (!required) return undefined;
    throw new TypeError(`${String(key)} is required`);
  }
  if (!('value' in descriptor)) throw new TypeError(`${String(key)} must be a data property`);
  return descriptor.value;
}

function snapshotManifest(value: unknown): TemplateModuleManifest {
  const snapshot = snapshotJsonValue(value, 'templateModule.manifest');
  if (snapshot === null || Array.isArray(snapshot) || typeof snapshot !== 'object') {
    throw new TypeError('Template module manifest must be an object');
  }
  const record = snapshot as Readonly<Record<string, JsonValue>>;
  const { id, apiVersion, environments: supported, priority } = record;
  if (
    typeof id !== 'string' ||
    !idPattern.test(id) ||
    typeof apiVersion !== 'string' ||
    !apiVersionPattern.test(apiVersion) ||
    !Array.isArray(supported) ||
    supported.length === 0 ||
    new Set(supported).size !== supported.length ||
    supported.some(
      (environment) => typeof environment !== 'string' || !environments.has(environment),
    ) ||
    typeof priority !== 'number' ||
    !Number.isSafeInteger(priority) ||
    Object.keys(record).some(
      (key) => !['id', 'apiVersion', 'environments', 'priority'].includes(key),
    )
  ) {
    throw new TypeError('Template module manifest is invalid');
  }
  return snapshot as unknown as TemplateModuleManifest;
}

function snapshotModule(value: TemplateModule): TemplateModule {
  if (value === null || typeof value !== 'object') {
    throw new TypeError('Template module must be an object');
  }
  const manifest = snapshotManifest(dataProperty(value, 'manifest'));
  const recognize = dataProperty(value, 'recognize');
  const transform = dataProperty(value, 'transform');
  const resolve = dataProperty(value, 'resolve');
  const layout = dataProperty(value, 'layout');
  const paint = dataProperty(value, 'paint');
  const dispose = dataProperty(value, 'dispose', false);
  if (
    typeof recognize !== 'function' ||
    typeof transform !== 'function' ||
    typeof resolve !== 'function' ||
    typeof layout !== 'function' ||
    typeof paint !== 'function' ||
    (dispose !== undefined && typeof dispose !== 'function')
  ) {
    throw new TypeError('Template module stages must be callable');
  }
  return Object.freeze({
    manifest,
    recognize: (node: TemplateModuleNode, context: TemplateModuleStageContext) =>
      Reflect.apply(recognize, value, [node, context]) as boolean | Promise<boolean>,
    transform: (node: TemplateModuleNode, context: TemplateModuleStageContext) =>
      Reflect.apply(transform, value, [node, context]) as JsonValue | Promise<JsonValue>,
    resolve: (ir: JsonValue, context: TemplateModuleStageContext) =>
      Reflect.apply(resolve, value, [ir, context]) as JsonValue | Promise<JsonValue>,
    layout: (ir: JsonValue, context: TemplateModuleStageContext) =>
      Reflect.apply(layout, value, [ir, context]) as JsonValue | Promise<JsonValue>,
    paint: (ir: JsonValue, context: TemplateModuleStageContext) =>
      Reflect.apply(paint, value, [ir, context]) as
        | readonly PrintDisplayCommand[]
        | Promise<readonly PrintDisplayCommand[]>,
    ...(dispose === undefined
      ? {}
      : {
          dispose: () => Reflect.apply(dispose, value, []) as void | Promise<void>,
        }),
  });
}

function snapshotRegistryOptions(options: TemplateModuleRegistryOptions): {
  readonly apiVersion: string;
  readonly environment: TemplateModuleRegistryOptions['environment'];
} {
  if (options === null || typeof options !== 'object') {
    throw new TypeError('Template module registry options must be an object');
  }
  const snapshot = snapshotJsonValue(options, 'templateModule.options');
  if (snapshot === null || Array.isArray(snapshot) || typeof snapshot !== 'object') {
    throw new TypeError('Template module registry options must be an object');
  }
  const record = snapshot as Readonly<Record<string, JsonValue>>;
  if (
    typeof record.apiVersion !== 'string' ||
    !apiVersionPattern.test(record.apiVersion) ||
    typeof record.environment !== 'string' ||
    !environments.has(record.environment) ||
    Object.keys(record).some((key) => !['apiVersion', 'environment'].includes(key))
  ) {
    throw new TypeError('Template module registry options are invalid');
  }
  return snapshot as {
    readonly apiVersion: string;
    readonly environment: TemplateModuleRegistryOptions['environment'];
  };
}

/** Creates one deterministic, independently disposable template module registry. */
export function createTemplateModuleRegistry(
  options: TemplateModuleRegistryOptions,
): TemplateModuleRegistry {
  let configuration: ReturnType<typeof snapshotRegistryOptions>;
  try {
    configuration = snapshotRegistryOptions(options);
  } catch (cause) {
    fail(
      'TEMPLATE_MODULE_OPTIONS_INVALID',
      'validate',
      'Template module registry options are invalid',
      undefined,
      cause,
    );
  }
  const state: RegistryState = {
    apiVersion: configuration.apiVersion,
    environment: configuration.environment,
    records: new Map(),
    idleWaiters: new Set(),
    activeExecutions: 0,
    disposed: false,
  };
  const waitForIdle = (): Promise<void> => {
    if (state.activeExecutions === 0) return Promise.resolve();
    return new Promise((resolve) => {
      state.idleWaiters.add(resolve);
    });
  };
  const disposedModules = new WeakSet<TemplateModule>();
  const disposeModule = async (
    registered: TemplateModule,
  ): Promise<readonly TemplateModuleDiagnostic[]> => {
    if (disposedModules.has(registered)) return Object.freeze([]);
    disposedModules.add(registered);
    if (registered.dispose === undefined) return Object.freeze([]);
    try {
      await registered.dispose();
      return Object.freeze([]);
    } catch (cause) {
      return Object.freeze([
        diagnostic(
          'TEMPLATE_MODULE_DISPOSE_FAILED',
          'dispose',
          `Template module ${registered.manifest.id} disposal failed`,
          { moduleId: registered.manifest.id },
          cause,
        ),
      ]);
    }
  };
  const registry: TemplateModuleRegistry = Object.freeze({
    register(candidate: TemplateModule): () => Promise<readonly TemplateModuleDiagnostic[]> {
      if (state.disposed) {
        fail(
          'TEMPLATE_REGISTRY_DISPOSED',
          'validate',
          'Cannot register a template module after registry disposal',
        );
      }
      let registered: TemplateModule;
      try {
        registered = snapshotModule(candidate);
      } catch (cause) {
        fail(
          'TEMPLATE_MODULE_MANIFEST_INVALID',
          'validate',
          'Template module registration is invalid',
          undefined,
          cause,
        );
      }
      const manifest = registered.manifest;
      if (major(manifest.apiVersion) !== major(state.apiVersion)) {
        fail(
          'TEMPLATE_MODULE_VERSION_UNSUPPORTED',
          'validate',
          `Template module ${manifest.id} uses an unsupported API version`,
          { moduleId: manifest.id, apiVersion: manifest.apiVersion },
        );
      }
      if (!manifest.environments.includes(state.environment)) {
        fail(
          'TEMPLATE_MODULE_ENVIRONMENT_UNSUPPORTED',
          'validate',
          `Template module ${manifest.id} does not support ${state.environment}`,
          { moduleId: manifest.id, environment: state.environment },
        );
      }
      if (state.records.has(manifest.id)) {
        fail(
          'TEMPLATE_MODULE_DUPLICATE',
          'validate',
          `Template module ${manifest.id} is already registered`,
          { moduleId: manifest.id },
        );
      }
      state.records.set(manifest.id, registered);
      let active = true;
      let unregisterPromise: Promise<readonly TemplateModuleDiagnostic[]> | undefined;
      return (): Promise<readonly TemplateModuleDiagnostic[]> => {
        if (unregisterPromise !== undefined) return unregisterPromise;
        unregisterPromise = (async () => {
          if (!active) return Object.freeze([]);
          active = false;
          if (state.records.get(manifest.id) === registered) state.records.delete(manifest.id);
          await waitForIdle();
          return disposeModule(registered);
        })();
        return unregisterPromise;
      };
    },
    list(): readonly TemplateModule[] {
      return Object.freeze([...state.records.values()].sort(compareModules));
    },
    dispose(): Promise<readonly TemplateModuleDiagnostic[]> {
      if (state.disposePromise !== undefined) return state.disposePromise;
      state.disposed = true;
      const modules = [...state.records.values()].sort(compareModules);
      state.records.clear();
      state.disposePromise = (async () => {
        await waitForIdle();
        const diagnostics: TemplateModuleDiagnostic[] = [];
        for (const registered of modules) {
          diagnostics.push(...(await disposeModule(registered)));
        }
        return Object.freeze(diagnostics);
      })();
      return state.disposePromise;
    },
  });
  states.set(registry, state);
  return registry;
}

function positiveLimit(value: number | undefined, fallback: number, name: string): number {
  const result = value ?? fallback;
  if (!Number.isSafeInteger(result) || result < 1) {
    throw new TypeError(`${name} must be a positive safe integer`);
  }
  return result;
}

function snapshotLimits(
  value: Partial<TemplateModulePipelineLimits> | undefined,
): Readonly<TemplateModulePipelineLimits> {
  if (value !== undefined && (value === null || typeof value !== 'object')) {
    throw new TypeError('Template module limits must be an object');
  }
  const snapshot =
    value === undefined
      ? {}
      : (snapshotJsonValue(value, 'templateModule.limits') as Readonly<Record<string, JsonValue>>);
  if (
    Object.keys(snapshot).some(
      (key) =>
        !['maximumNodes', 'maximumCommands', 'maximumOutputBytes', 'maximumMilliseconds'].includes(
          key,
        ),
    )
  ) {
    throw new TypeError('Template module limits contain unknown properties');
  }
  return Object.freeze({
    maximumNodes: positiveLimit(
      snapshot.maximumNodes as number | undefined,
      defaults.maximumNodes,
      'maximumNodes',
    ),
    maximumCommands: positiveLimit(
      snapshot.maximumCommands as number | undefined,
      defaults.maximumCommands,
      'maximumCommands',
    ),
    maximumOutputBytes: positiveLimit(
      snapshot.maximumOutputBytes as number | undefined,
      defaults.maximumOutputBytes,
      'maximumOutputBytes',
    ),
    maximumMilliseconds: positiveLimit(
      snapshot.maximumMilliseconds as number | undefined,
      defaults.maximumMilliseconds,
      'maximumMilliseconds',
    ),
  });
}

function snapshotDocument(value: unknown): {
  readonly requiredModules: readonly string[];
  readonly nodes: readonly TemplateModuleNode[];
} {
  const snapshot = snapshotJsonValue(value, 'templateModule.document');
  if (snapshot === null || Array.isArray(snapshot) || typeof snapshot !== 'object') {
    fail('TEMPLATE_NODE_INVALID', 'validate', 'Template module document must be an object');
  }
  const record = snapshot as Readonly<Record<string, JsonValue>>;
  if (
    !Array.isArray(record.requiredModules) ||
    record.requiredModules.some((id) => typeof id !== 'string' || !idPattern.test(id)) ||
    new Set(record.requiredModules).size !== record.requiredModules.length ||
    !Array.isArray(record.nodes) ||
    Object.keys(record).some((key) => !['requiredModules', 'nodes'].includes(key))
  ) {
    fail('TEMPLATE_NODE_INVALID', 'validate', 'Template module document is invalid');
  }
  const seen = new Set<string>();
  const nodes = record.nodes.map((value, index): TemplateModuleNode => {
    if (value === null || Array.isArray(value) || typeof value !== 'object') {
      fail('TEMPLATE_NODE_INVALID', 'validate', `Template node ${index} is invalid`);
    }
    const candidate = value as Readonly<Record<string, JsonValue>>;
    if (
      typeof candidate.id !== 'string' ||
      !idPattern.test(candidate.id) ||
      seen.has(candidate.id) ||
      typeof candidate.type !== 'string' ||
      !typePattern.test(candidate.type) ||
      !Object.hasOwn(candidate, 'data') ||
      Object.keys(candidate).some((key) => !['id', 'type', 'data'].includes(key))
    ) {
      fail('TEMPLATE_NODE_INVALID', 'validate', `Template node ${index} is invalid`);
    }
    seen.add(candidate.id);
    return candidate as unknown as TemplateModuleNode;
  });
  return Object.freeze({
    requiredModules: record.requiredModules as readonly string[],
    nodes: Object.freeze(nodes),
  });
}

function finite(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value);
}

function plainRecord(value: unknown): Readonly<Record<string, unknown>> | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Readonly<Record<string, unknown>>)
    : undefined;
}

function exactKeys(value: Readonly<Record<string, unknown>>, allowed: readonly string[]): boolean {
  return Object.keys(value).every((key) => allowed.includes(key));
}

function validRect(value: unknown): boolean {
  const rect = plainRecord(value);
  return (
    rect !== undefined &&
    exactKeys(rect, ['x', 'y', 'width', 'height']) &&
    finite(rect.x) &&
    finite(rect.y) &&
    finite(rect.width) &&
    rect.width >= 0 &&
    finite(rect.height) &&
    rect.height >= 0
  );
}

function validString(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0;
}

function validateCommands(value: JsonValue): {
  readonly commands: readonly PrintDisplayCommand[];
  readonly count: number;
} {
  if (!Array.isArray(value)) {
    fail('TEMPLATE_DRAW_COMMAND_INVALID', 'render', 'Template paint output must be an array');
  }
  let count = 0;
  const visit = (command: JsonValue, depth: number): void => {
    count += 1;
    if (depth > 64) {
      fail(
        'TEMPLATE_DRAW_COMMAND_INVALID',
        'render',
        'Template draw command nesting exceeds 64 levels',
      );
    }
    const record = plainRecord(command);
    if (record === undefined || typeof record.kind !== 'string') {
      fail('TEMPLATE_DRAW_COMMAND_INVALID', 'render', 'Template draw command is invalid');
    }
    let valid = false;
    switch (record.kind) {
      case 'fill-rect':
        valid =
          exactKeys(record, ['kind', 'rect', 'color']) &&
          validRect(record.rect) &&
          validString(record.color);
        break;
      case 'stroke-rect':
        valid =
          exactKeys(record, ['kind', 'rect', 'color', 'width']) &&
          validRect(record.rect) &&
          validString(record.color) &&
          finite(record.width) &&
          record.width >= 0;
        break;
      case 'text':
        valid =
          exactKeys(record, [
            'kind',
            'text',
            'x',
            'y',
            'maxWidth',
            'fontFamily',
            'fontSize',
            'color',
            'horizontalAlign',
          ]) &&
          typeof record.text === 'string' &&
          finite(record.x) &&
          finite(record.y) &&
          finite(record.maxWidth) &&
          record.maxWidth >= 0 &&
          validString(record.fontFamily) &&
          finite(record.fontSize) &&
          record.fontSize >= 0 &&
          validString(record.color) &&
          ['left', 'center', 'right'].includes(record.horizontalAlign as string);
        break;
      case 'line':
        valid =
          exactKeys(record, ['kind', 'x1', 'y1', 'x2', 'y2', 'color', 'width']) &&
          finite(record.x1) &&
          finite(record.y1) &&
          finite(record.x2) &&
          finite(record.y2) &&
          validString(record.color) &&
          finite(record.width) &&
          record.width >= 0;
        break;
      case 'image':
        valid =
          exactKeys(record, ['kind', 'resourceId', 'rect', 'fit']) &&
          validString(record.resourceId) &&
          validRect(record.rect) &&
          ['contain', 'cover', 'fill'].includes(record.fit as string);
        break;
      case 'path':
        valid =
          exactKeys(record, ['kind', 'data', 'fill', 'stroke', 'width']) &&
          typeof record.data === 'string' &&
          (record.fill === undefined || validString(record.fill)) &&
          (record.stroke === undefined || validString(record.stroke)) &&
          (record.width === undefined || (finite(record.width) && record.width >= 0));
        break;
      case 'link':
        valid =
          exactKeys(record, ['kind', 'rect', 'href', 'label']) &&
          validRect(record.rect) &&
          validString(record.href) &&
          validString(record.label);
        break;
      case 'clip':
        valid =
          exactKeys(record, ['kind', 'rect', 'commands']) &&
          validRect(record.rect) &&
          Array.isArray(record.commands);
        if (valid) {
          (record.commands as readonly JsonValue[]).forEach((nested) => visit(nested, depth + 1));
        }
        break;
      case 'group': {
        const origin = plainRecord(record.origin);
        valid =
          exactKeys(record, ['kind', 'rotation', 'origin', 'commands']) &&
          finite(record.rotation) &&
          record.rotation >= 0 &&
          record.rotation < 360 &&
          origin !== undefined &&
          exactKeys(origin, ['x', 'y']) &&
          finite(origin.x) &&
          finite(origin.y) &&
          Array.isArray(record.commands);
        if (valid) {
          (record.commands as readonly JsonValue[]).forEach((nested) => visit(nested, depth + 1));
        }
        break;
      }
    }
    if (!valid) {
      fail(
        'TEMPLATE_DRAW_COMMAND_INVALID',
        'render',
        `Template draw command kind ${String(record.kind)} is invalid`,
        { kind: String(record.kind) },
      );
    }
  };
  value.forEach((command) => visit(command, 0));
  return {
    commands: value as unknown as readonly PrintDisplayCommand[],
    count,
  };
}

/** Executes an atomic, bounded recognize → transform → resolve → layout → paint pipeline. */
async function executePipeline(
  registry: TemplateModuleRegistry,
  options: ExecuteTemplateModulePipelineOptions,
): Promise<TemplateModulePipelineOutput> {
  const state = states.get(registry);
  if (state === undefined || state.disposed) {
    fail(
      'TEMPLATE_REGISTRY_DISPOSED',
      'execute',
      'Template module registry is disposed or invalid',
    );
  }
  if (options === null || typeof options !== 'object') {
    fail('TEMPLATE_MODULE_OPTIONS_INVALID', 'validate', 'Template pipeline options are invalid');
  }
  const signal = dataProperty(options, 'signal');
  const documentSource = dataProperty(options, 'document');
  const limitsSource = dataProperty(options, 'limits', false);
  const clockSource = dataProperty(options, 'clock', false);
  if (
    !(signal instanceof AbortSignal) ||
    (clockSource !== undefined && typeof clockSource !== 'function')
  ) {
    fail('TEMPLATE_MODULE_OPTIONS_INVALID', 'validate', 'Template pipeline options are invalid');
  }
  let limits: Readonly<TemplateModulePipelineLimits>;
  try {
    limits = snapshotLimits(limitsSource as Partial<TemplateModulePipelineLimits> | undefined);
  } catch (cause) {
    fail(
      'TEMPLATE_MODULE_OPTIONS_INVALID',
      'validate',
      'Template pipeline limits are invalid',
      undefined,
      cause,
    );
  }
  const document = snapshotDocument(documentSource);
  if (document.nodes.length > limits.maximumNodes) {
    fail(
      'TEMPLATE_NODE_LIMIT_EXCEEDED',
      'validate',
      `Template document exceeds the ${limits.maximumNodes} node limit`,
      { actual: document.nodes.length, maximum: limits.maximumNodes },
    );
  }
  const modules = registry.list();
  const modulesById = new Map(modules.map((registered) => [registered.manifest.id, registered]));
  const missing = document.requiredModules.filter((id) => !modulesById.has(id));
  if (missing.length > 0) {
    fail(
      'TEMPLATE_MODULE_MISSING',
      'resolve',
      `Required template module ${missing[0]} is unavailable`,
      { missing },
    );
  }
  const clock =
    clockSource === undefined ? Date.now : () => Reflect.apply(clockSource, options, []) as number;
  const started = clock();
  const guard = (): void => {
    if (signal.aborted) {
      fail('TEMPLATE_PIPELINE_ABORTED', 'execute', 'Template module pipeline was aborted');
    }
    const elapsed = clock() - started;
    if (!Number.isFinite(elapsed) || elapsed > limits.maximumMilliseconds) {
      fail(
        'TEMPLATE_TIME_LIMIT_EXCEEDED',
        'execute',
        `Template module pipeline exceeded ${limits.maximumMilliseconds}ms`,
        { maximum: limits.maximumMilliseconds },
      );
    }
  };
  const context = (moduleId: string, nodeId: string): TemplateModuleStageContext =>
    Object.freeze({ moduleId, nodeId, signal });
  const invoke = async <Result>(
    stage: string,
    moduleId: string,
    nodeId: string,
    operation: () => Result | Promise<Result>,
  ): Promise<Result> => {
    guard();
    try {
      const result = await operation();
      guard();
      return result;
    } catch (cause) {
      if (cause instanceof TemplateModuleSdkError) throw cause;
      fail(
        'TEMPLATE_STAGE_FAILED',
        stage === 'paint' ? 'render' : stage === 'layout' ? 'layout' : 'execute',
        `Template module ${moduleId} ${stage} failed for node ${nodeId}`,
        { moduleId, nodeId, moduleStage: stage },
        cause,
      );
    }
  };

  const owners = new Map<string, TemplateModule>();
  for (const node of document.nodes) {
    const candidates: TemplateModule[] = [];
    for (const registered of modules) {
      const recognized = await invoke('recognize', registered.manifest.id, node.id, () =>
        registered.recognize(node, context(registered.manifest.id, node.id)),
      );
      if (typeof recognized !== 'boolean') {
        fail(
          'TEMPLATE_STAGE_FAILED',
          'execute',
          `Template module ${registered.manifest.id} recognize returned a non-boolean value`,
          {
            moduleId: registered.manifest.id,
            nodeId: node.id,
            moduleStage: 'recognize',
          },
        );
      }
      if (recognized) candidates.push(registered);
    }
    if (candidates.length > 1) {
      fail(
        'TEMPLATE_NODE_OWNERSHIP_CONFLICT',
        'resolve',
        `Template node ${node.id} is recognized by multiple modules`,
        {
          nodeId: node.id,
          moduleIds: candidates.map(({ manifest }) => manifest.id),
        },
      );
    }
    if (candidates[0] !== undefined) owners.set(node.id, candidates[0]);
  }

  let outputBytes = 0;
  let commandCount = 0;
  const outputs: TemplateModuleNodeOutput[] = [];
  const commands: PrintDisplayCommand[] = [];
  const snapshotOutput = (
    value: unknown,
    stage: string,
    moduleId: string,
    nodeId: string,
  ): JsonValue => {
    let snapshot: JsonValue;
    try {
      snapshot = snapshotJsonValue(value, `templateModule.${stage}`);
    } catch (cause) {
      fail(
        'TEMPLATE_STAGE_FAILED',
        'execute',
        `Template module ${moduleId} ${stage} returned invalid output`,
        { moduleId, nodeId, moduleStage: stage },
        cause,
      );
    }
    outputBytes += jsonSnapshotBytes(snapshot);
    if (outputBytes > limits.maximumOutputBytes) {
      fail(
        'TEMPLATE_OUTPUT_LIMIT_EXCEEDED',
        'execute',
        `Template module output exceeds ${limits.maximumOutputBytes} bytes`,
        { actual: outputBytes, maximum: limits.maximumOutputBytes },
      );
    }
    return snapshot;
  };

  for (const node of document.nodes) {
    const owner = owners.get(node.id);
    if (owner === undefined) continue;
    const moduleId = owner.manifest.id;
    const stageContext = context(moduleId, node.id);
    const transformed = snapshotOutput(
      await invoke('transform', moduleId, node.id, () => owner.transform(node, stageContext)),
      'transform',
      moduleId,
      node.id,
    );
    const resolved = snapshotOutput(
      await invoke('resolve', moduleId, node.id, () => owner.resolve(transformed, stageContext)),
      'resolve',
      moduleId,
      node.id,
    );
    const laidOut = snapshotOutput(
      await invoke('layout', moduleId, node.id, () => owner.layout(resolved, stageContext)),
      'layout',
      moduleId,
      node.id,
    );
    const painted = snapshotOutput(
      await invoke('paint', moduleId, node.id, () => owner.paint(laidOut, stageContext)),
      'paint',
      moduleId,
      node.id,
    );
    const validated = validateCommands(painted);
    commandCount += validated.count;
    if (commandCount > limits.maximumCommands) {
      fail(
        'TEMPLATE_COMMAND_LIMIT_EXCEEDED',
        'render',
        `Template module output exceeds the ${limits.maximumCommands} command limit`,
        { actual: commandCount, maximum: limits.maximumCommands },
      );
    }
    commands.push(...validated.commands);
    outputs.push(
      Object.freeze({
        nodeId: node.id,
        moduleId,
        ir: laidOut,
        commands: validated.commands,
      }),
    );
  }
  return Object.freeze({
    nodes: Object.freeze(outputs),
    commands: Object.freeze(commands),
  });
}

/** Executes an atomic, bounded recognize → transform → resolve → layout → paint pipeline. */
export function executeTemplateModulePipeline(
  registry: TemplateModuleRegistry,
  options: ExecuteTemplateModulePipelineOptions,
): Promise<TemplateModulePipelineOutput> {
  const state = states.get(registry);
  if (state === undefined || state.disposed) {
    return Promise.reject(
      new TemplateModuleSdkError([
        diagnostic(
          'TEMPLATE_REGISTRY_DISPOSED',
          'execute',
          'Template module registry is disposed or invalid',
        ),
      ]),
    );
  }
  state.activeExecutions += 1;
  return executePipeline(registry, options).finally(() => {
    state.activeExecutions -= 1;
    if (state.activeExecutions === 0) {
      for (const resolve of state.idleWaiters) resolve();
      state.idleWaiters.clear();
    }
  });
}
