import type { ValidationRequest, ValidationResult } from './model';
import { staticRuleAccepted, validationFailure } from './synchronous';

/** Restricted context supplied to a host validation resolver. */
export interface ValidationResolverContext {
  /** Cancellation signal owned by the validation engine. */
  readonly signal: AbortSignal;
}

/** Host list resolver with no implicit I/O capability. */
export type ValidationResolver = (
  context: ValidationResolverContext,
) => Promise<readonly string[]> | readonly string[];

/** Restricted context supplied to a custom-formula evaluator. */
export interface ValidationFormulaContext extends ValidationResolverContext {
  /** Immutable candidate and rule being evaluated. */
  readonly request: ValidationRequest;
}

/** Host-injected custom-formula evaluator with no implicit document mutation capability. */
export type ValidationFormulaEvaluator = (
  formula: string,
  context: ValidationFormulaContext,
) => boolean | Promise<boolean>;

/** Isolated resolver registry. */
export interface ValidationResolverRegistry {
  /** Registers a resolver and returns its disposer. */
  register(id: string, resolver: ValidationResolver): () => void;
  /** Resolves a previously registered resolver. */
  resolve(id: string): ValidationResolver | undefined;
}

/** Creates a validation resolver registry. */
export function createValidationResolverRegistry(): ValidationResolverRegistry {
  const entries = new Map<string, ValidationResolver>();
  return {
    register(id, resolver) {
      if (!/^[A-Za-z][A-Za-z0-9._-]*$/u.test(id) || entries.has(id)) {
        throw new TypeError(`Invalid or duplicate validation resolver ${id}`);
      }
      entries.set(id, resolver);
      return () => {
        entries.delete(id);
      };
    },
    resolve(id) {
      return entries.get(id);
    },
  };
}

/** Optional resolver capabilities and resource limits for validation. */
export interface ValidationEngineOptions {
  /** Optional host resolver registry. */
  readonly resolvers?: ValidationResolverRegistry;
  /** Optional restricted evaluator for custom-formula rules. */
  readonly evaluateCustomFormula?: ValidationFormulaEvaluator;
  /** Resource limits for dynamic validation sources. */
  readonly limits?: {
    /** Maximum number of values accepted from a static or resolved list. */
    readonly maxListItems?: number;
    /** Hard deadline for a resolver or custom-formula evaluator. */
    readonly resolverTimeoutMs?: number;
    /** Maximum custom-formula source length. */
    readonly maxFormulaLength?: number;
  };
}

/** Side-effect-free validation service used before document mutation. */
export interface ValidationEngine {
  /** Validates one immutable edit candidate. */
  validate(request: ValidationRequest): Promise<ValidationResult>;
  /** Aborts pending work and prevents later validation. */
  dispose?(): void;
}

type SourceErrorCode =
  | 'VALIDATION_SOURCE_ERROR'
  | 'VALIDATION_SOURCE_TOO_LARGE'
  | 'VALIDATION_SOURCE_TIMEOUT'
  | 'VALIDATION_SOURCE_ABORTED';

function sourceFailure(request: ValidationRequest, code: SourceErrorCode): ValidationResult {
  return {
    status: 'error',
    code,
    diagnostics: [{ code, ruleId: request.rule.id }],
  };
}

/** Creates a side-effect-free validation engine. */
export function createValidationEngine(options: ValidationEngineOptions = {}): ValidationEngine {
  const maxListItems = options.limits?.maxListItems ?? 10_000;
  const timeoutMs = options.limits?.resolverTimeoutMs ?? 5_000;
  const maxFormulaLength = options.limits?.maxFormulaLength ?? 4_096;
  const active = new Set<AbortController>();
  let disposed = false;

  async function bounded<Value>(
    request: ValidationRequest,
    operation: (signal: AbortSignal) => Value | Promise<Value>,
  ): Promise<
    | { readonly ok: true; readonly value: Value }
    | { readonly ok: false; readonly code: SourceErrorCode }
  > {
    if (disposed || request.signal?.aborted === true) {
      return { ok: false, code: 'VALIDATION_SOURCE_ABORTED' };
    }
    const controller = new AbortController();
    active.add(controller);
    const abort = () => controller.abort(request.signal?.reason);
    request.signal?.addEventListener('abort', abort, { once: true });
    let timeout: ReturnType<typeof setTimeout> | undefined;
    try {
      const operationPromise = Promise.resolve().then(() => operation(controller.signal));
      const abortPromise = new Promise<never>((_resolve, reject) => {
        controller.signal.addEventListener(
          'abort',
          () => reject(controller.signal.reason ?? new Error('Validation aborted')),
          { once: true },
        );
      });
      const timeoutPromise = new Promise<never>((_resolve, reject) => {
        timeout = setTimeout(() => {
          const error = new Error('Validation source timed out');
          controller.abort(error);
          reject(error);
        }, timeoutMs);
      });
      const value = await Promise.race([operationPromise, abortPromise, timeoutPromise]);
      return { ok: true, value };
    } catch {
      return {
        ok: false,
        code:
          (request.signal?.aborted ?? false) || disposed
            ? 'VALIDATION_SOURCE_ABORTED'
            : controller.signal.aborted
              ? 'VALIDATION_SOURCE_TIMEOUT'
              : 'VALIDATION_SOURCE_ERROR',
      };
    } finally {
      if (timeout !== undefined) clearTimeout(timeout);
      request.signal?.removeEventListener('abort', abort);
      active.delete(controller);
    }
  }

  return {
    async validate(request) {
      if (disposed || request.signal?.aborted === true) {
        return sourceFailure(request, 'VALIDATION_SOURCE_ABORTED');
      }
      if (request.value.type === 'blank' && request.rule.allowBlank) {
        return { status: 'accepted', diagnostics: [] };
      }

      let accepted = staticRuleAccepted(request);
      if (request.rule.type === 'list') {
        const source = request.rule.predicate.source;
        let values: readonly string[];
        if (source.type === 'static') {
          values = source.values;
        } else {
          const resolver = options.resolvers?.resolve(source.id);
          if (resolver === undefined) {
            return sourceFailure(request, 'VALIDATION_SOURCE_ERROR');
          }
          const resolution = await bounded(request, (signal) => resolver({ signal }));
          if (!resolution.ok) return sourceFailure(request, resolution.code);
          values = resolution.value;
        }
        if (values.length > maxListItems) {
          return sourceFailure(request, 'VALIDATION_SOURCE_TOO_LARGE');
        }
        accepted = request.value.type === 'string' && new Set(values).has(request.value.value);
      } else if (request.rule.type === 'custom-formula') {
        const formula = request.rule.predicate.formula;
        if (
          options.evaluateCustomFormula === undefined ||
          formula.length > maxFormulaLength ||
          !formula.startsWith('=')
        ) {
          return sourceFailure(request, 'VALIDATION_SOURCE_ERROR');
        }
        const evaluation = await bounded(request, (signal) =>
          options.evaluateCustomFormula!(formula, { request, signal }),
        );
        if (!evaluation.ok) return sourceFailure(request, evaluation.code);
        accepted = evaluation.value;
      }

      return accepted === true
        ? { status: 'accepted', diagnostics: [] }
        : validationFailure(request, request.rule.behavior === 'warn' ? 'warning' : 'rejected');
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      for (const controller of active) {
        controller.abort(new Error('Validation engine disposed'));
      }
      active.clear();
    },
  };
}
