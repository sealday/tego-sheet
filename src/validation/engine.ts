import type { ValidationRequest, ValidationResult } from './model';

/** Restricted context supplied to a host validation resolver. */
export interface ValidationResolverContext {
  /** Cancellation signal owned by the validation engine. */
  readonly signal: AbortSignal;
}

/** Host list resolver with no implicit I/O capability. */
export type ValidationResolver = (
  context: ValidationResolverContext,
) => Promise<readonly string[]> | readonly string[];

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
  /** Resource limits for dynamic validation sources. */
  readonly limits?: {
    readonly maxListItems?: number;
    readonly resolverTimeoutMs?: number;
  };
}

function failure(request: ValidationRequest, status: 'rejected' | 'warning'): ValidationResult {
  return {
    status,
    code: 'VALIDATION_REJECTED',
    diagnostics: [{ code: 'VALIDATION_REJECTED', ruleId: request.rule.id }],
  };
}

/** Creates a side-effect-free validation engine. */
export function createValidationEngine(options: ValidationEngineOptions = {}): {
  validate(request: ValidationRequest): Promise<ValidationResult>;
} {
  const maxListItems = options.limits?.maxListItems ?? 10_000;
  const timeoutMs = options.limits?.resolverTimeoutMs ?? 5_000;
  return {
    async validate(request) {
      if (request.value.type === 'blank' && request.rule.allowBlank) {
        return { status: 'accepted', diagnostics: [] };
      }
      let accepted = false;
      if (request.rule.type === 'number') {
        accepted =
          request.value.type === 'number' &&
          request.value.value >= request.rule.predicate.minimum &&
          request.value.value <= request.rule.predicate.maximum;
      } else {
        const source = request.rule.predicate.source;
        let values: readonly string[];
        if (source.type === 'static') {
          values = source.values;
        } else {
          const resolver = options.resolvers?.resolve(source.id);
          if (resolver === undefined) {
            return {
              status: 'error',
              code: 'VALIDATION_SOURCE_ERROR',
              diagnostics: [{ code: 'VALIDATION_SOURCE_ERROR', ruleId: request.rule.id }],
            };
          }
          const controller = new AbortController();
          const abort = () => controller.abort(request.signal?.reason);
          request.signal?.addEventListener('abort', abort, { once: true });
          const timeout = setTimeout(
            () => controller.abort(new Error('Validation resolver timed out')),
            timeoutMs,
          );
          try {
            values = await resolver({ signal: controller.signal });
          } catch {
            return {
              status: 'error',
              code: 'VALIDATION_SOURCE_ERROR',
              diagnostics: [{ code: 'VALIDATION_SOURCE_ERROR', ruleId: request.rule.id }],
            };
          } finally {
            clearTimeout(timeout);
            request.signal?.removeEventListener('abort', abort);
          }
        }
        if (values.length > maxListItems) {
          return {
            status: 'error',
            code: 'VALIDATION_SOURCE_TOO_LARGE',
            diagnostics: [{ code: 'VALIDATION_SOURCE_TOO_LARGE', ruleId: request.rule.id }],
          };
        }
        accepted = request.value.type === 'string' && new Set(values).has(request.value.value);
      }
      return accepted
        ? { status: 'accepted', diagnostics: [] }
        : failure(request, request.rule.behavior === 'warn' ? 'warning' : 'rejected');
    },
  };
}
