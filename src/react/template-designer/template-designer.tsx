import type { ChangeEvent, ReactElement } from 'react';
import type { BindingId, Diagnostic } from '../../document';
import type { SpreadsheetTemplate, TemplateBinding } from '../../template';

/** Inputs for the SDK-model template property panel. */
export interface TemplateDesignerProps {
  /** Immutable template being edited. */
  readonly template: SpreadsheetTemplate;
  /** Current compiler or render diagnostics. */
  readonly diagnostics: readonly Diagnostic[];
  /** Receives the next immutable template snapshot. */
  readonly onChange: (template: SpreadsheetTemplate) => void;
  /** Requests that the host reveal a binding's sheet location. */
  readonly onLocateBinding?: (bindingId: BindingId) => void;
}

function bindingExpression(binding: TemplateBinding): string {
  if (binding.type === 'value') return binding.expression;
  return binding.type === 'repeat-rows' ? binding.source : binding.when;
}

function updateBindingExpression(binding: TemplateBinding, value: string): TemplateBinding {
  if (binding.type === 'value') return { ...binding, expression: value };
  return binding.type === 'repeat-rows'
    ? { ...binding, source: value }
    : { ...binding, when: value };
}

/** Accessible property panel that edits the same immutable model consumed by the SDK. */
export function TemplateDesigner({
  template,
  diagnostics,
  onChange,
  onLocateBinding,
}: TemplateDesignerProps): ReactElement {
  const changeBinding =
    (bindingId: BindingId) =>
    (event: ChangeEvent<HTMLInputElement>): void => {
      onChange({
        ...template,
        bindings: template.bindings.map((binding) =>
          binding.id === bindingId
            ? updateBindingExpression(binding, event.currentTarget.value)
            : binding,
        ),
      });
    };
  return (
    <aside aria-label="Template designer">
      <h2>{template.name}</h2>
      <section aria-label="Bindings">
        {template.bindings.map((binding) => (
          <label key={binding.id}>
            {binding.id}
            <input
              aria-label={`Expression for ${binding.id}`}
              value={bindingExpression(binding)}
              onChange={changeBinding(binding.id)}
            />
          </label>
        ))}
      </section>
      <section aria-label="Print profiles">
        <ul>
          {template.printProfiles.map((profile) => (
            <li key={profile.id}>{profile.name}</li>
          ))}
        </ul>
      </section>
      <section aria-label="Template diagnostics" aria-live="polite">
        {diagnostics.map((diagnostic, index) => {
          const bindingId = diagnostic.location?.bindingId;
          return bindingId === undefined ? (
            <p key={`${diagnostic.code}-${index}`}>{diagnostic.message}</p>
          ) : (
            <button
              key={`${diagnostic.code}-${bindingId}-${index}`}
              type="button"
              onClick={() => onLocateBinding?.(bindingId)}
            >
              {diagnostic.message}
            </button>
          );
        })}
      </section>
    </aside>
  );
}
