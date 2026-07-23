import type { ChangeEvent, ReactElement } from 'react';
import type {
  BindingId,
  Diagnostic,
  DocumentCellAddress,
  DocumentCellRange,
  DocumentSheetId,
} from '../../document';
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
  /** Current grid selection used when creating bindings and print targets. */
  readonly selection?: DocumentCellRange;
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

function firstTarget(template: SpreadsheetTemplate): DocumentCellRange {
  const profileTarget = template.printProfiles[0]?.targets[0];
  if (profileTarget?.type === 'range') return profileTarget.range;
  if (profileTarget?.type === 'ranges' && profileTarget.ranges[0] !== undefined) {
    return profileTarget.ranges[0];
  }
  const binding = template.bindings[0];
  if (binding?.type === 'value') {
    return {
      sheetId: binding.target.sheetId,
      start: { row: binding.target.row, column: binding.target.column },
      end: { row: binding.target.row, column: binding.target.column },
    };
  }
  if (binding !== undefined) return binding.range;
  return {
    sheetId:
      profileTarget?.type === 'sheet' ? profileTarget.sheetId : ('sheet-1' as DocumentSheetId),
    start: { row: 0, column: 0 },
    end: { row: 0, column: 0 },
  };
}

function bindingTarget(range: DocumentCellRange): DocumentCellAddress {
  return { sheetId: range.sheetId, row: range.start.row, column: range.start.column };
}

/** Accessible property panel that edits the same immutable model consumed by the SDK. */
export function TemplateDesigner({
  template,
  diagnostics,
  onChange,
  onLocateBinding,
  selection,
}: TemplateDesignerProps): ReactElement {
  const selectedRange = selection ?? firstTarget(template);
  const updateBinding = (
    bindingId: BindingId,
    update: (binding: TemplateBinding) => TemplateBinding,
  ) =>
    onChange({
      ...template,
      bindings: template.bindings.map((binding) =>
        binding.id === bindingId ? update(binding) : binding,
      ),
    });
  const changeBinding =
    (bindingId: BindingId) =>
    (event: ChangeEvent<HTMLInputElement>): void => {
      updateBinding(bindingId, (binding) =>
        updateBindingExpression(binding, event.currentTarget.value),
      );
    };
  const addBinding = (type: TemplateBinding['type']): void => {
    const id = `binding-${template.bindings.length + 1}` as BindingId;
    const binding: TemplateBinding =
      type === 'value'
        ? {
            id,
            type,
            target: bindingTarget(selectedRange),
            expression: 'value',
          }
        : type === 'repeat-rows'
          ? {
              id,
              type,
              range: selectedRange,
              source: 'items',
              empty: 'remove',
              pageBreak: 'auto',
            }
          : { id, type, range: selectedRange, when: 'visible' };
    onChange({ ...template, bindings: [...template.bindings, binding] });
  };
  const updateFirstProfile = (
    update: (
      profile: SpreadsheetTemplate['printProfiles'][number],
    ) => SpreadsheetTemplate['printProfiles'][number],
  ): void => {
    if (template.printProfiles[0] === undefined) return;
    onChange({
      ...template,
      printProfiles: template.printProfiles.map((profile, index) =>
        index === 0 ? update(profile) : profile,
      ),
    });
  };
  return (
    <aside aria-label="Template designer">
      <h2>{template.name}</h2>
      <section aria-label="Bindings">
        <div role="group" aria-label="Add binding">
          <button type="button" onClick={() => addBinding('value')}>
            Add value
          </button>
          <button type="button" onClick={() => addBinding('repeat-rows')}>
            Add repeat rows
          </button>
          <button type="button" onClick={() => addBinding('conditional-range')}>
            Add conditional range
          </button>
        </div>
        {template.bindings.map((binding) => (
          <fieldset key={binding.id}>
            <legend>{binding.id}</legend>
            <label>
              Expression
              <input
                aria-label={`Expression for ${binding.id}`}
                value={bindingExpression(binding)}
                onChange={changeBinding(binding.id)}
              />
            </label>
            {binding.type === 'value' ? (
              <label>
                Formatter
                <input
                  aria-label={`Formatter for ${binding.id}`}
                  value={binding.formatter ?? ''}
                  onChange={(event) =>
                    updateBinding(binding.id, (current) =>
                      current.type === 'value'
                        ? {
                            ...current,
                            ...(event.currentTarget.value === ''
                              ? { formatter: undefined }
                              : { formatter: event.currentTarget.value }),
                          }
                        : current,
                    )
                  }
                />
              </label>
            ) : binding.type === 'repeat-rows' ? (
              <>
                <label>
                  Empty collection
                  <select
                    aria-label={`Empty policy for ${binding.id}`}
                    value={binding.empty}
                    onChange={(event) =>
                      updateBinding(binding.id, (current) =>
                        current.type === 'repeat-rows'
                          ? {
                              ...current,
                              empty: event.currentTarget.value as typeof current.empty,
                            }
                          : current,
                      )
                    }
                  >
                    <option value="remove">Remove rows</option>
                    <option value="keep-template-row">Keep template row</option>
                  </select>
                </label>
                <label>
                  Item page break
                  <select
                    aria-label={`Page break for ${binding.id}`}
                    value={binding.pageBreak}
                    onChange={(event) =>
                      updateBinding(binding.id, (current) =>
                        current.type === 'repeat-rows'
                          ? {
                              ...current,
                              pageBreak: event.currentTarget.value as typeof current.pageBreak,
                            }
                          : current,
                      )
                    }
                  >
                    <option value="auto">Automatic</option>
                    <option value="before-each-item">Before each item</option>
                  </select>
                </label>
              </>
            ) : null}
            <button
              type="button"
              onClick={() =>
                onChange({
                  ...template,
                  bindings: template.bindings.filter(({ id }) => id !== binding.id),
                })
              }
            >
              Delete {binding.id}
            </button>
          </fieldset>
        ))}
      </section>
      <section aria-label="Print profiles">
        {template.printProfiles[0] === undefined ? (
          <p>No print profile</p>
        ) : (
          <>
            <h3>{template.printProfiles[0].name}</h3>
            <label>
              Paper
              <select
                aria-label="Paper"
                value={template.printProfiles[0].page.paper.type}
                onChange={(event) =>
                  updateFirstProfile((profile) => ({
                    ...profile,
                    page: {
                      ...profile.page,
                      paper: { type: event.currentTarget.value as 'A4' | 'A5' | 'Letter' },
                    },
                  }))
                }
              >
                <option value="A4">A4</option>
                <option value="A5">A5</option>
                <option value="Letter">Letter</option>
              </select>
            </label>
            <label>
              Orientation
              <select
                aria-label="Orientation"
                value={template.printProfiles[0].page.orientation}
                onChange={(event) =>
                  updateFirstProfile((profile) => ({
                    ...profile,
                    page: {
                      ...profile.page,
                      orientation: event.currentTarget.value as 'portrait' | 'landscape',
                    },
                  }))
                }
              >
                <option value="portrait">Portrait</option>
                <option value="landscape">Landscape</option>
              </select>
            </label>
            <label>
              <input
                type="checkbox"
                aria-label="Show gridlines"
                checked={template.printProfiles[0].showGridlines}
                onChange={(event) =>
                  updateFirstProfile((profile) => ({
                    ...profile,
                    showGridlines: event.currentTarget.checked,
                  }))
                }
              />
              Show gridlines
            </label>
            <label>
              <input
                type="checkbox"
                aria-label="Show headings"
                checked={template.printProfiles[0].showHeadings}
                onChange={(event) =>
                  updateFirstProfile((profile) => ({
                    ...profile,
                    showHeadings: event.currentTarget.checked,
                  }))
                }
              />
              Show headings
            </label>
            <button
              type="button"
              onClick={() =>
                updateFirstProfile((profile) => ({
                  ...profile,
                  targets: [{ type: 'range', range: selectedRange }],
                }))
              }
            >
              Use selection as print range
            </button>
          </>
        )}
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
