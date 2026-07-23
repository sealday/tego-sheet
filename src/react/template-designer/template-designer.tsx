import type { ChangeEvent, ReactElement } from 'react';
import type {
  BindingId,
  Diagnostic,
  DocumentCellAddress,
  DocumentCellRange,
  DocumentSheetId,
} from '../../document';
import type { SpreadsheetTemplate, TemplateBinding, TemplatePrintProfile } from '../../template';

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
  const updateProfile = (
    profileId: string,
    update: (profile: TemplatePrintProfile) => TemplatePrintProfile,
  ): void => {
    onChange({
      ...template,
      printProfiles: template.printProfiles.map((profile) =>
        profile.id === profileId ? update(profile) : profile,
      ),
    });
  };
  const addProfile = (): void => {
    const profile: TemplatePrintProfile = {
      id: `profile-${template.printProfiles.length + 1}`,
      name: `Print profile ${template.printProfiles.length + 1}`,
      targets: [{ type: 'range', range: selectedRange }],
      page: {
        paper: { type: 'A4' },
        orientation: 'portrait',
        margins: { top: 20, right: 20, bottom: 20, left: 20 },
        scale: { type: 'fixed', value: 1 },
      },
      manualBreaks: [],
      showGridlines: true,
      showHeadings: false,
    };
    onChange({ ...template, printProfiles: [...template.printProfiles, profile] });
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
        <button type="button" onClick={addProfile}>
          Add print profile
        </button>
        {template.printProfiles.length === 0 ? <p>No print profile</p> : null}
        {template.printProfiles.map((profile, profileIndex) => (
          <fieldset key={profile.id}>
            <legend>{profile.name}</legend>
            <label>
              Profile name
              <input
                aria-label={`Profile name for ${profile.id}`}
                value={profile.name}
                onChange={(event) =>
                  updateProfile(profile.id, (current) => ({
                    ...current,
                    name: event.currentTarget.value,
                  }))
                }
              />
            </label>
            <label>
              Paper
              <select
                aria-label={profileIndex === 0 ? 'Paper' : `Paper for ${profile.id}`}
                value={profile.page.paper.type}
                onChange={(event) =>
                  updateProfile(profile.id, (current) => ({
                    ...current,
                    page: {
                      ...current.page,
                      paper:
                        event.currentTarget.value === 'custom'
                          ? { type: 'custom', width: 210, height: 297 }
                          : {
                              type: event.currentTarget.value as 'A4' | 'A5' | 'Letter',
                            },
                    },
                  }))
                }
              >
                <option value="A4">A4</option>
                <option value="A5">A5</option>
                <option value="Letter">Letter</option>
                <option value="custom">Custom</option>
              </select>
            </label>
            {profile.page.paper.type === 'custom' ? (
              <>
                <label>
                  Paper width
                  <input
                    type="number"
                    aria-label={`Paper width for ${profile.id}`}
                    value={profile.page.paper.width}
                    onChange={(event) =>
                      updateProfile(profile.id, (current) => ({
                        ...current,
                        page: {
                          ...current.page,
                          paper: {
                            type: 'custom',
                            width: Number(event.currentTarget.value),
                            height:
                              current.page.paper.type === 'custom'
                                ? current.page.paper.height
                                : 297,
                          },
                        },
                      }))
                    }
                  />
                </label>
                <label>
                  Paper height
                  <input
                    type="number"
                    aria-label={`Paper height for ${profile.id}`}
                    value={profile.page.paper.height}
                    onChange={(event) =>
                      updateProfile(profile.id, (current) => ({
                        ...current,
                        page: {
                          ...current.page,
                          paper: {
                            type: 'custom',
                            width:
                              current.page.paper.type === 'custom' ? current.page.paper.width : 210,
                            height: Number(event.currentTarget.value),
                          },
                        },
                      }))
                    }
                  />
                </label>
              </>
            ) : null}
            <label>
              Orientation
              <select
                aria-label={`Orientation for ${profile.id}`}
                value={profile.page.orientation}
                onChange={(event) =>
                  updateProfile(profile.id, (current) => ({
                    ...current,
                    page: {
                      ...current.page,
                      orientation: event.currentTarget.value as 'portrait' | 'landscape',
                    },
                  }))
                }
              >
                <option value="portrait">Portrait</option>
                <option value="landscape">Landscape</option>
              </select>
            </label>
            {(['top', 'right', 'bottom', 'left'] as const).map((side) => (
              <label key={side}>
                {side} margin
                <input
                  type="number"
                  aria-label={`${side} margin for ${profile.id}`}
                  value={profile.page.margins[side]}
                  onChange={(event) =>
                    updateProfile(profile.id, (current) => ({
                      ...current,
                      page: {
                        ...current.page,
                        margins: {
                          ...current.page.margins,
                          [side]: Number(event.currentTarget.value),
                        },
                      },
                    }))
                  }
                />
              </label>
            ))}
            <label>
              Scale
              <select
                aria-label={`Scale for ${profile.id}`}
                value={profile.page.scale.type}
                onChange={(event) =>
                  updateProfile(profile.id, (current) => ({
                    ...current,
                    page: {
                      ...current.page,
                      scale:
                        event.currentTarget.value === 'fit-page'
                          ? { type: 'fit-page' }
                          : event.currentTarget.value === 'fit-width'
                            ? { type: 'fit-width', pages: 1 }
                            : { type: 'fixed', value: 1 },
                    },
                  }))
                }
              >
                <option value="fixed">Fixed</option>
                <option value="fit-width">Fit width</option>
                <option value="fit-page">Fit page</option>
              </select>
            </label>
            {profile.page.scale.type === 'fixed' ? (
              <label>
                Scale value
                <input
                  type="number"
                  aria-label={`Scale value for ${profile.id}`}
                  value={profile.page.scale.value}
                  onChange={(event) =>
                    updateProfile(profile.id, (current) => ({
                      ...current,
                      page: {
                        ...current.page,
                        scale: { type: 'fixed', value: Number(event.currentTarget.value) },
                      },
                    }))
                  }
                />
              </label>
            ) : profile.page.scale.type === 'fit-width' ? (
              <label>
                Fit width pages
                <input
                  type="number"
                  aria-label={`Fit width pages for ${profile.id}`}
                  value={profile.page.scale.pages}
                  onChange={(event) =>
                    updateProfile(profile.id, (current) => ({
                      ...current,
                      page: {
                        ...current.page,
                        scale: { type: 'fit-width', pages: Number(event.currentTarget.value) },
                      },
                    }))
                  }
                />
              </label>
            ) : null}
            <label>
              <input
                type="checkbox"
                aria-label={`Show gridlines for ${profile.id}`}
                checked={profile.showGridlines}
                onChange={(event) =>
                  updateProfile(profile.id, (current) => ({
                    ...current,
                    showGridlines: event.currentTarget.checked,
                  }))
                }
              />
              Show gridlines
            </label>
            <label>
              <input
                type="checkbox"
                aria-label={`Show headings for ${profile.id}`}
                checked={profile.showHeadings}
                onChange={(event) =>
                  updateProfile(profile.id, (current) => ({
                    ...current,
                    showHeadings: event.currentTarget.checked,
                  }))
                }
              />
              Show headings
            </label>
            <button
              type="button"
              onClick={() =>
                updateProfile(profile.id, (current) => ({
                  ...current,
                  targets: [{ type: 'range', range: selectedRange }],
                }))
              }
            >
              Use selection as print range
            </button>
            <button
              type="button"
              onClick={() =>
                updateProfile(profile.id, (current) => ({
                  ...current,
                  repeatRows: selectedRange,
                }))
              }
            >
              Use selection as repeat rows
            </button>
            <button
              type="button"
              onClick={() =>
                updateProfile(profile.id, (current) => ({
                  ...current,
                  repeatColumns: selectedRange,
                }))
              }
            >
              Use selection as repeat columns
            </button>
            <button
              type="button"
              onClick={() =>
                updateProfile(profile.id, (current) => ({
                  ...current,
                  repeatRows: undefined,
                  repeatColumns: undefined,
                }))
              }
            >
              Clear repeat titles
            </button>
            <button
              type="button"
              onClick={() =>
                updateProfile(profile.id, (current) => ({
                  ...current,
                  manualBreaks: [
                    ...current.manualBreaks,
                    { sheetId: selectedRange.sheetId, beforeRow: selectedRange.start.row },
                  ],
                }))
              }
            >
              Add page break at selection
            </button>
            {profile.manualBreaks.map((pageBreak, index) => (
              <button
                key={`${pageBreak.sheetId}-${pageBreak.beforeRow}-${index}`}
                type="button"
                onClick={() =>
                  updateProfile(profile.id, (current) => ({
                    ...current,
                    manualBreaks: current.manualBreaks.filter(
                      (_pageBreak, candidateIndex) => candidateIndex !== index,
                    ),
                  }))
                }
              >
                Remove page break {pageBreak.beforeRow}
              </button>
            ))}
            {(['header', 'footer'] as const).flatMap((band) =>
              (['left', 'center', 'right'] as const).map((slot) => (
                <label key={`${band}-${slot}`}>
                  {band} {slot}
                  <input
                    aria-label={`${band} ${slot} for ${profile.id}`}
                    value={profile[band]?.[slot] ?? ''}
                    onChange={(event) =>
                      updateProfile(profile.id, (current) => ({
                        ...current,
                        [band]: {
                          ...current[band],
                          [slot]: event.currentTarget.value,
                        },
                      }))
                    }
                  />
                </label>
              )),
            )}
            <button
              type="button"
              onClick={() =>
                onChange({
                  ...template,
                  printProfiles: template.printProfiles.filter(({ id }) => id !== profile.id),
                })
              }
            >
              Delete profile {profile.id}
            </button>
          </fieldset>
        ))}
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
