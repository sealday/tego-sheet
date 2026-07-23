import type { FilterView } from './model';

/** Ephemeral selected-view state that never changes a document revision. */
export interface FilterViewSession {
  readonly revision: number;
  readonly documentRevision: 0;
  readonly selected: FilterView | undefined;
  select(view: FilterView | undefined): void;
}

/** Creates isolated, session-owned filter-view selection state. */
export function createFilterViewSession(): FilterViewSession {
  let revision = 0;
  let selected: FilterView | undefined;
  return {
    get revision() {
      return revision;
    },
    documentRevision: 0,
    get selected() {
      return selected;
    },
    select(view) {
      selected = view;
      revision += 1;
    },
  };
}
