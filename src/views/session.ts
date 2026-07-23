import type { FilterView } from './model';

/** Ephemeral selected-view state that never changes a document revision. */
export interface FilterViewSession {
  /** Monotonic session-only revision. */
  readonly revision: number;
  /** Constant proving view selection does not mutate the document. */
  readonly documentRevision: 0;
  /** Currently selected session view. */
  readonly selected: FilterView | undefined;
  /** Selects or clears the session view. */
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
