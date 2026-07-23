import type { DocumentCellAddress } from '../document';

/** @internal One latest-edit-wins validation lease for a document cell. */
export interface CellValidationLease {
  readonly signal: AbortSignal;
  readonly isCurrent: () => boolean;
  readonly abort: () => void;
  readonly release: () => void;
}

const activeByOwner = new WeakMap<object, Map<string, CellValidationLease>>();

function key(address: DocumentCellAddress): string {
  return `${address.sheetId}\u0000${address.row}\u0000${address.column}`;
}

/** @internal Supersedes pending validation for the same owner and cell. */
export function beginCellValidation(
  owner: object,
  address: DocumentCellAddress,
): CellValidationLease {
  const active = activeByOwner.get(owner) ?? new Map<string, CellValidationLease>();
  const cellKey = key(address);
  active.get(cellKey)?.abort();
  activeByOwner.set(owner, active);
  const controller = new AbortController();
  const isCurrent = () => active.get(cellKey) === lease && !controller.signal.aborted;
  const release = () => {
    if (active.get(cellKey) === lease) active.delete(cellKey);
    if (active.size === 0) activeByOwner.delete(owner);
  };
  const abort = () => {
    controller.abort();
    release();
  };
  const lease: CellValidationLease = { signal: controller.signal, isCurrent, abort, release };
  active.set(cellKey, lease);
  return lease;
}
