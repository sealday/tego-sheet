import { parsePlaygroundMode, type PlaygroundMode } from './playground-model';

export const PLAYGROUND_WORKSPACES = Object.freeze(['spreadsheet', 'output'] as const);
export type PlaygroundWorkspace = (typeof PLAYGROUND_WORKSPACES)[number];

export interface PlaygroundLocation {
  readonly workspace: PlaygroundWorkspace;
  readonly mode: PlaygroundMode;
}

export function readPlaygroundLocation(search: string): PlaygroundLocation {
  const params = new URLSearchParams(search);
  const workspace = PLAYGROUND_WORKSPACES.includes(params.get('workspace') as PlaygroundWorkspace)
    ? (params.get('workspace') as PlaygroundWorkspace)
    : 'spreadsheet';
  return Object.freeze({ workspace, mode: parsePlaygroundMode(params.get('mode')) });
}

export function writePlaygroundLocation(
  pathname: string,
  search: string,
  next: PlaygroundLocation,
): string {
  const params = new URLSearchParams(search);
  params.set('workspace', next.workspace);
  params.set('mode', next.mode);
  return `${pathname}?${params.toString()}`;
}
