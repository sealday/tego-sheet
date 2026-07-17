import type { JsonValue } from 'tego-sheet';

export const PLAYGROUND_MODES = Object.freeze([
  'uncontrolled',
  'controlled',
  'custom-chrome',
  'locales',
  'legacy-json',
] as const);

export type PlaygroundMode = (typeof PLAYGROUND_MODES)[number];

export type PlaygroundCallbackName =
  | 'onChange'
  | 'onActiveSheetChange'
  | 'onSelectionChange'
  | 'onCellEdit'
  | 'onPaste'
  | 'onError';

export interface PlaygroundEvent {
  readonly sequence: number;
  readonly callback: PlaygroundCallbackName;
  readonly payload: JsonValue;
}

const EVENT_LIMIT = 50;

export function parsePlaygroundMode(value: string | null): PlaygroundMode {
  return PLAYGROUND_MODES.includes(value as PlaygroundMode)
    ? (value as PlaygroundMode)
    : 'uncontrolled';
}

function cloneAndFreezeJson(value: JsonValue): JsonValue {
  if (value === null || typeof value !== 'object') return value;

  if (Array.isArray(value)) {
    return Object.freeze(value.map((entry) => cloneAndFreezeJson(entry)));
  }

  const clone: Record<string, JsonValue> = {};
  for (const [key, entry] of Object.entries(value)) {
    clone[key] = cloneAndFreezeJson(entry);
  }
  return Object.freeze(clone);
}

function cloneAndFreezeEvent(event: PlaygroundEvent): PlaygroundEvent {
  return Object.freeze({
    sequence: event.sequence,
    callback: event.callback,
    payload: cloneAndFreezeJson(event.payload),
  });
}

export function appendPlaygroundEvent(
  events: readonly PlaygroundEvent[],
  event: PlaygroundEvent,
): readonly PlaygroundEvent[] {
  const retained = events.slice(-(EVENT_LIMIT - 1));
  return Object.freeze([...retained, event].map((entry) => cloneAndFreezeEvent(entry)));
}
