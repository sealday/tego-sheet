import { describe, expect, it } from 'vitest';
import {
  readPlaygroundLocation,
  writePlaygroundLocation,
} from '../../../website/src/components/playground/playground-workspace';

describe('Playground workspace URL state', () => {
  it('keeps legacy mode links in the Spreadsheet workspace', () => {
    expect(readPlaygroundLocation('?mode=controlled')).toEqual({
      workspace: 'spreadsheet',
      mode: 'controlled',
    });
  });

  it('selects Output Studio without discarding the remembered spreadsheet mode', () => {
    expect(readPlaygroundLocation('?workspace=output&mode=locales')).toEqual({
      workspace: 'output',
      mode: 'locales',
    });
  });

  it('canonicalizes invalid values and preserves unrelated parameters', () => {
    expect(
      writePlaygroundLocation('/tego-sheet/playground', '?theme=dark&workspace=nope&mode=nope', {
        workspace: 'spreadsheet',
        mode: 'uncontrolled',
      }),
    ).toBe('/tego-sheet/playground?theme=dark&workspace=spreadsheet&mode=uncontrolled');
  });
});
