import type { LocaleDefinition } from 'tego-sheet';
import type { WorkbookData } from '../../../src/core';
import type { VisualFixtureName } from '../parity';

export type VisualLocale = 'de' | 'en';
export type VisualMaskName = 'blinking-caret' | 'native-scrollbars';

export interface VisualFixture {
  readonly name: VisualFixtureName;
  readonly workbook: WorkbookData;
  readonly locale?: VisualLocale;
  readonly masks?: readonly VisualMaskName[];
}

export type PublicLocale = LocaleDefinition;
