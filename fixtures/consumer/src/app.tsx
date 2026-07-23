import { useRef } from 'react';
import { createSpreadsheetDocument, TegoSheet, type TegoSheetHandle } from 'tego-sheet';
import { zhCN } from 'tego-sheet/locales/zh-cn';

const document = createSpreadsheetDocument({ sheetName: 'Consumer' });

export function App() {
  const sheet = useRef<TegoSheetHandle>(null);
  return <TegoSheet ref={sheet} defaultDocument={document} locale={zhCN} />;
}
