import BrowserOnly from '@docusaurus/BrowserOnly';
import { TegoSheet, type WorkbookData } from 'tego-sheet';
import 'tego-sheet/styles.css';
import { useState, type ReactElement } from 'react';

const previewFixture: WorkbookData = [
  {
    name: 'Release plan',
    freeze: 'A2',
    rows: {
      len: 18,
      0: {
        cells: {
          0: { text: 'Workstream' },
          1: { text: 'Owner' },
          2: { text: 'Status' },
          3: { text: 'Progress' },
        },
      },
      1: {
        cells: {
          0: { text: 'React API' },
          1: { text: 'Platform' },
          2: { text: 'Ready' },
          3: { text: '100%' },
        },
      },
      2: {
        cells: {
          0: { text: 'Documentation' },
          1: { text: 'DX' },
          2: { text: 'In review' },
          3: { text: '85%' },
        },
      },
      3: {
        cells: {
          0: { text: 'Playground' },
          1: { text: 'Web' },
          2: { text: 'Ready' },
          3: { text: '100%' },
        },
      },
    },
    cols: { len: 8 },
  },
];

function LivePreview(): ReactElement {
  const [value, setValue] = useState(previewFixture);

  return (
    <div className="tego-home-preview__sheet-host">
      <TegoSheet value={value} onChange={setValue} toolbar={false} sheetTabs={false} />
    </div>
  );
}

export function HomepagePreview(): ReactElement {
  return (
    <div className="tego-home-preview">
      <div className="tego-home-preview__bar" aria-hidden="true">
        <span />
        <span />
        <span />
        <strong>release-plan.sheet</strong>
      </div>
      <div className="tego-home-preview__viewport" aria-describedby="homepage-preview-note">
        <BrowserOnly
          fallback={<div className="tego-home-preview__loading">Loading spreadsheet preview…</div>}
        >
          {() => <LivePreview />}
        </BrowserOnly>
      </div>
      <p id="homepage-preview-note" className="tego-home-preview__note">
        Responsive preview · Full spreadsheet editing is designed for a desktop-sized workspace.
      </p>
    </div>
  );
}
