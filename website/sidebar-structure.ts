import type { SidebarsConfig } from '@docusaurus/plugin-content-docs';
import type { SidebarItem } from '@docusaurus/plugin-content-docs/lib/sidebars/types.js';

export function createDocumentationSidebars(typedocSidebar: SidebarItem[]): SidebarsConfig {
  return {
    docsSidebar: [
      {
        type: 'category',
        label: 'Getting Started',
        items: [
          'getting-started/installation',
          'getting-started/quick-start',
          'getting-started/styling-and-sizing',
        ],
      },
      {
        type: 'category',
        label: 'Core Concepts',
        items: [
          'concepts/controlled-and-uncontrolled',
          'concepts/workbook-data',
          'concepts/refs-and-commands',
          'concepts/callbacks-and-errors',
        ],
      },
      {
        type: 'category',
        label: 'Guides',
        items: [
          'guides/custom-chrome',
          'guides/locales',
          'guides/validation-and-filtering',
          'guides/frozen-panes-and-layout',
          'guides/printing',
        ],
      },
      {
        type: 'category',
        label: 'Migration',
        items: ['migration/from-x-data-spreadsheet'],
      },
      {
        type: 'category',
        label: 'API Reference',
        link: { type: 'doc', id: 'api/index' },
        items: typedocSidebar,
      },
    ],
  };
}
