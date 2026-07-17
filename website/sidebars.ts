import { createRequire } from 'node:module';
import type { SidebarsConfig } from '@docusaurus/plugin-content-docs';
import type { SidebarItem } from '@docusaurus/plugin-content-docs/lib/sidebars/types.js';

const require = createRequire(import.meta.url);
const typedocSidebar = require('./docs/api/typedoc-sidebar.cjs') as SidebarItem[];

const sidebars: SidebarsConfig = {
  docsSidebar: [
    {
      type: 'category',
      label: 'Getting Started',
      items: ['getting-started/installation', 'getting-started/quick-start'],
    },
    {
      type: 'category',
      label: 'Core Concepts',
      link: { type: 'generated-index', slug: '/concepts' },
      items: [],
    },
    {
      type: 'category',
      label: 'Guides',
      link: { type: 'generated-index', slug: '/guides' },
      items: [],
    },
    {
      type: 'category',
      label: 'Migration',
      link: { type: 'generated-index', slug: '/migration' },
      items: [],
    },
    {
      type: 'category',
      label: 'API Reference',
      link: { type: 'doc', id: 'api/index' },
      items: typedocSidebar,
    },
  ],
};

export default sidebars;
