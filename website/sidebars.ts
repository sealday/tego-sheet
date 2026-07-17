import { createRequire } from 'node:module';
import type { SidebarItem } from '@docusaurus/plugin-content-docs/lib/sidebars/types.js';
import { createDocumentationSidebars } from './sidebar-structure';

const require = createRequire(import.meta.url);
const typedocSidebar = require('./docs/api/typedoc-sidebar.cjs') as SidebarItem[];

const sidebars = createDocumentationSidebars(typedocSidebar);

export default sidebars;
