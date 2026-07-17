import { createRequire } from 'node:module';
import { createDocumentationSidebars, parseGeneratedSidebar } from './sidebar-structure';

const require = createRequire(import.meta.url);
const typedocSidebar = parseGeneratedSidebar(require('./docs/api/typedoc-sidebar.cjs'));

const sidebars = createDocumentationSidebars(typedocSidebar);

export default sidebars;
