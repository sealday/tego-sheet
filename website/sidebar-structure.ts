import type { SidebarsConfig } from '@docusaurus/plugin-content-docs';
import type {
  SidebarCategoriesShorthand,
  SidebarItemCategoryLinkConfig,
  SidebarItemConfig,
} from '@docusaurus/plugin-content-docs/lib/sidebars/types.js';

function sidebarError(path: string, expected: string): never {
  throw new TypeError(`Generated TypeDoc sidebar ${path} ${expected}.`);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function parseString(value: unknown, path: string): string {
  if (typeof value !== 'string') sidebarError(path, 'must be a string');
  return value;
}

function parseNonEmptyString(value: unknown, path: string): string {
  if (typeof value !== 'string' || value.length === 0)
    sidebarError(path, 'must be a non-empty string');
  return value;
}

function parseBoolean(value: unknown, path: string): boolean {
  if (typeof value !== 'boolean') sidebarError(path, 'must be a boolean');
  return value;
}

function assertAllowedFields(
  value: Record<string, unknown>,
  path: string,
  allowedFields: readonly string[],
): void {
  const allowed = new Set(allowedFields);
  for (const field of Object.keys(value)) {
    if (!allowed.has(field)) sidebarError(`${path}.${field}`, 'is not a supported field');
  }
}

interface SidebarItemBaseConfig {
  key?: string;
  className?: string;
  customProps?: Record<string, unknown>;
}

const baseFields = ['key', 'className', 'customProps'] as const;

function parseBaseFields(value: Record<string, unknown>, path: string): SidebarItemBaseConfig {
  const parsed: SidebarItemBaseConfig = {};
  if (value.key !== undefined) parsed.key = parseString(value.key, `${path}.key`);
  if (value.className !== undefined)
    parsed.className = parseString(value.className, `${path}.className`);
  if (value.customProps !== undefined) {
    if (!isRecord(value.customProps)) sidebarError(`${path}.customProps`, 'must be an object');
    parsed.customProps = value.customProps;
  }
  return parsed;
}

function parseOptionalString(value: unknown, path: string): string | undefined {
  return value === undefined ? undefined : parseString(value, path);
}

function parseOptionalBoolean(value: unknown, path: string): boolean | undefined {
  return value === undefined ? undefined : parseBoolean(value, path);
}

function parseCategoryLink(value: unknown, path: string): SidebarItemCategoryLinkConfig {
  if (!isRecord(value)) sidebarError(path, 'must be a category link object');
  const type = parseNonEmptyString(value.type, `${path}.type`);
  if (type === 'doc') {
    assertAllowedFields(value, path, ['type', 'id']);
    return { type, id: parseNonEmptyString(value.id, `${path}.id`) };
  }
  if (type !== 'generated-index') {
    sidebarError(`${path}.type`, `has unsupported value ${JSON.stringify(type)}`);
  }
  assertAllowedFields(value, path, ['type', 'slug', 'title', 'description', 'image', 'keywords']);
  let keywords: string | readonly string[] | undefined;
  if (value.keywords !== undefined) {
    if (typeof value.keywords === 'string') {
      keywords = value.keywords;
    } else if (
      Array.isArray(value.keywords) &&
      value.keywords.every((keyword) => typeof keyword === 'string')
    ) {
      keywords = [...value.keywords];
    } else {
      sidebarError(`${path}.keywords`, 'must be a string or an array of strings');
    }
  }
  const slug = parseOptionalString(value.slug, `${path}.slug`);
  const title = parseOptionalString(value.title, `${path}.title`);
  const description = parseOptionalString(value.description, `${path}.description`);
  const image = parseOptionalString(value.image, `${path}.image`);
  return {
    type,
    ...(slug === undefined ? {} : { slug }),
    ...(title === undefined ? {} : { title }),
    ...(description === undefined ? {} : { description }),
    ...(image === undefined ? {} : { image }),
    ...(keywords === undefined ? {} : { keywords }),
  };
}

function parseSidebarShorthand(
  value: Record<string, unknown>,
  path: string,
): SidebarCategoriesShorthand {
  const parsed: SidebarCategoriesShorthand = {};
  for (const [label, items] of Object.entries(value)) {
    parseNonEmptyString(label, `${path} category label`);
    parsed[label] = parseSidebarConfig(items, `${path}.${label}`);
  }
  return parsed;
}

function parseSidebarConfig(
  value: unknown,
  path: string,
): SidebarCategoriesShorthand | SidebarItemConfig[] {
  if (Array.isArray(value)) {
    return value.map((item, index) => parseSidebarItem(item, `${path}[${index}]`));
  }
  if (!isRecord(value)) sidebarError(path, 'must be an array or category shorthand object');
  return parseSidebarShorthand(value, path);
}

function parseSidebarItem(value: unknown, path: string): SidebarItemConfig {
  if (typeof value === 'string') return value;
  if (!isRecord(value)) sidebarError(path, 'must be a string or sidebar item object');
  if (value.type === undefined || Array.isArray(value.type) || isRecord(value.type)) {
    return parseSidebarShorthand(value, path);
  }

  const type = parseNonEmptyString(value.type, `${path}.type`);
  switch (type) {
    case 'doc':
    case 'ref': {
      assertAllowedFields(value, path, ['type', 'id', 'label', ...baseFields]);
      const label =
        value.label === undefined ? undefined : parseNonEmptyString(value.label, `${path}.label`);
      return {
        ...parseBaseFields(value, path),
        type,
        id: parseNonEmptyString(value.id, `${path}.id`),
        ...(label === undefined ? {} : { label }),
      };
    }
    case 'category': {
      assertAllowedFields(value, path, [
        'type',
        'label',
        'items',
        'collapsed',
        'collapsible',
        'description',
        'link',
        ...baseFields,
      ]);
      const collapsed = parseOptionalBoolean(value.collapsed, `${path}.collapsed`);
      const collapsible = parseOptionalBoolean(value.collapsible, `${path}.collapsible`);
      const description = parseOptionalString(value.description, `${path}.description`);
      const link =
        value.link === undefined ? undefined : parseCategoryLink(value.link, `${path}.link`);
      return {
        ...parseBaseFields(value, path),
        type,
        label: parseNonEmptyString(value.label, `${path}.label`),
        items: parseSidebarConfig(value.items, `${path}.items`),
        ...(collapsed === undefined ? {} : { collapsed }),
        ...(collapsible === undefined ? {} : { collapsible }),
        ...(description === undefined ? {} : { description }),
        ...(link === undefined ? {} : { link }),
      };
    }
    case 'autogenerated': {
      assertAllowedFields(value, path, ['type', 'dirName', ...baseFields]);
      return {
        ...parseBaseFields(value, path),
        type,
        dirName: parseNonEmptyString(value.dirName, `${path}.dirName`),
      };
    }
    case 'html': {
      assertAllowedFields(value, path, ['type', 'value', 'defaultStyle', ...baseFields]);
      const defaultStyle = parseOptionalBoolean(value.defaultStyle, `${path}.defaultStyle`);
      return {
        ...parseBaseFields(value, path),
        type,
        value: parseString(value.value, `${path}.value`),
        ...(defaultStyle === undefined ? {} : { defaultStyle }),
      };
    }
    case 'link': {
      assertAllowedFields(value, path, [
        'type',
        'href',
        'label',
        'autoAddBaseUrl',
        'description',
        ...baseFields,
      ]);
      const autoAddBaseUrl = parseOptionalBoolean(value.autoAddBaseUrl, `${path}.autoAddBaseUrl`);
      const description = parseOptionalString(value.description, `${path}.description`);
      return {
        ...parseBaseFields(value, path),
        type,
        href: parseNonEmptyString(value.href, `${path}.href`),
        label: parseNonEmptyString(value.label, `${path}.label`),
        ...(autoAddBaseUrl === undefined ? {} : { autoAddBaseUrl }),
        ...(description === undefined ? {} : { description }),
      };
    }
    default:
      sidebarError(`${path}.type`, `has unsupported value ${JSON.stringify(type)}`);
  }
}

export function parseGeneratedSidebar(value: unknown): SidebarItemConfig[] {
  if (!Array.isArray(value)) sidebarError('export', 'must be an array');
  return value.map((item, index) => parseSidebarItem(item, `export[${index}]`));
}

export function createDocumentationSidebars(typedocSidebar: SidebarItemConfig[]): SidebarsConfig {
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
        label: 'Product Roadmap',
        link: { type: 'doc', id: 'roadmap/index' },
        items: [
          'roadmap/foundation',
          'roadmap/template-printing',
          'roadmap/formulas-data',
          'roadmap/analysis-visualization',
          'roadmap/extensibility',
          'roadmap/host-integrations',
        ],
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
