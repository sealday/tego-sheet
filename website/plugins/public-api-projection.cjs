// @ts-check

const callbackNames = [
  'onActiveSheetChange',
  'onCellEdit',
  'onChange',
  'onError',
  'onPaste',
  'onSelectionChange',
];
const callbackNameSet = new Set(callbackNames);
const inheritedPrefix = 'TegoSheetCallbacks.';

/**
 * Validate and flatten the private callback helper from one TegoSheetProps documentation model.
 * No mutation occurs unless the complete heritage and inherited-child shape matches.
 *
 * @param {import('typedoc').DeclarationReflection} reflection
 * @param {Pick<import('typedoc').Logger, 'error'>} logger
 * @param {string | undefined} projectPackageName
 * @returns {boolean}
 */
function projectTegoSheetProps(reflection, logger, projectPackageName) {
  const extendedTypes = reflection.extendedTypes;
  if (
    !Array.isArray(extendedTypes) ||
    extendedTypes.length !== 1 ||
    extendedTypes[0]?.type !== 'reference' ||
    extendedTypes[0].name !== 'TegoSheetCallbacks' ||
    extendedTypes[0].qualifiedName !== 'TegoSheetCallbacks' ||
    projectPackageName === undefined ||
    extendedTypes[0].package !== projectPackageName
  ) {
    logger.error(
      'public API projection expected TegoSheetProps to extend the project TegoSheetCallbacks helper',
    );
    return false;
  }

  const children = reflection.children ?? [];
  const callbackChildren = children.filter((child) => callbackNameSet.has(child.name));
  const inheritedCallbackChildren = children.filter((child) =>
    child.inheritedFrom?.name.startsWith(inheritedPrefix),
  );
  const callbackChildNames = callbackChildren.map((child) => child.name);
  const hasExactNames =
    callbackChildNames.length === callbackNames.length &&
    new Set(callbackChildNames).size === callbackNames.length &&
    callbackNames.every((name) => callbackChildNames.includes(name));
  const hasExactInheritance =
    inheritedCallbackChildren.length === callbackNames.length &&
    callbackChildren.every(
      (child) =>
        child.inheritedFrom?.type === 'reference' &&
        child.inheritedFrom.name === `${inheritedPrefix}${child.name}`,
    );
  if (!hasExactNames || !hasExactInheritance) {
    logger.error(
      'public API projection expected exactly six unique TegoSheetCallbacks inherited properties',
    );
    return false;
  }

  delete reflection.extendedTypes;
  for (const child of callbackChildren) delete child.inheritedFrom;
  return true;
}

/**
 * Flatten the private callback helper from TypeDoc's TegoSheetProps model only.
 * The compiler-facing `interface TegoSheetProps extends TegoSheetCallbacks` heritage is unchanged.
 *
 * @param {import('typedoc').Application} app
 */
function load(app) {
  app.converter.on('resolveBegin', (context) => {
    const matches = /** @type {import('typedoc').DeclarationReflection[]} */ (
      (context.project.children ?? []).filter((child) => child.name === 'TegoSheetProps')
    );
    if (matches.length !== 1) {
      app.logger.error(
        'public API projection expected exactly one direct TegoSheetProps project child',
      );
      return;
    }

    const reflection = matches[0];
    if (!reflection || !('extendedTypes' in reflection)) {
      app.logger.error('public API projection could not find TegoSheetProps interface heritage');
      return;
    }

    projectTegoSheetProps(reflection, app.logger, context.project.packageName);
  });
}

module.exports = { load, projectTegoSheetProps };
