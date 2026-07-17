// @ts-check

/**
 * Flatten the private callback helper from TypeDoc's TegoSheetProps model only.
 * The compiler-facing `interface TegoSheetProps extends TegoSheetCallbacks` heritage is unchanged.
 *
 * @param {import('typedoc').Application} app
 */
function load(app) {
  app.converter.on('resolveBegin', (context) => {
    const reflection = /** @type {import('typedoc').DeclarationReflection | undefined} */ (
      context.project.getChildByName('TegoSheetProps')
    );
    if (!reflection || !('extendedTypes' in reflection)) {
      app.logger.error('public API projection could not find TegoSheetProps interface heritage');
      return;
    }

    const extendedTypes = reflection.extendedTypes;
    if (
      !Array.isArray(extendedTypes) ||
      extendedTypes.length !== 1 ||
      extendedTypes[0]?.type !== 'reference' ||
      extendedTypes[0].name !== 'TegoSheetCallbacks'
    ) {
      app.logger.error(
        'public API projection expected TegoSheetProps to extend only TegoSheetCallbacks',
      );
      return;
    }

    const callbackNames = [
      'onActiveSheetChange',
      'onCellEdit',
      'onChange',
      'onError',
      'onPaste',
      'onSelectionChange',
    ];
    const callbackChildren = callbackNames.map((name) =>
      reflection.children?.find((child) => child.name === name),
    );
    if (
      callbackChildren.some(
        (child, index) =>
          child?.inheritedFrom?.type !== 'reference' ||
          child.inheritedFrom.name !== `TegoSheetCallbacks.${callbackNames[index]}`,
      )
    ) {
      app.logger.error(
        'public API projection expected six TegoSheetCallbacks inherited properties',
      );
      return;
    }

    delete reflection.extendedTypes;
    for (const child of callbackChildren) {
      if (child) delete child.inheritedFrom;
    }
  });
}

module.exports = { load };
