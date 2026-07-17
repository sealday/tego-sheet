// @ts-check

/** @type {Promise<typeof import('typedoc')> | undefined} */
let typedocPromise;

const loadTypeDoc = () => {
  typedocPromise ??= import('typedoc');
  return typedocPromise;
};

module.exports = {
  async bootstrap(/** @type {import('typedoc').TypeDocOptions} */ options) {
    const typedoc = await loadTypeDoc();

    return typedoc.Application.bootstrapWithPlugins(options, [
      new typedoc.TypeDocReader(),
      new typedoc.PackageJsonReader(),
      new typedoc.TSConfigReader(),
    ]);
  },
};
