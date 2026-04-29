'use strict';

const path = require('path');

/**
 * SPFx webpack customization hook (consumed by customize-spfx-webpack-configuration-plugin).
 *
 * Problem: @microsoft/sp-property-pane (a linked SPFx external) imports several
 * @msinternal/* and @ms/* packages that are internal Microsoft packages bundled
 * only with the SharePoint Online runtime CDN — they are not available on npm.
 * During a local webpack build, source-map-loader processes sp-property-pane's
 * source files and follows those imports, producing 40+ "Module not found" errors
 * even though sp-property-pane itself is never bundled.
 *
 * Fix: use NormalModuleReplacementPlugin to redirect all @msinternal/* and the
 * affected @ms/* imports to a local empty stub module. This keeps them out of
 * webpack's externals array (avoiding SPFx ManifestPlugin validation errors)
 * while still resolving the imports successfully during a local build.
 *
 * At runtime on SharePoint Online these modules are provided by the SPFx AMD
 * loader; the stub is never shipped in the production bundle.
 */
module.exports = function customizeSpfxWebpackConfig(config, _taskSession, _heftConfiguration, webpack) {
  const stubPath = path.resolve(__dirname, 'msinternal-empty-stub.js');

  config.plugins = config.plugins || [];
  config.plugins.push(
    new webpack.NormalModuleReplacementPlugin(
      /^(@msinternal\/|@ms\/office-ui-fabric-react-bundle)/,
      stubPath
    )
  );
};
