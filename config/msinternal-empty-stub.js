// Empty stub module.
// Provides a safe default export for @msinternal/* packages that are part of the
// SharePoint Online runtime CDN and are not available as npm packages.
// At runtime on SharePoint Online these modules are loaded by the SPFx AMD loader;
// this stub is only used during local webpack builds to suppress "Module not found" errors.
'use strict';
module.exports = {};
