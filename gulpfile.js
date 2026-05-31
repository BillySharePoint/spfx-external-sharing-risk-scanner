'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
    var result = getTasks.call(build.rig);
    result.set('serve', result.get('serve-deprecated'));
    return result;
};

// Fix CSS class-name mismatch: the sass task writes .scss.ts with single-hash
// names (e.g. shell_92cf71c3), but sp-css-loader's DEFAULT_GENERATE_CSS_CLASS_NAME
// appends a *second* hash to every selector in the CSS (shell_92cf71c3_aecf9af3).
// The component reads from .scss.ts (single hash), so the DOM class never matches
// the CSS rule.  Override generateCssClassName to be the identity function so the
// CSS keeps the same class names the sass task already produced.
build.configureWebpack.mergeConfig({
    additionalConfiguration: (generatedConfiguration) => {
        const rules = generatedConfiguration.module && generatedConfiguration.module.rules;
        if (rules) {
            for (const rule of rules) {
                if (rule.use && Array.isArray(rule.use)) {
                    for (const loader of rule.use) {
                        if (typeof loader === 'object' &&
                            loader.options &&
                            loader.options.generateCssClassName) {
                            loader.options.generateCssClassName = (name) => name;
                        }
                    }
                }
            }
        }
        return generatedConfiguration;
    }
});

build.initialize(require('gulp'));
