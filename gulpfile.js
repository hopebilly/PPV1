"use strict";

const gulp = require('gulp');
const build = require("@microsoft/sp-build-web");

// Suppress SASS warnings about non-camelCase class names
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
build.addSuppression(`Warning - [sass] The local CSS class 'ms-List-page' is not camelCase and will not be type-safe.`);
build.addSuppression(`Warning - [sass] The local CSS class 'ms-List-cell' is not camelCase and will not be type-safe.`);
build.addSuppression(`Warning - [sass] The local CSS class 'svg-chart-container' is not camelCase and will not be type-safe.`);

// Extend Webpack config to support Babel and optional chaining
build.configureWebpack.mergeConfig({
  additionalConfiguration: function (config) {
    if (!config.module) {
      config.module = { rules: [] };
    }

    // Add Babel loader to process JavaScript (including optional chaining)
    config.module.rules.push({
      test: /\.js$/,
      exclude: /node_modules\/(?!orgchart)/,  // Allow 'orgchart' to be processed by Babel
      use: {
        loader: 'babel-loader',
        options: {
          presets: ['@babel/preset-env'],
          plugins: ['@babel/plugin-proposal-optional-chaining'],
          compact: false  // Prevent Babel from compacting large files
        }
      }
    });

     // Update css-loader configuration to preserve original (kebab-case) class names
     if (config.module && config.module.rules) {
      config.module.rules.forEach(rule => {
        if (rule.use) {
          // rule.use might be an array or a single object
          const loaders = Array.isArray(rule.use) ? rule.use : [rule.use];
          loaders.forEach(loader => {
            if (
              typeof loader === 'object' &&
              loader.loader &&
              loader.loader.includes('css-loader') &&
              loader.options &&
              loader.options.modules
            ) {
              // Preserve the original class names (kebab-case) by not converting them to camelCase
              loader.options.modules.exportLocalsConvention = 'asIs';
            }
          });
        }
      });
    }

    return config;
  }
});

// Modify tasks to use the deprecated 'serve' task (use cautiously)
var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);
  result.set("serve", result.get("serve-deprecated"));  // Ensure this is really needed
  return result;
};

// Initialize the build pipeline with Gulp
build.initialize(require("gulp"));