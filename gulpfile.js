'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Suppress all lint warnings
build.addSuppression(/Warning - \[lint\].*/);
build.addSuppression(/Warning - \[tslint\].*/);
build.addSuppression(/error @typescript-eslint.*/);
build.addSuppression(/Warning - \[webpack\].*/);
build.addSuppression(/Warning - \[sass\].*/);
build.addSuppression(/Warning - \[typescript\].*/);
build.addSuppression(/.*wrote output to stderr.*/);
build.addSuppression(/.*build failed because a task wrote output to stderr.*/);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

// Disable the lint subtask completely if it exists
if (build.tslintCmd) {
  build.tslintCmd.enabled = false;
}
if (build.eslintCmd) {
  build.eslintCmd.enabled = false;
}

build.initialize(require('gulp'));
