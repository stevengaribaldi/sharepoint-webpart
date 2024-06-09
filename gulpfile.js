/* eslint-disable no-undef */
'use strict'

// eslint-disable-next-line @typescript-eslint/no-var-requires
const build = require('@microsoft/sp-build-web')

build.addSuppression(
  `Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`
)

var getTasks = build.rig.getTasks
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig)

  result.set('serve', result.get('serve-deprecated'))

  return result
}
build.tslintCmd.enabled = false
// eslint-disable-next-line @typescript-eslint/no-var-requires
build.initialize(require('gulp'))
