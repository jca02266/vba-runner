#!/usr/bin/env node
// Test runner using tsx CJS loader.
// Usage: node scripts/run-test.cjs <test-file.ts>
//
// Background: this project has "type":"module" in package.json, so Node.js
// defaults to the ESM loader. tsx's ESM loader cannot resolve .ts extensions
// on Node.js v20. Requiring tsx/cjs first installs the CJS require hooks,
// allowing .ts files to be require()'d directly.
require('../node_modules/tsx/dist/cjs/index.cjs');
const path = require('path');
const file = path.resolve(process.argv[2]);
require(file);
