const assert = require("node:assert");
const { name } = require("../package.json");

assert.ok(require(name).default, `require("${name}") has no default export`);
