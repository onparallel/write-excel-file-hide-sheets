import assert from "node:assert";
import pkg from "../package.json" with { type: "json" };

const exported = await import(pkg.name);
assert.ok(exported.default, `import("${pkg.name}") has no default export`);
