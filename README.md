# @onparallel/write-excel-file-hide-sheets

Hide-sheets custom feature for [`write-excel-file`](https://www.npmjs.com/package/write-excel-file).

Adds support for marking entire sheets as **hidden** in the generated `.xlsx`, without requiring a fork of the base package. Hidden sheets are preserved in the file (Excel users can right-click the tab strip → "Unhide" to reveal them); they are not deleted.

## Install

```sh
npm install write-excel-file @onparallel/write-excel-file-hide-sheets
```

`write-excel-file` is a peer dependency (`^4.0.0`). The package is published publicly under the `@onparallel` scope — no authentication required to install.

## Usage

Register the feature when calling `writeXlsxFile()`. `write-excel-file/node`'s built-in `Sheet` type does not know about `hidden`, so intersect it with the `HideSheetsSheetOptions` type exported by this package:

```ts
import writeXlsxFile, { type Sheet } from "write-excel-file/node";
import hideSheets, { type HideSheetsSheetOptions } from "@onparallel/write-excel-file-hide-sheets";

const sheets: (Sheet<any> & HideSheetsSheetOptions)[] = [
  { sheet: "Visible", data: [["shown"]] },
  { sheet: "Hidden", data: [["not shown"]], hidden: true },
];

await writeXlsxFile(sheets, { features: [hideSheets] }).toFile("out.xlsx");
```

If you use the intersection in many places, alias it locally: `type MySheet = Sheet<any> & HideSheetsSheetOptions`.

Why not module augmentation? `write-excel-file/node` re-exports its types with `export type { ... }`, which TypeScript does not allow downstream packages to augment.

## API

```ts
export interface HideSheetsSheetOptions {
  hidden?: boolean;
}
```

Set `hidden: true` on any sheet you want hidden in the resulting workbook. The default is `false` (visible).

## Limitations

- **At least one sheet must remain visible.** Excel requires this — a workbook with every sheet hidden either fails to open or shows a recovery error. If the caller marks _every_ sheet as hidden, this feature does nothing rather than producing a broken file. The base library's output (all sheets visible) is kept as a safe fallback. Validate caller input upstream if you want to surface a clearer error.

## How it works

The feature subscribes to the `files.write.files` hook, which `write-excel-file` invokes after all worksheet XMLs have been generated. The hook:

1. Reads `xl/workbook.xml` via the `read` callback the library provides.
2. Adds `state="hidden"` to each `<sheet/>` element whose corresponding `sheetOptions.hidden` is truthy.
3. When the first sheet is hidden, rewrites `<workbookView/>` with `firstSheet="N" activeTab="N"` pointing at the first visible sheet's index. Without this, Excel opens the file targeting a hidden tab and refuses to render it.
4. Returns the patched XML, which the library writes back over the original.

This stateless, single-hook approach avoids any closure or factory state across calls.

See [`src/hideSheets.ts`](./src/hideSheets.ts) for the implementation.

## Examples

The [`examples/`](./examples) folder contains runnable TypeScript scripts that emit `.xlsx` files next to themselves. Open the generated files in Excel/Numbers/LibreOffice to confirm the hidden sheets behave as expected.

```sh
npm install
npx tsx examples/hide-second-sheet.ts   # run a single example
npm run examples                          # run them all
```

See [`examples/README.md`](./examples/README.md) for the full list.

## Development

```sh
npm install
npm test          # vitest
npm run typecheck # tsc --noEmit
npm run build     # emit dist/ via tsc
```

## License

[MIT](./LICENSE)
