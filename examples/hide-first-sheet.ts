import writeXlsxFile, { type Sheet } from "write-excel-file/node";
import hideSheets, { type HideSheetsSheetOptions } from "../src/index.js";

// First sheet is hidden. This is the case where Excel needs `firstSheet` and
// `activeTab` on `<workbookView/>` to know which tab to open — without it the
// file either fails to open or shows a recovery error.
const sheets: (Sheet<any> & HideSheetsSheetOptions)[] = [
  {
    sheet: "Hidden",
    hidden: true,
    data: [[{ value: "Hidden by default", fontWeight: "bold" }]],
  },
  {
    sheet: "Visible",
    data: [
      [{ value: "This is what opens", fontWeight: "bold" }],
      ["The first sheet is hidden — right-click the tab strip → Unhide to see it"],
    ],
  },
];

const output = new URL("./hide-first-sheet.xlsx", import.meta.url).pathname;

await writeXlsxFile(sheets, { features: [hideSheets] }).toFile(output);

console.log(`Wrote ${output}`);
