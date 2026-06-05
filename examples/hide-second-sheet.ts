import writeXlsxFile, { type Sheet } from "write-excel-file/node";
import hideSheets, { type HideSheetsSheetOptions } from "../src/index.js";

// Two sheets — the second is hidden. Excel users can right-click → Unhide.
const sheets: (Sheet<any> & HideSheetsSheetOptions)[] = [
  {
    sheet: "Visible",
    data: [
      [{ value: "Open this sheet", fontWeight: "bold" }],
      ["The other sheet is hidden — right-click the tab strip → Unhide"],
    ],
  },
  {
    sheet: "Hidden",
    hidden: true,
    data: [[{ value: "You found me!", fontWeight: "bold" }], ["This sheet was hidden."]],
  },
];

const output = new URL("./hide-second-sheet.xlsx", import.meta.url).pathname;

await writeXlsxFile(sheets, { features: [hideSheets] }).toFile(output);

console.log(`Wrote ${output}`);
