import {
  findElement,
  findElementInsideElement,
  getChildElements,
  replaceElement,
  getSelfClosingTagMarkup,
} from "write-excel-file/utility";
import type { Feature } from "write-excel-file/node";

import type { HideSheetsSheetOptions } from "./types.js";

// `Feature<any>` rather than `Feature<unknown>` so the feature is assignable to
// `Feature<FileContent>` for any `FileContent` (Buffer/Blob/etc) chosen by the caller.
// This feature only patches `xl/workbook.xml` (a string), so `FileContent` is irrelevant.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const hideSheets: Feature<any> = {
  files: {
    write: {
      files: (sheetsOptions, { read }) => {
        const hiddenFlags = (sheetsOptions as readonly HideSheetsSheetOptions[]).map((o) =>
          Boolean(o.hidden),
        );

        // Nothing to do if no sheet is marked hidden.
        if (!hiddenFlags.some((h) => h)) return undefined;

        // Excel requires at least one visible sheet. If the caller marked every sheet
        // hidden, leave the workbook untouched rather than producing a file Excel will
        // refuse to open.
        const firstVisibleSheetIndex = hiddenFlags.indexOf(false);
        if (firstVisibleSheetIndex === -1) return undefined;

        const workbookXml = read("xl/workbook.xml");
        if (typeof workbookXml !== "string") return undefined;

        return {
          "xl/workbook.xml": patchWorkbookXml(workbookXml, hiddenFlags, firstVisibleSheetIndex),
        };
      },
    },
  },
};

export default hideSheets;

function patchWorkbookXml(
  xml: string,
  hiddenFlags: readonly boolean[],
  firstVisibleSheetIndex: number,
): string {
  // 1. Mark each `<sheet/>` whose flag is true with `state="hidden"`.
  // Iterate in reverse so each replaceElement doesn't shift the offsets of
  // the earlier (still-cached) elements.
  const sheetsElement = findElement(xml, "sheets");
  if (sheetsElement) {
    const sheetElements = getChildElements(xml, sheetsElement);
    for (let i = sheetElements.length - 1; i >= 0; i--) {
      if (!hiddenFlags[i]) continue;
      const el = sheetElements[i];
      if (!el || el.tagName !== "sheet") continue;
      xml = replaceElement(
        xml,
        el,
        getSelfClosingTagMarkup("sheet", {
          ...el.openingTagAttributes,
          state: "hidden",
        }),
      );
    }
  }

  // 2. When the first sheet is hidden, point `<workbookView/>` at the first
  // visible sheet. Without `activeTab`, Excel opens to sheet 0 — which is
  // hidden — and either refuses the file or shows a recovery error.
  if (firstVisibleSheetIndex > 0) {
    const bookViewsElement = findElement(xml, "bookViews");
    if (bookViewsElement) {
      const workbookViewElement = findElementInsideElement(xml, "workbookView", bookViewsElement);
      if (workbookViewElement) {
        xml = replaceElement(
          xml,
          workbookViewElement,
          getSelfClosingTagMarkup("workbookView", {
            ...workbookViewElement.openingTagAttributes,
            firstSheet: firstVisibleSheetIndex,
            activeTab: firstVisibleSheetIndex,
          }),
        );
      }
    }
  }

  return xml;
}
