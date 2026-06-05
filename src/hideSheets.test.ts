import { describe, it, expect } from "vitest";

import hideSheets from "./hideSheets.js";
import type { HideSheetsSheetOptions } from "./types.js";

const writeFiles = hideSheets.files!.write!.files!;

// Realistic shape of `xl/workbook.xml` as emitted by `write-excel-file`. The
// xmlns declarations are present on `<workbook>` but stripped in fixtures here
// for readability — `findElement` matches on local name only.
function workbookXml(sheets: Array<{ name: string; sheetId: number }>): string {
  const sheetTags = sheets
    .map((s) => `<sheet name="${s.name}" sheetId="${s.sheetId}" r:id="rId${s.sheetId}"/>`)
    .join("");
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    "<workbook>" +
    "<workbookPr/>" +
    "<bookViews><workbookView/></bookViews>" +
    "<sheets>" +
    sheetTags +
    "</sheets>" +
    "<definedNames/>" +
    "<calcPr/>" +
    "</workbook>"
  );
}

function run(
  sheetsOptions: HideSheetsSheetOptions[],
  xml: string,
): Record<string, string> | undefined {
  const result = writeFiles(sheetsOptions as never, {
    read: (path: string) => (path === "xl/workbook.xml" ? xml : undefined),
  });
  return result as Record<string, string> | undefined;
}

describe("hideSheets — pass-through", () => {
  it("returns undefined when no sheet is hidden", () => {
    const xml = workbookXml([
      { name: "A", sheetId: 1 },
      { name: "B", sheetId: 2 },
    ]);
    expect(run([{}, {}], xml)).toBeUndefined();
  });

  it("returns undefined when hidden is explicitly false", () => {
    const xml = workbookXml([{ name: "A", sheetId: 1 }]);
    expect(run([{ hidden: false }], xml)).toBeUndefined();
  });

  it("returns undefined when ALL sheets are hidden (refuses to corrupt the file)", () => {
    const xml = workbookXml([
      { name: "A", sheetId: 1 },
      { name: "B", sheetId: 2 },
    ]);
    expect(run([{ hidden: true }, { hidden: true }], xml)).toBeUndefined();
  });
});

describe("hideSheets — hidden non-first sheet", () => {
  it('adds state="hidden" only to the targeted <sheet/>; leaves <workbookView/> alone', () => {
    const xml = workbookXml([
      { name: "Visible", sheetId: 1 },
      { name: "Hidden", sheetId: 2 },
    ]);
    const out = run([{}, { hidden: true }], xml);
    expect(out).toBeDefined();
    const patched = out!["xl/workbook.xml"]!;

    // Sheet 1 untouched
    expect(patched).toContain('<sheet name="Visible" sheetId="1" r:id="rId1"/>');
    // Sheet 2 gets state="hidden", existing attributes preserved
    expect(patched).toContain('<sheet name="Hidden" sheetId="2" r:id="rId2" state="hidden"/>');
    // <workbookView/> NOT modified — first sheet is still visible
    expect(patched).toContain("<bookViews><workbookView/></bookViews>");
    expect(patched).not.toContain("firstSheet=");
    expect(patched).not.toContain("activeTab=");
  });

  it("preserves the third sheet when only the middle one is hidden", () => {
    const xml = workbookXml([
      { name: "A", sheetId: 1 },
      { name: "B", sheetId: 2 },
      { name: "C", sheetId: 3 },
    ]);
    const out = run([{}, { hidden: true }, {}], xml)!;
    const patched = out["xl/workbook.xml"]!;
    expect(patched).toContain('<sheet name="A" sheetId="1" r:id="rId1"/>');
    expect(patched).toContain('<sheet name="B" sheetId="2" r:id="rId2" state="hidden"/>');
    expect(patched).toContain('<sheet name="C" sheetId="3" r:id="rId3"/>');
  });
});

describe("hideSheets — hidden first sheet", () => {
  it("marks the first <sheet/> hidden AND points <workbookView/> at the first visible one", () => {
    const xml = workbookXml([
      { name: "Hidden", sheetId: 1 },
      { name: "Visible", sheetId: 2 },
    ]);
    const out = run([{ hidden: true }, {}], xml)!;
    const patched = out["xl/workbook.xml"]!;

    expect(patched).toContain('<sheet name="Hidden" sheetId="1" r:id="rId1" state="hidden"/>');
    expect(patched).toContain('<sheet name="Visible" sheetId="2" r:id="rId2"/>');
    expect(patched).toContain('<workbookView firstSheet="1" activeTab="1"/>');
  });

  it("points <workbookView/> at sheet index 2 when the first two are hidden", () => {
    const xml = workbookXml([
      { name: "H1", sheetId: 1 },
      { name: "H2", sheetId: 2 },
      { name: "V", sheetId: 3 },
    ]);
    const out = run([{ hidden: true }, { hidden: true }, {}], xml)!;
    const patched = out["xl/workbook.xml"]!;

    expect(patched).toContain('<sheet name="H1" sheetId="1" r:id="rId1" state="hidden"/>');
    expect(patched).toContain('<sheet name="H2" sheetId="2" r:id="rId2" state="hidden"/>');
    expect(patched).toContain('<sheet name="V" sheetId="3" r:id="rId3"/>');
    expect(patched).toContain('<workbookView firstSheet="2" activeTab="2"/>');
  });
});

describe("hideSheets — attribute preservation", () => {
  it("preserves pre-existing attributes on <workbookView/>", () => {
    const xml =
      "<workbook>" +
      '<bookViews><workbookView xWindow="0" yWindow="0"/></bookViews>' +
      "<sheets>" +
      '<sheet name="A" sheetId="1" r:id="rId1"/>' +
      '<sheet name="B" sheetId="2" r:id="rId2"/>' +
      "</sheets>" +
      "</workbook>";
    const out = run([{ hidden: true }, {}], xml)!;
    const patched = out["xl/workbook.xml"]!;
    expect(patched).toContain(
      '<workbookView xWindow="0" yWindow="0" firstSheet="1" activeTab="1"/>',
    );
  });

  it("preserves the surrounding XML structure (header, definedNames, etc.)", () => {
    const xml = workbookXml([
      { name: "A", sheetId: 1 },
      { name: "B", sheetId: 2 },
    ]);
    const out = run([{}, { hidden: true }], xml)!;
    const patched = out["xl/workbook.xml"]!;
    expect(patched.startsWith('<?xml version="1.0"')).toBe(true);
    expect(patched).toContain("<workbookPr/>");
    expect(patched).toContain("<definedNames/>");
    expect(patched).toContain("<calcPr/>");
    expect(patched.endsWith("</workbook>")).toBe(true);
  });
});

describe("hideSheets — defensive", () => {
  it("returns undefined when read() cannot find xl/workbook.xml", () => {
    const result = writeFiles([{ hidden: true } as never], {
      read: () => undefined,
    });
    expect(result).toBeUndefined();
  });
});
