// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeAll, describe, expect, it } from "vitest";

import { resolveWorkerLoadAssets } from "./controller";
import type { XlsxWorkerLoadSnapshot } from "./worker-client";

// vitest runs from the package directory.
const WELCOME = readFileSync(
  join(process.cwd(), "../../apps/playground/public/examples/welcome.xlsx")
);

function welcomeBuffer(): ArrayBuffer {
  return WELCOME.buffer.slice(
    WELCOME.byteOffset,
    WELCOME.byteOffset + WELCOME.byteLength
  ) as ArrayBuffer;
}

// A worker snapshot as the grid receives it; its contents are opaque here,
// what matters is that they reach the grid untouched.
function snapshot(): XlsxWorkerLoadSnapshot {
  return {
    chartsByWorkbookSheetIndex: [[{ id: "chart-1" }]],
    chartsheets: [],
    formControlsByWorkbookSheetIndex: [[{ id: "checkbox-1" }]],
    sheets: [{ name: "Sheet1", styleById: { 3: { fontWeight: "bold" } } }],
    tablesByWorkbookSheetIndex: [[]],
    tabs: [{ id: "tab-0", index: 0, name: "Sheet1" }]
  } as unknown as XlsxWorkerLoadSnapshot;
}

describe("resolveWorkerLoadAssets", () => {
  // jsdom has no canvas text metrics or object URLs, both of which the drawing
  // parse uses in a browser; minimal stand-ins let it run.
  beforeAll(() => {
    HTMLCanvasElement.prototype.getContext = (() => ({
      font: "",
      measureText: () => ({ width: 7 })
    })) as unknown as HTMLCanvasElement["getContext"];
    URL.createObjectURL = () => "blob:test";
    URL.revokeObjectURL = () => undefined;
  });

  it("parses images and shapes by default", () => {
    const loaded = resolveWorkerLoadAssets({
      snapshot: snapshot(),
      buffer: welcomeBuffer(),
      skipXmlParsing: false,
      loadDrawingAssets: true
    });
    expect(loaded.imageAssets).not.toBeNull();
    expect(loaded.imageAssets?.imagesByWorkbookSheetIndex.length).toBeGreaterThan(0);
  });

  it("skips the drawing parse with loadDrawingAssets: false", () => {
    const loaded = resolveWorkerLoadAssets({
      snapshot: snapshot(),
      buffer: welcomeBuffer(),
      skipXmlParsing: false,
      loadDrawingAssets: false
    });
    expect(loaded.imageAssets).toBeNull();
  });

  it("keeps charts, form controls, cell styles, tabs and tables either way", () => {
    for (const loadDrawingAssets of [true, false]) {
      const worker = snapshot();
      const loaded = resolveWorkerLoadAssets({
        snapshot: worker,
        buffer: welcomeBuffer(),
        skipXmlParsing: false,
        loadDrawingAssets
      });
      expect(loaded.chartsByWorkbookSheetIndex).toBe(worker.chartsByWorkbookSheetIndex);
      expect(loaded.formControlsByWorkbookSheetIndex).toBe(
        worker.formControlsByWorkbookSheetIndex
      );
      expect(loaded.sheets).toBe(worker.sheets);
      expect(loaded.tabs).toBe(worker.tabs);
      expect(loaded.tablesByWorkbookSheetIndex).toBe(worker.tablesByWorkbookSheetIndex);
    }
  });

  it("never parses drawings when XML parsing is skipped", () => {
    const loaded = resolveWorkerLoadAssets({
      snapshot: snapshot(),
      buffer: welcomeBuffer(),
      skipXmlParsing: true,
      loadDrawingAssets: true
    });
    expect(loaded.imageAssets).toBeNull();
  });
});
