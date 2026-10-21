import { vi } from "vitest"
import type { IWordSession } from "../../src/word/session.js"
import type { IDocumentProxy, ISelectionProxy, IRangeProxy } from "../../src/word/com-proxy/types.js"
import { createMockDocProxy, createMockSelProxy, createMockRangeProxy } from "../../src/word/com-proxy/com-proxy.mock.js"

export function createMockDoc(): Record<string, unknown> {
  return {
    Name: "test.docx",
    FullName: "C:\\test.docx",
    Content: { End: 0, Start: 0, Text: "" },
    Range: vi.fn(() => ({
      End: 0, Start: 0, Select: vi.fn(),
      Hyperlinks: { Add: vi.fn() },
      Shading: { BackgroundPatternColor: 0 },
    })),
    Paragraphs: { Count: 0, Item: vi.fn() },
    Tables: { Count: 0, Item: vi.fn() },
    Bookmarks: { Count: 0, Item: vi.fn(), Add: vi.fn() },
    Styles: { Count: 0, Item: vi.fn() },
    Sections: { Count: 1, Item: vi.fn() },
    Saved: true,
    ExportAsFixedFormat: vi.fn(),
    TrackRevisions: false,
    InlineShapes: { AddPicture: vi.fn() },
    Revisions: { Count: 0, AcceptAll: vi.fn(), RejectAll: vi.fn() },
    Comments: { Count: 0, Item: vi.fn() },
    Hyperlinks: { Add: vi.fn() },
    Footnotes: { Add: vi.fn() },
    Select: vi.fn(),
    Undo: vi.fn(),
    Redo: vi.fn(),
  } as unknown as Record<string, unknown>
}

export function createMockSel(): Record<string, unknown> {
  return {
    Start: 0, End: 0, StoryType: 1, Type: 1,
    TypeText: vi.fn(),
    TypeParagraph: vi.fn(),
    TypeBackspace: vi.fn(),
    Collapse: vi.fn(),
    MoveStart: vi.fn(),
    EndKey: vi.fn(),
    Style: "",
    Font: {},
    ParagraphFormat: { LeftIndent: 0, Alignment: 0 },
    Range: {
      Text: "", Font: {},
      Hyperlinks: { Add: vi.fn() },
      Shading: { BackgroundPatternColor: 0 },
      Select: vi.fn(),
      ConvertToTable: vi.fn(),
      InsertFile: vi.fn(),
      ListFormat: {
        ApplyBulletDefault: vi.fn(),
        ApplyNumberDefault: vi.fn(),
        RemoveNumbers: vi.fn(),
        IncreaseIndent: vi.fn(),
      },
    },
    InlineShapes: { AddHorizontalLineStandard: vi.fn(), AddPicture: vi.fn() },
    Find: {
      ClearFormatting: vi.fn(),
      MatchCase: false, MatchWholeWord: false,
      Style: "", Text: "", Forward: true, Wrap: 0, Format: false,
      Execute: vi.fn(),
      Replacement: {
        ClearFormatting: vi.fn(),
        ParagraphFormat: { FirstLineIndent: 0 },
        Text: "",
      },
    },
    Information: vi.fn(() => false),
    Tables: { Item: vi.fn(), Count: 0 },
    ShapeRange: { Count: 0 },
    InsertBreak: vi.fn(),
    HomeKey: vi.fn(),
    WholeStory: vi.fn(),
    Delete: vi.fn(),
    Copy: vi.fn(),
    Cut: vi.fn(),
    Paste: vi.fn(),
    GoTo: vi.fn(),
    Select: vi.fn(),
    Expand: vi.fn(),
  } as unknown as Record<string, unknown>
}

export function createMockSession(
  overrides?: Partial<IWordSession>,
  customDoc?: Record<string, unknown>,
  customSel?: Record<string, unknown>,
): IWordSession {
  const doc = customDoc ?? createMockDoc()
  const sel = customSel ?? createMockSel()
  const app = {
    Selection: sel,
    ScreenUpdating: true,
    ActiveDocument: doc,
    ScreenRefresh: vi.fn(),
  } as Record<string, unknown>

  let docProxy: IDocumentProxy = createMockDocProxy(doc)
  let selProxy: ISelectionProxy = createMockSelProxy(sel)

  const base: IWordSession = {
    application: app,
    activeDoc: doc,
    activeDocPath: null,
    setActiveDoc: vi.fn((d) => {
      docProxy = createMockDocProxy(d ?? {})
    }),
    setActiveDocPath: vi.fn(),
    ensureAlive: vi.fn(),
    isAlive: () => true,
    start: vi.fn(),
    attachIfRunning: () => true,
    release: vi.fn(),
    quit: vi.fn(),
    setOnLog: vi.fn(),
    setScreenUpdating: vi.fn(),
    healthCheck: () => true,
    recover: async () => {},
    comCall: <T>(fn: () => T) => fn(),
    withScreenOff: async <T>(fn: () => Promise<T>) => fn(),
    markHealthy: vi.fn(),
    markUnhealthy: vi.fn(),
    isUnhealthy: () => false,
    getDocProxy: () => docProxy,
    getSelectionProxy: () => selProxy,
    wrapRange: (raw: Record<string, unknown>) => createMockRangeProxy(raw),
    lockPrintView: vi.fn(),
    tryAdoptActiveDoc: vi.fn().mockReturnValue(false),
  }

  return { ...base, ...overrides }
}
