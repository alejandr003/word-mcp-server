import { describe, it, expect, vi } from "vitest"
import { mkdtempSync, writeFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { WordApplicationManager } from "../../../src/word/application.js"
import { createMockSession } from "../test-helpers.js"
import type { IWordSession } from "../../../src/word/session.js"

const LOCAL_PATH =
  "C:\\Users\\aleja\\OneDrive - Universidad\\Documentos\\TI91\\Unit 2 Evaluation.docx"
const SHAREPOINT_URL =
  "https://universi799-my.sharepoint.com/personal/aleja/Documents/Documentos/TI91/Unit%202%20Evaluation.docx"

function makeDoc(fullName: string, name = fullName.split(/[\\/]/).pop() ?? "doc.docx") {
  return {
    Name: name,
    FullName: fullName,
    Path: fullName.replace(/[\\/][^\\/]+$/, ""),
    Saved: true,
    Activate: vi.fn(),
    ActiveWindow: { Activate: vi.fn() },
  } as unknown as Record<string, unknown>
}

function makeApp(docs: Record<string, unknown>[]) {
  const open = vi.fn((path: string) => {
    const doc = makeDoc(path)
    docs.push(doc)
    return doc
  })
  const add = vi.fn(() => {
    const doc = makeDoc("Document1", "Document1")
    docs.push(doc)
    return doc
  })
  const app: Record<string, unknown> = {
    Visible: true,
    ActiveDocument: docs[0],
    Documents: {
      get Count() { return docs.length },
      Item: (i: number) => docs[i - 1],
      Open: open,
      Add: add,
    },
  }
  return { app, open, add }
}

function makeSession(app: Record<string, unknown>, overrides?: Partial<IWordSession>) {
  return createMockSession({
    application: app,
    activeDoc: null,
    isAlive: () => true,
    attachIfRunning: () => true,
    ...overrides,
  })
}

describe("WordApplicationManager", () => {
  it("reuses and activates an already-open document instead of opening a duplicate", async () => {
    const doc = makeDoc(LOCAL_PATH)
    const { app, open } = makeApp([doc])
    const session = makeSession(app)
    const manager = new WordApplicationManager(session)

    const result = await manager.ensureDocument(LOCAL_PATH)

    expect(result).toBe("reused")
    expect(open).not.toHaveBeenCalled()
    expect(doc.Activate).toHaveBeenCalled()
    expect(session.setActiveDoc).toHaveBeenCalledWith(doc)
  })

  it("matches a local OneDrive path to an open SharePoint URL document", async () => {
    const doc = makeDoc(SHAREPOINT_URL)
    const { app, open } = makeApp([doc])
    const session = makeSession(app)
    const manager = new WordApplicationManager(session)

    const result = await manager.ensureDocument(LOCAL_PATH)

    expect(result).toBe("reused")
    expect(open).not.toHaveBeenCalled()
    expect(doc.Activate).toHaveBeenCalled()
  })

  it("opens the file when it is not already open", async () => {
    const dir = mkdtempSync(join(tmpdir(), "word-mcp-test-"))
    const filePath = join(dir, "Unit 2 Evaluation.docx")
    writeFileSync(filePath, "fake docx")
    try {
      const { app, open } = makeApp([])
      const session = makeSession(app)
      const manager = new WordApplicationManager(session)

      const result = await manager.ensureDocument(filePath)

      expect(result).toBe("opened")
      expect(open).toHaveBeenCalledWith(filePath, false, false, false)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it("creates an untitled document with the requested title", async () => {
    const { app, add } = makeApp([])
    const session = makeSession(app)
    const manager = new WordApplicationManager(session)

    const result = await manager.ensureDocument(undefined, "My report")

    expect(result).toBe("created")
    expect(add).toHaveBeenCalled()
    expect(session.setActiveDoc).toHaveBeenCalled()
  })

  it("getStatus does not create Word when it is not running", async () => {
    const session = createMockSession({
      isAlive: () => false,
      attachIfRunning: () => false,
      application: undefined as unknown as Record<string, unknown>,
    })
    const manager = new WordApplicationManager(session)

    const status = await manager.getStatus()

    expect(status.wordRunning).toBe(false)
    expect(status.openDocuments).toEqual([])
  })

  it("getStatus reports an open document after attaching to the user's Word", async () => {
    const doc = makeDoc(LOCAL_PATH)
    const { app } = makeApp([doc])
    const session = makeSession(app)
    const manager = new WordApplicationManager(session)

    const status = await manager.getStatus(LOCAL_PATH)

    expect(status.wordRunning).toBe(true)
    expect(status.activeDocument?.name).toBe("Unit 2 Evaluation.docx")
    expect(status.pathStatus?.isOpenInWord).toBe(true)
    expect(status.pathStatus?.isActive).toBe(true)
  })
})
