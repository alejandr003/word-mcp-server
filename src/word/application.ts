import { tmpdir } from "node:os"
import { join } from "node:path"
import { existsSync, copyFileSync } from "node:fs"
import type { IWordSession } from "./session.js"
import { DocumentRegistry } from "./document-registry.js"
import { WordMcpError } from "../security/errors.js"
import { normalizePath, pathsReferToSameFile } from "./types.js"

const FORMAT_EXT_MAP: Record<string, number> = {
  ".docx": 16, ".doc": 0, ".pdf": 17, ".rtf": 6, ".txt": 2,
  ".htm": 8, ".html": 8, ".mht": 9, ".mhtml": 9, ".xml": 11,
  ".odt": 23, ".dotx": 15, ".dotm": 13, ".docm": 12,
}

const FORMAT_NAMES: Record<string, number> = {
  docx: 16, doc: 0, pdf: 17, rtf: 6, txt: 2,
  htm: 8, html: 8, mht: 9, mhtml: 9, xml: 11,
  odt: 23, dotx: 15, dotm: 13, docm: 12,
}

function formatForPath(p: string): number {
  const ext = p.match(/\.(\w+)$/)?.[0]?.toLowerCase()
  return FORMAT_EXT_MAP[ext ?? ""] ?? 16
}

interface ActiveDocumentInfo {
  name: string; fullName: string; path: string; saved: boolean
}

interface OpenDocumentInfo {
  name: string; fullName: string; path: string
}

interface PathStatus {
  path: string
  existsOnDisk: boolean
  isOpenInWord: boolean
  isActive: boolean
  isTrackedByRegistry: boolean
}

interface WordStatusResult {
  wordRunning: boolean
  dialogBlocked?: boolean
  activeDocument: ActiveDocumentInfo | null
  openDocuments: OpenDocumentInfo[]
  trackedByRegistry: { name: string; path: string }[]
  pathStatus?: PathStatus
}

export class WordApplicationManager {
  public readonly registry: DocumentRegistry

  constructor(private session: IWordSession) {
    this.registry = new DocumentRegistry()
  }

  private getWord(): Record<string, unknown> {
    return this.session.application as Record<string, unknown>
  }

  /**
   * Access Word for read-only queries without launching a new WINWORD.EXE.
   * If Word is running but not attached yet, attach to it; otherwise report null.
   */
  private tryGetWordSafe(): Record<string, unknown> | null {
    try {
      if (!this.session.isAlive() && !this.session.attachIfRunning()) return null
      return this.session.application
    } catch {
      return null
    }
  }

  /** Bring a document window to the front without opening a second copy. */
  private activateDocument(doc: Record<string, unknown>): void {
    try {
      this.session.comCall(() => {
        ;(doc.Activate as (() => void) | undefined)?.()
        const win = doc.ActiveWindow as Record<string, unknown> | undefined
        ;(win?.Activate as (() => void) | undefined)?.()
      })
    } catch {
      // Activation failing must never abort the operation.
    }
  }

  private getDoc(): Record<string, unknown> {
    const doc = this.session.activeDoc ?? (() => { try { return this.getWord().ActiveDocument as Record<string, unknown> | undefined } catch { return undefined } })()
    if (!doc) throw new WordMcpError("No document is open", "NO_DOCUMENT", false, "Use word_document or word_stream_start first.")
    return doc
  }

  adoptDocument(path: string, doc: Record<string, unknown>): void {
    this.session.setActiveDoc(doc)
    this.session.setActiveDocPath(path || null)
    if (path && !this.registry.isOpen(path)) {
      this.registry.register(path, doc)
    }
  }

  isDocumentActive(): boolean {
    return this.session.activeDoc !== null
  }

  getActiveDocumentName(): string | null {
    const doc = this.session.activeDoc
    return doc ? (doc.Name as string) : null
  }

  private isPathOpenInWord(path: string): Record<string, unknown> | null {
    const word = this.tryGetWordSafe()
    if (!word) return null
    const keyPath = normalizePath(path)
    const docs = word.Documents as { Count: number; Item: (i: number) => Record<string, unknown> }
    for (let i = 1; i <= docs.Count; i++) {
      const doc = docs.Item(i) as Record<string, unknown>
      try {
        const raw = (doc.FullName as string) ?? ""
        if (!raw) continue
        // Handles OneDrive/SharePoint documents whose FullName is an https URL
        // while the caller passes the local synced path.
        if (normalizePath(raw) === keyPath || pathsReferToSameFile(path, raw)) return doc
      } catch { /* skip stale */ }
    }
    return null
  }

  async getStatus(resolvedPath?: string): Promise<WordStatusResult> {
    const word = this.tryGetWordSafe()
    if (!word) {
      return { wordRunning: false, activeDocument: null, openDocuments: [], trackedByRegistry: [] }
    }

    // 轻量 COM 探针：检测 Word 是否被弹窗阻塞（RPC_E_CALL_REJECTED）
    try {
      const _ = (word as any).DisplayAlerts
    } catch (e) {
      const code = (e as Record<string, unknown>)?.number ?? (e as Record<string, unknown>)?.code
      if (typeof code === "number" && code === 0x80010005) {
        return { wordRunning: true, dialogBlocked: true, activeDocument: null, openDocuments: [], trackedByRegistry: [] }
      }
    }

    const openDocuments: OpenDocumentInfo[] = []
    let activeDocument: ActiveDocumentInfo | null = null

    const docs = word.Documents as { Count: number; Item: (i: number) => Record<string, unknown> }
    const comActiveDoc = word.ActiveDocument as Record<string, unknown> | undefined
    const cachedActiveDoc = this.session.activeDoc

    for (let i = 1; i <= docs.Count; i++) {
      const doc = docs.Item(i) as Record<string, unknown>
      try {
        const name = doc.Name as string
        const fullName = doc.FullName as string
        const docPath = doc.Path as string
        openDocuments.push({ name, fullName, path: docPath })
        const docFullName = (doc.FullName as string) ?? ""
        const comActiveFullName = (comActiveDoc?.FullName as string) ?? ""
        if (comActiveDoc && docFullName && docFullName === comActiveFullName) {
          const saved = (doc.Saved as boolean) ?? false
          activeDocument = { name, fullName, path: docPath, saved }
        }
      } catch { /* skip stale */ }
    }

    if (!activeDocument && cachedActiveDoc && openDocuments.length > 0) {
      try {
        const name = cachedActiveDoc.Name as string
        const fullName = cachedActiveDoc.FullName as string
        activeDocument = { name, fullName, path: cachedActiveDoc.Path as string, saved: (cachedActiveDoc.Saved as boolean) ?? false }
      } catch { /* cached doc is stale */ }
    }

    this.registry.pruneStale()
    const trackedByRegistry = this.registry.listAll().map(e => ({ name: e.displayName, path: e.path }))

    let pathStatus: PathStatus | undefined
    if (resolvedPath) {
      pathStatus = {
        path: resolvedPath,
        existsOnDisk: existsSync(resolvedPath),
        isOpenInWord: this.isPathOpenInWord(resolvedPath) !== null,
        isActive: activeDocument
          ? pathsReferToSameFile(resolvedPath, activeDocument.fullName)
          : false,
        isTrackedByRegistry: this.registry.isOpen(resolvedPath),
      }
    }

    return { wordRunning: true, activeDocument, openDocuments, trackedByRegistry, pathStatus }
  }

  async ensureDocument(path?: string, title?: string): Promise<"reused" | "opened" | "created"> {
    if (path) {
      const alreadyOpen = this.isPathOpenInWord(path)
      if (alreadyOpen) {
        this.adoptDocument(path, alreadyOpen)
        this.activateDocument(alreadyOpen)
        return "reused"
      }

      if (existsSync(path)) {
        await this.openDocument(path)
        return "opened"
      }

      throw new WordMcpError(
        `File does not exist: "${path}". Use word_stream_start to create a new document, or provide a path to an existing file.`,
        "FILE_NOT_FOUND", false,
        "Check the path with word_get_status(path), or use word_stream_start to create a new document."
      )
    }

    if (this.session.activeDoc) {
      this.activateDocument(this.session.activeDoc)
      return "reused"
    }

    await this.createDocument({ title })
    return "created"
  }

  async createDocument(params?: { title?: string; author?: string }): Promise<{ name: string; fullName: string }> {
    const w = this.getWord()
    try { ;(w as any).Visible = true } catch (e) { this.session.logger?.warn({ err: e }, "Visible=true failed") }
    const doc = this.session.comCall(() => {
      const docs = w.Documents as { Add: () => Record<string, unknown> }
      return docs.Add() as Record<string, unknown>
    })
    try {
      if (params?.title) {
        const props = doc.BuiltInDocumentProperties as { Item: (n: string) => { Value: string } }
        props.Item("Title").Value = params.title
      }
      if (params?.author) {
        const props = doc.BuiltInDocumentProperties as { Item: (n: string) => { Value: string } }
        props.Item("Author").Value = params.author
      }
    } catch {
      // Suppress property errors on some Word versions
    }
    this.session.setActiveDoc(doc)
    this.session.setActiveDocPath(null)
    const fullName = doc.FullName as string
    if (fullName && fullName !== doc.Name as string) {
      this.registry.register(fullName, doc)
    }
    return { name: doc.Name as string, fullName }
  }

  async createDocumentFromTemplate(
    templatePath: string,
    params?: { title?: string; author?: string },
  ): Promise<{ name: string; fullName: string }> {
    const w = this.getWord()
    try { ;(w as any).Visible = true } catch (e) { this.session.logger?.warn({ err: e }, "Visible=true failed") }
    const doc = this.session.comCall(() => {
      const docs = w.Documents as { Add: (t: string) => Record<string, unknown> }
      return docs.Add(templatePath) as Record<string, unknown>
    })
    try {
      if (params?.title) {
        const props = doc.BuiltInDocumentProperties as { Item: (n: string) => { Value: string } }
        props.Item("Title").Value = params.title
      }
      if (params?.author) {
        const props = doc.BuiltInDocumentProperties as { Item: (n: string) => { Value: string } }
        props.Item("Author").Value = params.author
      }
    } catch {
      // Suppress property errors on some Word versions
    }
    this.session.setActiveDoc(doc)
    this.session.setActiveDocPath(null)
    return { name: doc.Name as string, fullName: doc.FullName as string }
  }

  async openDocument(path: string): Promise<{ name: string; fullName: string }> {
    // Word may already have this file open in the same instance (for example
    // OneDrive documents exposed as a URL). Reuse and focus it instead of
    // calling Documents.Open again, which creates a second window.
    const alreadyOpen = this.isPathOpenInWord(path)
    if (alreadyOpen) {
      this.adoptDocument(path, alreadyOpen)
      this.activateDocument(alreadyOpen)
      return {
        name: (alreadyOpen.Name as string) ?? path,
        fullName: (alreadyOpen.FullName as string) ?? path,
      }
    }

    const existing = this.registry.getByPath(path)
    if (existing) {
      try {
        const _ = (existing.doc as Record<string, unknown>).Name
        this.adoptDocument(path, existing.doc)
        this.activateDocument(existing.doc)
        return { name: existing.displayName, fullName: existing.path }
      } catch {
        this.registry.unregister(path)
      }
    }
    const w = this.getWord()
    try { ;(w as any).Visible = true } catch (e) { this.session.logger?.warn({ err: e }, "Visible=true failed") }
    try { ;(w as Record<string, unknown>).FileValidation = 2 } catch { /* some Word versions lack this */ }
    const doc = this.session.comCall(() => {
      const docs = w.Documents as { Open: (p: string, confirm?: boolean, readOnly?: boolean, addToRecent?: boolean) => Record<string, unknown> }
      return docs.Open(path, false, false, false) as Record<string, unknown>
    })
    try { ;(doc as Record<string, unknown>).AutoUpdate = false } catch { /* ignore */ }
    this.adoptDocument(path, doc)
    this.activateDocument(doc)
    return { name: doc.Name as string, fullName: doc.FullName as string }
  }

  async saveDocument(): Promise<void> {
    const doc = this.getDoc()
    const docPath = doc.Path as string
    if (!docPath || docPath === "") {
      const tempPath = join(tmpdir(), "word-mcp-doc.docx")
      try {
        const bakPath = tempPath + ".bak"
        if (existsSync(tempPath)) {
          copyFileSync(tempPath, bakPath)
        }
      } catch (e) { this.session.logger?.warn({ err: e }, "backup failed") }
      this.session.comCall(() => { ;(doc.SaveAs as (p: string, f: number) => void)(tempPath, 16) })
      this.session.setActiveDoc(doc)
      this.session.setActiveDocPath(tempPath)
    } else {
      const origPath = doc.FullName as string
      try {
        const bakPath = origPath + ".bak"
        if (existsSync(origPath)) {
          copyFileSync(origPath, bakPath)
        }
      } catch (e) { this.session.logger?.warn({ err: e }, "backup failed") }
      this.session.comCall(() => { ;(doc.Save as () => void)() })
      this.session.setActiveDocPath(doc.FullName as string)
    }
  }

  async saveDocumentAs(path: string, format?: unknown): Promise<{ path: string }> {
    const doc = this.getDoc()
    const f = typeof format === "string" ? (FORMAT_NAMES[format] ?? formatForPath(path)) : (format as number | undefined) ?? formatForPath(path)
    try {
      const bakPath = path + ".bak"
      if (existsSync(path)) {
        copyFileSync(path, bakPath)
      }
    } catch (e) { this.session.logger?.warn({ err: e }, "backup failed") }
    this.session.comCall(() => { ;(doc.SaveAs as (p: string, f: number) => void)(path, f) })
    this.session.setActiveDoc(doc)
    this.session.setActiveDocPath(path)
    this.registry.register(path, doc)
    return { path }
  }

  async closeDocument(saveChanges?: boolean): Promise<void> {
    const path = this.session.activeDocPath
    const prevDoc = this.session.activeDoc

    // 关闭 COM 文档；异常时仍然执行状态清理
    try {
      const doc = prevDoc ?? (() => {
        try { return this.session.comCall(() => this.getWord().ActiveDocument as Record<string, unknown> | undefined) } catch { return undefined }
      })()
      if (doc) {
        this.session.comCall(() => { ;(doc.Close as (s: boolean) => void)(saveChanges ?? false) })
      }
    } catch {
      // COM close 异常 — 继续清理内存状态
    }

    // 清理已追踪的状态
    if (path) this.registry.unregister(path)
    this.session.setActiveDoc(null)
    this.session.setActiveDocPath(null)

    // 后处理：接管剩余文档（Count > 0 时自动切换 activeDoc）
    try {
      const word = this.getWord()
      const docs = word.Documents as { Count: number; Item: (i: number) => Record<string, unknown> }
      if (docs.Count > 0) {
        const nextDoc = word.ActiveDocument as Record<string, unknown> | undefined
        if (nextDoc) {
          const nextPath = (nextDoc.FullName as string) ?? (nextDoc.Name as string) ?? ""
          this.adoptDocument(nextPath, nextDoc)
        }
      }
      // Count == 0: 保留 Word 空白窗口（原生语义，不隐藏）
    } catch { /* ignore */ }
  }

  async quit(): Promise<void> {
    this.registry.clear()
    this.session.quit()
  }

  /**
   * Detach from Word without closing documents or quitting the app. Used when
   * the MCP server shuts down — the user's open work must survive a restart.
   */
  release(): void {
    this.registry.clear()
    this.session.release()
  }
}
