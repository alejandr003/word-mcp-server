import { z } from "zod"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { WordApplicationManager } from "../../word/application.js"
import { WordDocument } from "../../word/document.js"
import { SecurityManager } from "../../security/policy.js"
import type { ServerContext } from "../server-context.js"
import { createRegTool } from "./shared.js"

export function registerDocumentTools(
  server: McpServer,
  context: ServerContext,
  appManager: WordApplicationManager,
  docOps: WordDocument,
  security: SecurityManager,
): void {
  const regTool = createRegTool(server, security, context)
  regTool("word_get_status",
    {
      description: "WHEN: always call first to understand what state Word is in before deciding the next action. WHAT: returns current state: NO_WORD (Word not running), NO_DOC (Word running, no document), DOC_ACTIVE (document open with details), or DIALOG (Word blocked by modal dialog). CONSTRAINT: read-only; zero side effects on Word state. Essential before any operation to avoid errors.",
      inputSchema: {
        path: z.string().min(1).max(4096).optional().describe("Optional file path to check if a specific file is open, tracked, and active"),
      },
    },
    async ({ path }) => {
      const safePath = path ? security.pathSanitizer.resolveAndValidate(path) : undefined
      const status = await appManager.getStatus(safePath)

      if (!status.wordRunning) {
        return "State: NO_WORD — Word is not running."
      }

      if (status.dialogBlocked) {
        return "State: DIALOG — Word is blocked by a dialog (macro security, file-in-use, etc.). Close the dialog in the Word window to continue, or wait for auto-recovery."
      }

      if (safePath) {
        const active = status.activeDocument
        const p = status.pathStatus!
        const activeDocLine = active
          ? `Active: "${active.name}"`
          : "Active: NONE"
        const targetLine = `Target: ${p.path} — Exists: ${p.existsOnDisk} — Open: ${p.isOpenInWord} — Active: ${p.isActive}`
        let actionLine = ""
        if (p.existsOnDisk) {
          actionLine = p.isOpenInWord
            ? `→ Already open. Use word_document("${p.path}") to switch.`
            : `→ On disk. Use word_document("${p.path}") to open.`
        } else {
          actionLine = "→ Not on disk. Use word_stream_start to create."
        }
        return `State: DOC_ACTIVE — ${activeDocLine}\n${targetLine}\n${actionLine}`
      }

      if (status.activeDocument) {
        const d = status.activeDocument
        return `State: DOC_ACTIVE — Active: "${d.name}" — ${status.openDocuments.length} document(s) open — Saved: ${d.saved}`
      }
      return "State: NO_DOC — Word is running, no active document."
    },
    { preconditions: [] },
  )

  regTool("word_document",
    {
      description: "WHEN: need to switch to an already-open document, open an existing file, or create an untitled document. WHAT: if path is provided and file exists, opens it; if already open (including OneDrive/SharePoint copies), activates that window instead of opening a duplicate. If no path, creates/reuses an untitled document. CONSTRAINT: preferred over word_open. Error if path does not exist on disk.",
      inputSchema: {
        path: z.string().min(1).max(4096).optional().describe("Path to an existing .docx file. Omit to create an untitled document or reuse the active one."),
        title: z.string().max(255).optional().describe("Document title shown in Word title bar (used only when creating untitled doc with no path)"),
      },
    },
    async ({ path, title }) => {
      const safePath = path ? security.pathSanitizer.resolveAndValidate(path) : undefined
      const action = await appManager.ensureDocument(safePath, title)
      context.director?.enterEditMode()
      const doc = await appManager.getStatus()
      const name = doc.activeDocument?.name ?? "untitled"
      const labels: Record<string, string> = { created: "Created", opened: "Opened", reused: "Reused" }
      return `Action: ${labels[action]} "${name}"\nNext: word_get_info() or word_get_structure() or word_insert_at(...)`
    },
    { preconditions: [] },
  )

  regTool("word_open",
    {
      description: "WHEN: need to open a .docx file from disk. WHAT: opens the file, or activates it if that document is already open in Word (no duplicate window). CONSTRAINT: file must exist on disk. word_document is the preferred general-purpose tool; word_open is equivalent but focused on opening from disk.",
      inputSchema: {
        path: z.string().min(1).max(4096).describe("Full path to the .docx file to open in a new Word window"),
      },
    },
    async ({ path }) => {
      const safePath = security.pathSanitizer.validateForRead(path)
      const result = await appManager.openDocument(safePath)
      context.director?.enterEditMode()
      return `Action: Document opened "${result.name}"\nDetail: ${result.fullName}\nNext: word_get_info() or word_get_structure() or word_insert_at(...)`
    },
  )

  regTool("word_save",
    {
      description: "WHEN: want to persist changes to the current file. WHAT: saves the active document, creating a .bak backup before overwriting. CONSTRAINT: undo history is cleared after save. Cannot undo past a save boundary. For saving to a new path or format, use word_save_as.",
    },
    async () => {
      await appManager.saveDocument()
      return "Action: Document saved\nNext: word_close({saveChanges: true}) or word_save_as({path:\"C:\\output.docx\"})"
    },
  )

  regTool("word_save_as",
    {
      description: "WHEN: need to create a copy, export to different format (PDF, RTF, TXT, HTML), or save to a new location. WHAT: saves the current document with a new path/format, then switches to the new file. CONSTRAINT: format is auto-detected from file extension. For PDF export without changing the active document, use word_export_to_pdf.",
      inputSchema: {
        path: z.string().min(1).max(4096).describe("Full save path (e.g., 'C:\\output\\report.pdf' or 'C:\\output\\copy.docx')"),
        format: z.enum(["docx", "doc", "pdf", "rtf", "txt", "html", "mht", "xml", "odt", "dotx", "dotm", "docm"]).optional().describe("File format (default: determined from file extension). Use 'pdf' for PDF, 'txt' for plain text."),
      },
    },
    async ({ path, format }) => {
      const safePath = security.pathSanitizer.validateForWrite(path)
      const result = await appManager.saveDocumentAs(safePath, format)
      return `Action: Document saved to ${result.path}\nNext: word_close({saveChanges: false})`
    },
  )

  regTool("word_close",
    {
      description: "WHEN: done editing and want to return to NO_DOC state or switch to a different document workflow. WHAT: closes the active document (does NOT close Word). CONSTRAINT: if saveChanges=false, unsaved changes may be lost. After close, use word_document to open another file or word_stream_start to create new.",
      inputSchema: {
        saveChanges: z.boolean().optional().describe("Whether to save changes before closing (default: false). Set true to avoid losing edits."),
      },
    },
    async ({ saveChanges }) => {
      await appManager.closeDocument(saveChanges)
      context.director?.exitEditMode()
      context.director?.releaseStreamLock()
      const saved = saveChanges ?? false
      return `Action: Document closed\nDetail: Changes saved: ${saved}\nNext: word_document({path:"C:\\file.docx"}) or word_stream_start({title:"New Doc"})`
    },
  )

  regTool("word_get_info",
    {
      description: "WHEN: need to check document size, word count, page count, table count, or save status. WHAT: returns document statistics (words, paragraphs, pages, characters, sections, saved state). CONSTRAINT: read-only; does not modify document. For heading outline, use word_get_structure.",
    },
    async () => {
      const status = await appManager.getStatus()
      if (!status.activeDocument) {
        return "No active document. Use word_document(path) to open a file, or word_stream_start to create a new one."
      }
      const info = await docOps.getInfo()
      return `Action: Document info\nDetail: "${info.name}" — ${info.wordCount} words, ${info.paragraphCount} paras, ${info.pageCount} pages, ${info.characterCount} chars, ${info.sectionCount} sections, saved: ${info.saved}\nNext: word_get_structure() or word_insert_at(...)`
    },
  )

  regTool("word_get_structure",
    {
      description: "WHEN: need to navigate or understand the document's organization via heading hierarchy. WHAT: returns a hierarchical outline with paragraph indices (e.g., 'H1 ¶3 — Introduction'). CONSTRAINT: read-only. For full text content, use word_get_text. Paragraph indices from this output can be used with word_go_to_paragraph.",
    },
    async () => {
      const structure = await docOps.getStructure()
      if (structure.headings.length === 0) {
        return `Action: No headings found\nDetail: ${structure.totalParagraphs} total paragraphs\nNext: word_go_to_paragraph({index:2}) to navigate or word_apply_style({styleName:"Heading 1"}) to add`
      }
      const lines = [
        `Action: Document structure — ${structure.totalParagraphs} paragraphs, ${structure.headings.length} headings`,
        `Next: word_go_to_paragraph({index:H.paragraphIndex}) to navigate`,
        "",
      ]
      for (const h of structure.headings) {
        const indent = "  ".repeat(h.level - 1)
        lines.push(`${indent}H${h.level} ¶${h.paragraphIndex} — ${h.text}`)
      }
      return lines.join("\n")
    },
  )

  regTool("word_quit",
    {
      description: "WHEN: need to completely shut down Microsoft Word (e.g., to force-restart after a hang). WHAT: closes all documents and quits Word.exe. CONSTRAINT: already-saved files are preserved; unsaved changes may be lost. Server restarts Word automatically on next operation. Use word_close instead for normal document closing.",
    },
    async () => {
      await appManager.quit()
      return "Action: Word closed\nNext: Restart MCP server to reconnect to Word"
    },
    { preconditions: [] },
  )
}
