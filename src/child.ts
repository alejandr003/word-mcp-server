import { createServer } from "./server/create-server.js"

const instructions = [
  "Word MCP Server — Microsoft Word COM automation.",
  "",
  "States (call word_get_status first):",
  "- NO_WORD: restart server.",
  "- NO_DOC: word_stream_start (new) or word_document(path).",
  "- DOC_ACTIVE: all tools available.",
  "- DIALOG: dismiss dialog manually.",
  "",
"Streaming (new documents): word_stream_start → word_stream_block(×N) → word_stream_end",
"Style convention: pre-configure fonts/spacing via word_stream_start.baseStyleProfile, then write pure markdown — styles auto-inherited.",
"Streaming supports: headings, bold, italic, code, lists, tables, blockquotes, code blocks, links, rules.",
  "Rich elements: word_insert_chart/image/textbox, word_set_page_region/page_numbers/watermark (add quiet:true for pipeline mode).",
  "Editing existing: word_document → word_insert_at with markdown text.",
  "",
  "ANTI-PATTERNS:",
  "1. word_stream_start while doc active → auto-closes existing doc.",
  "2. word_save then word_undo_redo → undo cleared.",
  "3. Format before typing → type first, select, format.",
  "4. word_document({path}) activates an already-open file; do not retry with word_open.",
  "5. word_find_text + word_go_to_paragraph → find already moves cursor.",
].join("\n")

createServer({ instructions, enableHeartbeat: true, logPrefix: "word-mcp" }).catch((error) => {
  console.error("[word-mcp/child] Fatal error:", error)
  process.exit(1)
})
