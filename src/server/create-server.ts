import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import type { ServerContext } from "./server-context.js"
import { WordSession } from "../word/session.js"
import { WordApplicationManager } from "../word/application.js"
import { WordDocument } from "../word/document.js"
import { ChartDataBridge } from "../word/chart-data-bridge.js"
import { WordContentWriter } from "../word/word-content-writer.js"
import { WordCursor } from "../word/word-cursor.js"
import { WordFormatter } from "../word/word-formatter.js"
import { PositionMap } from "../word/position-map.js"
import { SecurityManager } from "../security/policy.js"
import { SessionDirector } from "./session-director.js"
import { registerDocumentTools } from "./tools/document.js"
import { registerContentTools } from "./tools/content.js"
import { registerCursorTools } from "./tools/cursor.js"
import { registerFormattingTools } from "./tools/formatting.js"
import { registerTableTools } from "./tools/tables.js"
import { registerMediaTools } from "./tools/media.js"
import { registerStructureTools } from "./tools/structure.js"
/* markdown tools removed — streaming is the sole content path */
import { registerReaderTools } from "./tools/reader.js"
import { registerVariableTool } from "./tools/variable.js"
import { registerSemanticTools } from "./tools/semantic.js"
import { registerBatchTools } from "./tools/batch.js"
import { registerWhereAmITool } from "./tools/whereami.js"
import { registerDocumentStructureResource } from "./resources/document-structure-resource.js"
import { registerDiagnosticResource } from "./resources/diagnostic-resource.js"

import { StreamingMarkdownWriter } from "../word/word-stream-writer.js"
import { registerStreamTools } from "./tools/stream.js"
import { registerReportPrompts } from "./prompts/report-prompts.js"
import { registerStateMachinePrompt } from "./prompts/state-machine.js"
import { createRootLogger, createModuleLogger } from "../logger.js"


export interface CreateServerOptions {
  instructions: string
  enableHeartbeat?: boolean
  logPrefix?: string
}

export async function createServer(options: CreateServerOptions): Promise<void> {
  const rootLogger = createRootLogger()
  const log = createModuleLogger(rootLogger, "server")

  log.info("Starting Word MCP Server...")

  const heartbeat = options.enableHeartbeat
    ? setInterval(() => { log.debug("heartbeat alive") }, 10_000)
    : null

  const sessionLogger = createModuleLogger(rootLogger, "session")
  const directorLogger = createModuleLogger(rootLogger, "director")

  const session = new WordSession()
  session.setLogger(sessionLogger)
  const appManager = new WordApplicationManager(session)
  const positionMap = new PositionMap(session)
  const docOps = new WordDocument(session, positionMap)
  const formatter = new WordFormatter(session)
  const chartBridge = new ChartDataBridge()
  const contentWriter = new WordContentWriter(session, chartBridge)
  const cursor = new WordCursor(session)
  const security = new SecurityManager()
  const director = new SessionDirector(session, positionMap, appManager)
  director.setLogger(directorLogger)
  director.startWatchdog()
  const context: ServerContext = { session, positionMap, director, logger: rootLogger }
  const streamWriter = new StreamingMarkdownWriter(
    session, contentWriter, appManager, formatter, director,
  )

  const cleanup = async () => {
    log.info("Shutting down...")
    if (heartbeat) clearInterval(heartbeat)
    // Detach only: quitting Word here would discard the user's unsaved work
    // every time the MCP client reconnects or the child process restarts.
    try { appManager.release() } catch { /* ignore */ }
    process.exit(0)
  }
  process.on("SIGINT", cleanup)
  process.on("SIGTERM", cleanup)
  process.on("unhandledRejection", (reason) => {
    log.error("Unhandled rejection:", reason)
  })
  process.on("uncaughtException", (err) => {
    log.error({ err }, "Uncaught exception")
    process.exit(1)
  })

  const server = new McpServer(
    { name: "word-mcp-server", version: "1.0.0" },
    { instructions: options.instructions },
  )

  registerDocumentTools(server, context, appManager, docOps, security)
  registerContentTools(server, context, contentWriter, security)
  registerCursorTools(server, context, cursor, security)
  registerFormattingTools(server, context, formatter, security)
  registerTableTools(server, context, contentWriter, formatter, security)
  registerMediaTools(server, context, contentWriter, security)
  registerStructureTools(server, context, formatter, security)
  registerReportPrompts(server)
  registerStateMachinePrompt(server)
  registerVariableTool(server, context, contentWriter, security)
  registerReaderTools(server, context, docOps, security)
  registerSemanticTools(server, context, cursor, formatter, contentWriter, positionMap, security)
  registerBatchTools(server, context, cursor, contentWriter, positionMap, security)

  registerStreamTools(server, context, streamWriter, security)
  registerWhereAmITool(server, context, positionMap, security)
  registerDocumentStructureResource(server, context, docOps, positionMap)
  registerDiagnosticResource(server, context)

  const transport = new StdioServerTransport()
  await server.connect(transport)

  log.info("Word MCP Server connected and ready")
}
