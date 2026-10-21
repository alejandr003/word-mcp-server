import type { IWordSession } from "../word/session.js"
import type { PositionMap } from "../word/position-map.js"
import type { IStreamLock } from "../word/types.js"
import type { WordApplicationManager } from "../word/application.js"
import { CircuitBreaker } from "../word/circuit-breaker.js"
import { SessionPathMachine, type SessionPath } from "./session-path-machine.js"
import { isReadOnlyTool } from "./tools/shared.js"
import type { ILogger } from "../logger.js"
import type { SystemMeta } from "./tool-result.js"

export type Precondition = "DOC" | "NO_DOC"

const STREAM_BLOCKED_TOOLS = new Set([
  "word_document", "word_open", "word_quit",
])

const EDIT_BLOCKED_TOOLS = new Set([
  "word_stream_start",
])

const WATCHDOG_INTERVAL = 5000
const HUNG_THRESHOLD = 30000

export class SessionDirector implements IStreamLock {
  private _pathMachine: SessionPathMachine
  private _session: IWordSession | null
  private _positionMap: PositionMap | null
  private _appManager: WordApplicationManager | undefined
  private _logger: ILogger | null = null
  private _circuitBreaker: CircuitBreaker
  private _watchdogTimer: ReturnType<typeof setInterval> | null = null

  constructor(session: IWordSession | null, positionMap: PositionMap | null, appManager?: WordApplicationManager) {
    this._pathMachine = new SessionPathMachine()
    this._session = session
    this._positionMap = positionMap
    this._appManager = appManager
    this._circuitBreaker = new CircuitBreaker()
  }

  setLogger(logger: ILogger): void {
    this._logger = logger
    this._pathMachine.setLogger(logger)
  }

  private _log(level: string, message: string): void {
    if (level === "error") this._logger?.error(message)
    else if (level === "warn") this._logger?.warn(message)
    else if (level === "debug") this._logger?.debug(message)
    else this._logger?.info(message)
  }

  get isStreamingActive(): boolean {
    return this._pathMachine.isStreamingActive
  }

  get currentPath(): SessionPath {
    return this._pathMachine.currentPath
  }

  get session(): IWordSession | null {
    return this._session
  }

  get circuitBreaker(): CircuitBreaker {
    return this._circuitBreaker
  }

  startWatchdog(): void {
    if (this._watchdogTimer) return
    this._watchdogTimer = setInterval(() => this._watchdogTick(), WATCHDOG_INTERVAL)
  }

  stopWatchdog(): void {
    if (this._watchdogTimer) {
      clearInterval(this._watchdogTimer)
      this._watchdogTimer = null
    }
  }

  private _watchdogTick(): void {
    const session = this._session
    if (!session || !session.isAlive()) return

    if (session.isUnhealthy()) {
      this._log("warn", "Watchdog: session unhealthy, attempting recovery")
      session.recover().then(() => {
        this._circuitBreaker.forceReset()
        this._log("info", "Watchdog: session recovered successfully")
      }).catch((err) => {
        this._log("error", `Watchdog: recovery failed: ${err}`)
      })
      return
    }

    const lastCallStart = session.getLastCallStart()
    if (lastCallStart > 0 && Date.now() - lastCallStart > HUNG_THRESHOLD) {
      this._log("warn", `Watchdog: COM call hung for >${HUNG_THRESHOLD}ms, killing Word`)
      session.recover().then(() => {
        this._circuitBreaker.forceReset()
        this._log("info", "Watchdog: Word recovered after hung call")
      }).catch((err) => {
        this._log("error", `Watchdog: recovery failed: ${err}`)
      })
    }
  }

  get positionMap(): PositionMap | null {
    return this._positionMap
  }

  acquireStreamLock(toolName: string): string | null {
    return this._pathMachine.acquireStreamLock(toolName)
  }

  releaseStreamLock(): void {
    this._pathMachine.releaseStreamLock()
  }

  enterEditMode(): void {
    this._pathMachine.enterEditMode()
  }

  exitEditMode(): void {
    this._pathMachine.exitEditMode()
  }

  refreshWatchdog(): void {
    this._pathMachine.refreshWatchdog()
  }

  async precheck(
    toolName: string,
    precondition?: Precondition | Precondition[],
  ): Promise<{ ok: true } | { ok: false; error: string }> {
    if (this._pathMachine.currentPath === "streaming" && STREAM_BLOCKED_TOOLS.has(toolName)) {
      return {
        ok: false,
        error: `[STREAMING] Currently in a stream session.\n>> Recovery: End the stream session with word_stream_end first, then use ${toolName}.`,
      }
    }
    if (this._pathMachine.currentPath === "editing" && EDIT_BLOCKED_TOOLS.has(toolName)) {
      return {
        ok: false,
        error: "[EDIT_MODE] Currently in document editing mode.\n>> Recovery: Close the current document with word_close() first, then use word_stream_start to create a new document.",
      }
    }

    const session = this._session
    if (session) {
      if (session.isBusy() && session.isAlive() && !session.isUnhealthy()) {
        this._log("info", "Clearing stale busy flag")
        session.setBusy(false)
      }
      if (session.isUnhealthy()) {
        try {
          await session.recover()
          this._circuitBreaker.forceReset()
        } catch (e) { this._logger?.warn({ err: e }, "precheck recovery failed") }
      } else if (!session.isAlive()) {
        try {
          // Read-only tools must not launch Word. Attach to an already-running
          // instance when possible; other tools force a start (attach-first).
          if (isReadOnlyTool(toolName)) session.attachIfRunning()
          else session.start()
        } catch (e) { this._logger?.warn({ err: e }, "precheck session start failed") }
      }
    }

    try {
      this._circuitBreaker.check()
    } catch (err) {
      if (session?.isAlive() && !session.isUnhealthy()) {
        this._circuitBreaker.forceReset()
      } else {
        const breakerErr = err as Error
        return { ok: false, error: breakerErr.message }
      }
    }

    const preErr = this._checkPreconditions(precondition)
    if (preErr) return { ok: false, error: preErr }

    return { ok: true }
  }

  private _checkPreconditions(precondition?: Precondition | Precondition[]): string | null {
    const session = this._session
    if (!session) return null

    let p: Precondition | undefined
    if (Array.isArray(precondition)) {
      if (precondition.length === 0) return null
      p = precondition[0]
    } else {
      p = precondition ?? "DOC"
    }
    if (p === "DOC" && !session.activeDoc) {
      try {
        // Do not spawn Word just to check a precondition.
        if (!session.isAlive() && !session.attachIfRunning()) {
          return "[NO_DOCUMENT] Microsoft Word is not running.\n>> Recovery: Use word_document({path:'...'}) to open a file, or word_stream_start({title:'...'}) to create a new document."
        }
        if (session.tryAdoptActiveDoc()) {
          const doc = session.activeDoc
          const fullName = session.activeDocPath
          if (doc && fullName) {
            try {
              this._appManager?.adoptDocument(fullName, doc)
            } catch { this._logger?.warn("adoptDocument failed during precheck") }
          }
        }
      } catch { this._logger?.warn("activeDoc resolution failed during precheck") }
      if (!session.activeDoc) {
        return "[NO_DOCUMENT] No document is currently open.\n>> Recovery: Use word_stream_start({title:'...'}) to create a new document, or word_document({path:'...'}) to open an existing one."
      }
    }
    if (p === "NO_DOC" && session.activeDoc) {
      return "[DOC_ACTIVE] A document is currently open.\n>> Recovery: Close the current document with word_close() first, then create a new document."
    }
    return null
  }

  buildToolMeta(): SystemMeta | null {
    try {
      const session = this._session
      if (!session) return null

      // struct
      let struct: SystemMeta["struct"] | undefined
      const pm = this._positionMap
      if (pm) {
        struct = {
          v: pm.docVersion,
          p: pm.cachedParaCount,
          h: pm.cachedHeadingCount,
          t: pm.cachedTableCount,
        }
      }

      // doc
      let doc: SystemMeta["doc"] | undefined
      if (!session.activeDoc) {
        doc = { name: null, state: "none" }
      } else {
        const path = session.activeDocPath
        doc = path
          ? { name: path.split(/[\\/]/).pop() ?? "?", state: "named" }
          : { name: null, state: "untitled" }
      }

      if (!doc) return null
      const result: SystemMeta = { struct, doc }
      return result
    } catch {
      return null
    }
  }

  markHealthy(): void {
    this._session?.markHealthy()
  }

  markDirtyIfNeeded(toolName: string): void {
    if (!isReadOnlyTool(toolName)) this._positionMap?.markDirty()
  }

  schedulePositionRefresh(): void {
    this._positionMap?.scheduleRefresh()
  }

  async recoverSession(): Promise<void> {
    await this._session?.recover()
    this._circuitBreaker.forceReset()
  }
}
