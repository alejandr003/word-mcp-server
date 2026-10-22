import pino from "pino"
import { AsyncLocalStorage } from "node:async_hooks"
import { randomUUID } from "node:crypto"

export interface LogEntry {
  level: string
  msg: string
  time: number
}

interface AlsStore {
  traceId: string
  logs: LogEntry[]
}

const als = new AsyncLocalStorage<AlsStore>()

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type LogFn = (...args: any[]) => void

export interface ILogger {
  fatal: LogFn
  error: LogFn
  warn: LogFn
  info: LogFn
  debug: LogFn
  trace: LogFn
  child(bindings: Record<string, unknown>): ILogger
}

export function generateTraceId(): string {
  return randomUUID()
}

export function getTraceId(): string | undefined {
  return als.getStore()?.traceId
}

export function getLogBuffer(): LogEntry[] {
  return als.getStore()?.logs ?? []
}

export function runWithTraceId<T>(traceId: string, fn: () => T): T {
  return als.run({ traceId, logs: [] }, fn)
}

export function createRootLogger(): ILogger {
  // stdout carries the MCP JSON-RPC stream. Logging there corrupts the
  // protocol, so every log line is written to stderr instead.
  const raw = pino(
    { name: "word-mcp", level: process.env.LOG_LEVEL ?? "info" },
    pino.destination(2),
  )
  return alsAware(raw)
}

export function createModuleLogger(root: ILogger, module: string): ILogger {
  return root.child({ module })
}

function alsAware(raw: pino.Logger): ILogger {
  return {
    fatal: (...args) => { send(raw, "fatal", args) },
    error: (...args) => { send(raw, "error", args) },
    warn: (...args) => { send(raw, "warn", args) },
    info: (...args) => { send(raw, "info", args) },
    debug: (...args) => { send(raw, "debug", args) },
    trace: (...args) => { send(raw, "trace", args) },
    child: (bindings) => alsAware(raw.child(bindings)),
  }
}

const sessionLogRing: LogEntry[] = []
const MAX_SESSION_LOGS = 500

function send(raw: pino.Logger, level: string, args: unknown[]): void {
  const msg = String(args.map(a => typeof a === "object" ? JSON.stringify(a) : a).join(" "))
  const entry: LogEntry = { level, msg, time: Date.now() }

  sessionLogRing.push(entry)
  if (sessionLogRing.length > MAX_SESSION_LOGS) sessionLogRing.shift()

  const store = als.getStore()
  if (store) {
    store.logs.push(entry)
    if (store.logs.length > 200) store.logs.splice(0, store.logs.length - 200)
    if (args.length > 0 && typeof args[0] === "object" && args[0] !== null) {
      args[0] = { ...args[0] as Record<string, unknown>, traceId: store.traceId }
    } else {
      args.unshift({ traceId: store.traceId })
    }
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ;(raw as any)[level](...args)
}

export function getSessionLogs(count = 50, level?: string): LogEntry[] {
  const logs = count >= sessionLogRing.length ? sessionLogRing : sessionLogRing.slice(-count)
  return level ? logs.filter(l => l.level === level) : logs
}
