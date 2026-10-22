import { spawn } from "node:child_process"
import { resolve, dirname } from "node:path"
import { fileURLToPath } from "node:url"
import { randomUUID } from "node:crypto"
import pino from "pino"

const __filename = fileURLToPath(import.meta.url)
const __dirname = dirname(__filename)

const CHILD_SCRIPT = resolve(__dirname, "child.js")
const WATCHDOG_TIMEOUT_MS = parseInt(process.env.WATCHDOG_TIMEOUT_MS ?? "30000", 10)
const WATCHDOG_INTERVAL_MS = parseInt(process.env.WATCHDOG_INTERVAL_MS ?? "5000", 10)

// stdout is the MCP transport — logs must go to stderr.
const logger = pino(
  { name: "word-mcp/parent", level: process.env.LOG_LEVEL ?? "info" },
  pino.destination(2),
)

class Watchdog {
  private child: ReturnType<typeof spawn> | null = null
  private lastChildOutput = Date.now()
  private restarting = false
  private shutdown = false
  private stdinHandler: ((chunk: Buffer) => void) | null = null
  private generation = 0

  start(): void {
    this.spawnChild()
    this.startWatchdog()
    this.setupSignalHandlers()
  }

  private spawnChild(): void {
    if (this.restarting) return
    this.generation++
    const myGeneration = this.generation
    const childTraceId = randomUUID()
    logger.info({ childTraceId, generation: myGeneration }, "Spawning child")

    this.lastChildOutput = Date.now()
    this.child = spawn("node", [CHILD_SCRIPT], {
      // Dedicated IPC channel for heartbeats. STDIO stays untouched for MCP.
      stdio: ["pipe", "pipe", "pipe", "ipc"],
      windowsHide: false,
      env: { ...process.env, CHILD_TRACE_ID: childTraceId },
    })

    const child = this.child
    const childStdin = child.stdin!
    const childStdout = child.stdout!
    const childStderr = child.stderr!

    // 清理旧的 stdin listener (restart 时)
    const ps = process.stdin!
    if (this.stdinHandler) ps.removeListener("data", this.stdinHandler)

    this.stdinHandler = (chunk: Buffer) => {
      try { childStdin.write(chunk) } catch { /* child may be dead */ }
    }
    ps.on("data", this.stdinHandler)

    // 子进程 stdout → 父进程 stdout (MCP 响应)
    childStdout.on("data", (chunk: Buffer) => {
      this.lastChildOutput = Date.now()
      process.stdout.write(chunk)
    })

    // 子进程 stderr → 父进程 stderr (日志 + 心跳检测)
    childStderr.on("data", (chunk: Buffer) => {
      this.lastChildOutput = Date.now()
      process.stderr.write(chunk)
    })

    // 子进程退出处理
    child.on("exit", (code, signal) => {
      logger.info({ code, signal, generation: myGeneration }, "Child exited")
      if (this.generation !== myGeneration) return
      if (!this.shutdown) {
        logger.warn("Restarting child after unexpected exit...")
        this.spawnChild()
      }
    })

    // Heartbeats are sent over IPC so they do not pollute the MCP stdio stream.
    child.on("message", () => {
      this.lastChildOutput = Date.now()
    })

    child.on("error", (err) => {
      logger.error({ err: err.message }, "Child error")
    })

    logger.info("Child process spawned")
  }

  private startWatchdog(): void {
    setInterval(() => {
      if (this.shutdown || !this.child) return
      if (this.child.killed) return

      const idle = Date.now() - this.lastChildOutput
      if (idle > WATCHDOG_TIMEOUT_MS) {
        logger.warn({ idleMs: idle }, "Watchdog: no child output, killing and restarting...")
        this.restarting = true
        this.child.kill("SIGTERM")
        // 等待进程退出
        setTimeout(() => {
          this.restarting = false
          this.spawnChild()
        }, 2000)
      }
    }, WATCHDOG_INTERVAL_MS)
  }

  private setupSignalHandlers(): void {
    const cleanup = () => {
      this.shutdown = true
      logger.info("Shutting down parent...")
      if (this.child && !this.child.killed) {
        this.child.kill("SIGTERM")
      }
      process.exit(0)
    }

    process.on("SIGINT", cleanup)
    process.on("SIGTERM", cleanup)
    process.on("unhandledRejection", (reason) => {
      logger.error({ err: reason }, "Unhandled rejection")
    })
  }
}

const watchdog = new Watchdog()
watchdog.start()
