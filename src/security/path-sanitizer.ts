import { resolve, normalize, relative, parse, isAbsolute, join, sep } from "node:path"
import { homedir } from "node:os"
import { existsSync, statSync } from "node:fs"
import { PathSecurityError } from "./errors.js"

export interface PathPolicy {
  allowedDirectories: string[]
  allowNetworkPaths: boolean
  maxFileSize: number
}

export class PathSanitizer {
  private policy: PathPolicy

  constructor(policy: PathPolicy) {
    this.policy = policy
  }

  resolveAndValidate(rawPath: string): string {
    const expanded = this.expandRawPath(rawPath)
    const resolved = this.resolveInput(expanded)

    this.rejectIfTraversal(expanded, resolved)
    this.rejectIfSpecialDevicePath(expanded)
    this.rejectIfNetworkPath(resolved)
    this.rejectIfBlocklisted(resolved)

    if (this.policy.allowedDirectories.length > 0) {
      this.rejectIfNotAllowed(resolved)
    }

    return resolved
  }

  /**
   * Accept the path shapes MCP clients commonly send: surrounding quotes,
   * `~` for the home directory, and forward slashes.
   */
  private expandRawPath(rawPath: string): string {
    let p = rawPath.trim()
    if (p.length >= 2) {
      const first = p[0]
      const last = p[p.length - 1]
      if ((first === '"' && last === '"') || (first === "'" && last === "'")) {
        p = p.slice(1, -1).trim()
      }
    }
    if (p === "~") return homedir()
    if (p.startsWith("~/") || p.startsWith("~\\")) return join(homedir(), p.slice(2))
    return p
  }

  private resolveInput(expanded: string): string {
    const direct = resolve(normalize(expanded))
    if (isAbsolute(expanded) || existsSync(direct)) return direct
    // Relative paths are resolved against the server cwd first; if that misses,
    // try the configured allowed roots so "Documents\\file.docx" still works.
    for (const dir of this.policy.allowedDirectories) {
      const candidate = resolve(normalize(dir), expanded)
      if (existsSync(candidate)) return candidate
    }
    return direct
  }

  validateForRead(rawPath: string): string {
    const resolved = this.resolveAndValidate(rawPath)

    if (!existsSync(resolved)) {
      throw new PathSecurityError(`File does not exist: ${resolved}`)
    }

    const stats = statSync(resolved)
    if (!stats.isFile()) {
      throw new PathSecurityError(`Not a file: ${resolved}`)
    }

    if (stats.size > this.policy.maxFileSize) {
      throw new PathSecurityError(
        `File too large: ${stats.size} bytes (max ${this.policy.maxFileSize} bytes)`
      )
    }

    return resolved
  }

  validateForWrite(rawPath: string): string {
    const resolved = this.resolveAndValidate(rawPath)
    return resolved
  }

  private rejectIfTraversal(rawPath: string, resolved: string): void {
    const cwd = process.cwd()
    if (rawPath.includes("..") || rawPath.startsWith("~") || rawPath.includes("\\~") || rawPath.includes("/~")) {
      const resRoot = parse(resolved).root
      const cwdRoot = parse(cwd).root
      if (resRoot !== cwdRoot) {
        throw new PathSecurityError(`Cross-drive path not allowed: ${resolved}`)
      }
      const rel = relative(cwd, resolved)
      if (rel.startsWith("..")) {
        throw new PathSecurityError(`Path traversal detected: ${resolved}`)
      }
    }
  }

  private rejectIfNetworkPath(resolved: string): void {
    if (!this.policy.allowNetworkPaths) {
      if (resolved.startsWith("\\\\") || resolved.startsWith("//")) {
        throw new PathSecurityError(`Network paths not allowed: ${resolved}`)
      }
    }
  }

  private rejectIfBlocklisted(resolved: string): void {
    const blocklist = [
      /^[a-zA-Z]:\\windows(\\|$)/i,
      /^[a-zA-Z]:\\program files(\\|$)/i,
      /^[a-zA-Z]:\\program files \(x86\)(\\|$)/i,
      /^[a-zA-Z]:\\programdata(\\|$)/i,
    ]
    for (const pattern of blocklist) {
      if (pattern.test(resolved)) {
        throw new PathSecurityError(`Access to system directory blocked: ${resolved}`)
      }
    }
  }

  private rejectIfSpecialDevicePath(rawPath: string): void {
    if (rawPath.startsWith("\\\\.") || rawPath.startsWith("\\\\.\\") ||
        rawPath.startsWith("\\\\?\\") || rawPath.startsWith("\\??\\")) {
      throw new PathSecurityError(`Device path not allowed: ${rawPath}`)
    }
    const normal = rawPath.replace(/\//g, "\\")
    const fileName = normal.split("\\").pop() ?? normal
    if (fileName.includes(":") && normal !== fileName) {
      throw new PathSecurityError(`Alternate data stream not allowed: ${rawPath}`)
    }
  }

  private rejectIfNotAllowed(resolved: string): void {
    const target = resolved.toLowerCase().replace(/[\\/]+$/, "")
    const isAllowed = this.policy.allowedDirectories.some((d) => {
      const dir = resolve(normalize(d)).toLowerCase().replace(/[\\/]+$/, "")
      return target === dir || target.startsWith(dir + sep)
    })
    if (!isAllowed) {
      throw new PathSecurityError(`Path not in allowed directories: ${resolved}`)
    }
  }
}
