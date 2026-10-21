import { describe, it, expect } from "vitest"
import { PathSanitizer } from "../../../src/security/path-sanitizer.js"

const defaultPolicy = {
  allowedDirectories: [] as string[],
  allowNetworkPaths: false,
  maxFileSize: 50 * 1024 * 1024,
}

describe("PathSanitizer", () => {
  it("should resolve a normal path", () => {
    const s = new PathSanitizer(defaultPolicy)
    const result = s.resolveAndValidate("C:\\Users\\test\\doc.docx")
    if (process.platform === "win32") {
      expect(result).toBe("C:\\Users\\test\\doc.docx")
    } else {
      expect(result).toBeDefined()
    }
  })

  it("should reject Windows system directory", () => {
    const s = new PathSanitizer(defaultPolicy)
    if (process.platform === "win32") {
      expect(() => s.resolveAndValidate("C:\\Windows\\System32\\evil.exe")).toThrow()
    }
  })

  it("should reject network paths when disabled", () => {
    const s = new PathSanitizer({ ...defaultPolicy, allowNetworkPaths: false })
    if (process.platform === "win32") {
      expect(() => s.resolveAndValidate("\\\\server\\share\\doc.docx")).toThrow()
    }
  })

  it("should allow network paths when enabled", () => {
    const s = new PathSanitizer({ ...defaultPolicy, allowNetworkPaths: true })
    const result = s.resolveAndValidate("\\\\server\\share\\doc.docx")
    if (process.platform === "win32") {
      expect(result).toContain("server")
    } else {
      expect(result).toBeDefined()
    }
  })

  it("should reject traversal paths", () => {
    const s = new PathSanitizer(defaultPolicy)
    expect(() => s.resolveAndValidate("..\\..\\windows\\file.exe")).toThrow()
  })

  it("should allow paths within allowed directories", () => {
    const s = new PathSanitizer({
      ...defaultPolicy,
      allowedDirectories: ["C:\\Users\\test\\Documents"],
    })
    const result = s.resolveAndValidate("C:\\Users\\test\\Documents\\report.docx")
    if (process.platform === "win32") {
      expect(result).toBe("C:\\Users\\test\\Documents\\report.docx")
    } else {
      expect(result).toBeDefined()
    }
  })

  it("should reject paths outside allowed directories", () => {
    const s = new PathSanitizer({
      ...defaultPolicy,
      allowedDirectories: ["C:\\Users\\test\\Documents"],
    })
    expect(() =>
      s.resolveAndValidate("C:\\Users\\test\\Desktop\\file.docx")
    ).toThrow()
  })

  it("should validate for read on existing file", () => {
    const s = new PathSanitizer(defaultPolicy)
    expect(() => s.validateForRead("C:\\nonexistent\\file.docx")).toThrow()
  })

  it("should reject over-large files", () => {
    const s = new PathSanitizer({ ...defaultPolicy, maxFileSize: 1 })
    expect(() => s.validateForRead(__filename)).toThrow()
  })

  it("should strip matching quotes around a path", () => {
    const s = new PathSanitizer(defaultPolicy)
    const result = s.resolveAndValidate('"C:\\Users\\test\\doc.docx"')
    if (process.platform === "win32") {
      expect(result).toBe("C:\\Users\\test\\doc.docx")
    } else {
      expect(result).not.toContain('"')
    }
  })

  it("should expand a leading tilde to the home directory", () => {
    const s = new PathSanitizer(defaultPolicy)
    const result = s.resolveAndValidate("~/Documents/doc.docx")
    expect(result).toContain("Documents")
    expect(result).not.toContain("~")
  })

  it("should not allow sibling directories with a shared prefix", () => {
    const s = new PathSanitizer({
      ...defaultPolicy,
      allowedDirectories: ["C:\\Users\\test\\Documents"],
    })
    expect(() =>
      s.resolveAndValidate("C:\\Users\\test\\Documents-other\\file.docx")
    ).toThrow()
  })
})
