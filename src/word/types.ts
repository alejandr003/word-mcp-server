export interface DocumentInfo {
  name: string
  fullName: string
  path: string
  wordCount: number
  paragraphCount: number
  pageCount: number
  characterCount: number
  sectionCount: number
  saved: boolean
}

export interface HeadingEntry {
  level: number
  text: string
  paragraphIndex: number
}

import { resolve, normalize } from "node:path"

export interface IStreamLock {
  acquireStreamLock(toolName: string): string | null
  releaseStreamLock(): void
  refreshWatchdog(): void
}

/** Word 颜色索引映射 */
export const COLOR_INDEX: Record<string, number> = {
  auto: 0, black: 1, blue: 2, turquoise: 3, bright_green: 4, pink: 5,
  red: 6, yellow: 7, white: 8, dark_blue: 9, teal: 10, green: 11,
  violet: 12, dark_red: 13, dark_yellow: 14, gray_50: 15, gray_25: 16,
}

/** PDF 导出格式代码（wdExportFormatPDF） */
export const EXPORT_FORMAT_PDF = 17

/** 规范化文件路径：解析 + 标准化 + 小写 */
export function normalizePath(p: string): string {
  return resolve(normalize(p.trim())).toLowerCase()
}

function safeDecode(p: string): string {
  try {
    return decodeURIComponent(p)
  } catch {
    return p
  }
}

function pathSegments(p: string): string[] {
  return safeDecode(p)
    .replace(/[?#].*$/, "")
    .split(/[\\/]+/)
    .filter(Boolean)
    .map(s => s.toLowerCase())
}

function isWebUrl(p: string): boolean {
  return /^https?:\/\//i.test(p.trim())
}

/**
 * Decide whether a path supplied by the caller and a document's FullName refer to
 * the same file.
 *
 * Word reports cloud-backed documents (OneDrive/SharePoint) as an https:// URL in
 * `Document.FullName`, while callers pass the local synced path — for example:
 *   C:\Users\me\OneDrive - Org\Documentos\TI91\Unit 2.docx
 *   https://org-my.sharepoint.com/personal/me/Documents/Documentos/TI91/Unit%202.docx
 * A plain string comparison considers them different and opens a duplicate copy,
 * so cloud URLs are compared by file name plus the shared parent-folder suffix.
 */
export function pathsReferToSameFile(
  target: string,
  documentFullName: string,
): boolean {
  const t = target.trim()
  const d = documentFullName.trim()
  if (!t || !d) return false

  if (isWebUrl(t) && isWebUrl(d)) {
    return safeDecode(t).replace(/\/+$/, "").toLowerCase() ===
      safeDecode(d).replace(/\/+$/, "").toLowerCase()
  }
  if (isWebUrl(t) !== isWebUrl(d)) {
    const url = isWebUrl(t) ? t : d
    const local = isWebUrl(t) ? d : t
    const urlSegments = pathSegments(url)
    const localSegments = pathSegments(local)
    if (urlSegments.length === 0 || localSegments.length === 0) return false
    if (urlSegments[urlSegments.length - 1] !== localSegments[localSegments.length - 1]) {
      return false
    }
    // Require the file name plus at least one shared parent folder so two
    // different files with the same name are not confused with each other.
    let shared = 0
    const max = Math.min(urlSegments.length, localSegments.length)
    for (let i = 1; i <= max; i++) {
      if (urlSegments[urlSegments.length - i] === localSegments[localSegments.length - i]) shared++
      else break
    }
    return shared >= 2 || urlSegments.length === 1 || localSegments.length === 1
  }

  return normalizePath(t) === normalizePath(d)
}
