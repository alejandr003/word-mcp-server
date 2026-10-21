import { describe, it, expect } from "vitest"
import { pathsReferToSameFile } from "../../../src/word/types.js"

const LOCAL =
  "C:\\Users\\aleja\\OneDrive - Universidad Tecnologica\\Documentos\\TI91\\Unit 2 Evaluation.docx"
const SHAREPOINT =
  "https://universi799-my.sharepoint.com/personal/aleja/Documents/Documentos/TI91/Unit%202%20Evaluation.docx"

describe("pathsReferToSameFile", () => {
  it("matches identical local paths regardless of case and slashes", () => {
    expect(pathsReferToSameFile(LOCAL, LOCAL.toLowerCase().replace(/\\/g, "/"))).toBe(true)
  })

  it("matches a local path with the SharePoint URL Word reports", () => {
    expect(pathsReferToSameFile(LOCAL, SHAREPOINT)).toBe(true)
    expect(pathsReferToSameFile(SHAREPOINT, LOCAL)).toBe(true)
  })

  it("does not match different files with the same name", () => {
    const other = "C:\\Users\\aleja\\Desktop\\Unrelated\\Unit 2 Evaluation.docx"
    expect(pathsReferToSameFile(other, SHAREPOINT)).toBe(false)
  })

  it("does not match different file names", () => {
    expect(pathsReferToSameFile("C:\\tmp\\a.docx", "C:\\tmp\\b.docx")).toBe(false)
  })

  it("matches identical URLs", () => {
    expect(pathsReferToSameFile(SHAREPOINT, SHAREPOINT)).toBe(true)
  })

  it("returns false for empty values", () => {
    expect(pathsReferToSameFile("", LOCAL)).toBe(false)
    expect(pathsReferToSameFile(LOCAL, "")).toBe(false)
  })
})
