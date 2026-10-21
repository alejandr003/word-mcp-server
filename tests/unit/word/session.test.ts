import { describe, it, expect, vi } from "vitest"
import { WordSession } from "../../../src/word/session.js"
import { ComError, FatalComError, TransientComError } from "../../../src/word/com-errors.js"

function createMockWinax() {
  const mockApp: Record<string, unknown> = {
    Version: "16.0",
    Visible: false,
    DisplayAlerts: 999,
    AutomationSecurity: 999,
    Documents: { Count: 0, Item: vi.fn(), Add: vi.fn() },
    Quit: vi.fn(),
    Release: vi.fn(),
  }
  const MockObject = vi.fn(function (this: unknown, progid: string, _options?: unknown) {
    return mockApp
  })
  const mockWinax = { Object: MockObject as unknown as new (progid: string, options?: { activate?: boolean }) => Record<string, unknown> }
  return { mockWinax, mockApp, MockObject }
}

function createSession(mockWinax: ReturnType<typeof createMockWinax>["mockWinax"], pids: number[] = []) {
  return new WordSession(() => mockWinax, { listWinwordPids: () => pids })
}

describe("WordSession", () => {
  it("initial state is not running, no active doc", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    expect(session.isAlive()).toBe(false)
    expect(session.activeDoc).toBeNull()
  })

  it("start attaches to a running Word instance before creating one", () => {
    const { mockWinax, mockApp } = createMockWinax()
    const session = createSession(mockWinax, [4242])
    session.start()
    expect(mockWinax.Object).toHaveBeenCalledWith("Word.Application", { activate: true })
    expect(mockApp.Visible).toBe(true)
    expect(mockApp.DisplayAlerts).toBe(0)
    expect(mockApp.AutomationSecurity).toBe(3)
    expect(session.isAlive()).toBe(true)
  })

  it("start is idempotent", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    session.start()
    session.start()
    expect(mockWinax.Object).toHaveBeenCalledTimes(1)
  })

  it("ensureAlive starts if not running", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    session.ensureAlive()
    expect(session.isAlive()).toBe(true)
  })

  it("application lazy-starts", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    const app = session.application
    expect(app).not.toBeNull()
    expect(session.isAlive()).toBe(true)
  })

  it("application getter throws FatalComError when unhealthy", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    session.markUnhealthy()
    expect(() => session.application).toThrow(FatalComError)
  })

  it("setActiveDoc stores and activeDoc returns it", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    const fakeDoc = { Name: "test.docx" }
    session.setActiveDoc(fakeDoc as unknown as Record<string, unknown>)
    expect(session.activeDoc).toBe(fakeDoc)
  })

  it("quit clears state when no real Word process", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    session.start()
    session.quit()
    expect(session.isAlive()).toBe(false)
    expect(session.activeDoc).toBeNull()
  })

  it("quit is safe when not started", () => {
    const session = new WordSession(() => ({ Object: vi.fn() }))
    expect(() => session.quit()).not.toThrow()
  })

  it("comCall throws FatalComError on COM failure", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    session.start()
    expect(() => session.comCall(() => { throw new Error("COM error") })).toThrow(FatalComError)
  })

  it("comCall throws TransientComError for transient HRESULT", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    session.start()
    const err = { number: 0x80010005, message: "Call was rejected" }
    expect(() => session.comCall(() => { throw err })).toThrow(TransientComError)
  })

  it("healthCheck returns true after start (optimistic alive)", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    session.start()
    expect(session.healthCheck()).toBe(true)
  })

  it("setLogger stores logger and uses it for logging", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    const infoFn = vi.fn()
    session.setLogger({ info: infoFn } as any)
    session.start()
    expect(infoFn).toHaveBeenCalledWith(expect.stringContaining("Word.Application"))
  })

  it("recover() restarts session asynchronously", async () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax)
    session.start()
    await session.recover()
    expect(session.isAlive()).toBe(true)
  })

  it("attachIfRunning does not create Word when no process is running", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax, [])
    expect(session.attachIfRunning()).toBe(false)
    expect(mockWinax.Object).not.toHaveBeenCalled()
  })

  it("attachIfRunning attaches to an existing Word process", () => {
    const { mockWinax } = createMockWinax()
    const session = createSession(mockWinax, [1234])
    expect(session.attachIfRunning()).toBe(true)
    expect(mockWinax.Object).toHaveBeenCalledWith("Word.Application", { activate: true })
  })

  it("attachIfRunning discards a fallback process it had to create", () => {
    const { mockWinax, mockApp } = createMockWinax()
    let call = 0
    const pids = () => {
      call++
      if (call === 1) return [1234]       // process seen before start()
      if (call === 2) return [1234]       // snapshot before creating
      return [1234, 9999]                 // winax had to create a new instance
    }
    const session = new WordSession(() => mockWinax, { listWinwordPids: pids })
    expect(session.attachIfRunning()).toBe(false)
    expect(mockApp.Quit).toHaveBeenCalled()
    expect(session.isAlive()).toBe(false)
  })

  it("release() detaches without quitting Word", () => {
    const { mockWinax, mockApp } = createMockWinax()
    const session = createSession(mockWinax, [1234])
    session.start()
    session.release()
    expect(mockApp.Quit).not.toHaveBeenCalled()
    expect(session.isAlive()).toBe(false)
  })
})
