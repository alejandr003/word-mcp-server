# Domain Context

## Core entities

**WordSession** — a COM connection to `WINWORD.EXE`. Manages the `Word.Application` lifecycle (start/quit), health tracking (`isAlive` / `healthCheck` / `recover`) and the active document pointer. All word-processing modules depend on this session.

**Document** — a Word document tracked by a resolved path. In `DocumentRegistry` it has an active state (the session's `activeDoc`), a path and a saved state.

**Selection** — the cursor position or text selection in the current document. All editing operations go through the selection.

**PositionMap** — an in-process cache of document heading and table positions, built from a single batched COM read. Supports O(1) lookups for "go to heading X" operations.

## Core value objects

**DocumentInfo** — `wordCount`, `paragraphCount`, `pageCount`, `characterCount`, `sectionCount`, `saved`

**HeadingEntry** — `level`, `text`, `paragraphIndex`

**Locator** — discriminated union of heading/paragraph/table/bookmark locators with match modes and offsets

**ResolvedPosition** — `found`, `paragraphIndex`, `headingContext`, `tableIndex`

**WordMcpError** — `code`, `recoverable` flag, `recoveryHint` string

**Document states** — `NO_WORD`, `NO_DOC`, `DOC_ACTIVE`, `DIALOG`

## Modules

**CursorPosition** — cursor state manager. Ensures that after every call to `ensureMainBody()` the selection is in the document's main body (not in a header, footer, table or shape). Tracks position with markers when switching between these contexts. Referenced in `WordBase` and shared by all editor classes.

**ChartDataBridge** — chart data child-process manager. Spawns a separate Node.js child process (`chart-data-worker.js`) over IPC to set chart data. Manages an idle timeout (60 s auto-kill) and a per-operation timeout (15 s). Referenced in `WordMediaEditor` and injected through the `IChartDataBridge` seam.

**WordStreamWriter** — streaming document writer engine. Supports writing Markdown content in chunks within a single COM session, with content appearing instantly in the Word window. Works with the `word_stream_start` / `word_stream_block` / `word_stream_end` tool chain for zero-boilerplate document generation. Automatically manages style inheritance, cursor positioning and batch flushing.

**SessionDirector** — session orchestrator. Routes document-level operations such as `word_document` / `word_create` / `word_get_status`, and coordinates interactions between `WordSession` and `DocumentRegistry`. Provides a unified document entry-point abstraction at the `server/` layer.

**ServerContext** — shared session dependencies required by the MCP tool pipeline. Contains references to `IWordSession`, `PositionMap` and `DocumentRegistry`. Does not perform hidden module-level global connections; all tool registration functions receive it as an explicit parameter.

**ProcessMonitor** — child-process health watchdog. Monitors the lifecycle of spawned processes such as `chart-data-worker`, manages idle timeouts and cleans up resources after abnormal termination. Provides process-level guarantees in the `word/` layer.

**StreamingMarkdownWriter** — streaming document writer engine (`word/word-stream-writer.ts`). Supports chunked Markdown writes within a single COM session, with instant preview in Word. Works with the `word_stream_start` / `word_stream_block` / `word_stream_end` tool chain for zero-boilerplate document generation. Automatically manages style inheritance, cursor positioning and batch flushing. Supports `baseStyleProfile` to preconfigure style definitions (font, paragraph) when creating a document; Markdown writes inherit them automatically.

**COM Proxy** — typed COM object wrapper layer (`word/com-proxy/`). `DocumentProxy` / `SelectionProxy` / `RangeProxy` wrap the three most-used COM objects, eliminating `Record<string, unknown>` type unsafety across the codebase. Lazily constructed in the session layer through `getDocProxy()` / `getSelectionProxy()` / `wrapRange()` factories. Sub-objects (Find, Font, Shading, etc.) keep the `Record` escape hatch and will be refined later.

## Architecture overview

```
┌─────────────────────────────────────────────────────────────────┐
│                        MCP Client (stdio)                        │
└──────────────────────────────┬──────────────────────────────────┘
                               │
                    ┌──────────▼──────────┐
                    │     parent.js        │  <- watchdog process: monitors/restarts the child
                    │  (Watchdog + spawn)  │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │      child.js        │  <- MCP server main process
                    │  (createServer)      │
                    └──────────┬──────────┘
                               │
              ┌────────────────┼────────────────┐
              │                │                │
     ┌────────▼──────┐ ┌──────▼───────┐ ┌──────▼────────┐
     │ Server Layer  │ │  Word Layer  │ │ Security Layer│
     │               │ │              │ │               │
     │ SessionDirector│ │ WordSession  │ │ SecurityManager│
     │ ServerContext │ │ COM Proxies  │ │ PathSanitizer │
     │ Tool Modules  │ │ ContentWriter│ │ RateLimiter   │
     │ Prompts       │ │ Stream Writer│ │ Audit         │
     │ Resources     │ │ PositionMap  │ │               │
     └───────────────┘ └──────┬───────┘ └───────────────┘
                               │
                    ┌──────────▼──────────┐
                    │    WINWORD.EXE      │  <- COM automation
                    │   (winax addon)     │
                    └─────────────────────┘
```

## Data flow

1. The **MCP client** sends JSON-RPC requests over stdio.
2. **parent.js**, the watchdog, receives requests and forwards them to child.js.
3. `createServer` in **child.js** initializes all modules.
4. **SessionDirector** runs prechecks (state machine, rate limit, circuit breaker).
5. **Tool handlers** call Word-layer modules to perform the actual operations.
6. The **COM Proxy** layer translates operations into Word COM API calls.
7. Results are returned to the MCP client over stdio.

## Key design patterns

### Streaming document writing

```
word_stream_start({baseStyleProfile}) -> word_stream_block(x N) -> word_stream_end()
```

- Configure styles once; subsequent content inherits them automatically.
- Batched flushing (200 ms batches) avoids overloading COM calls.
- Markdown parsing -> segmented rendering -> instant display in Word.

### Tool registration pipeline

```
createRegTool -> mcpCall middleware:
  rate limit -> precheck -> ensureReady -> timeout -> handler -> audit -> circuit breaker -> recovery
```

### Cursor management

- `ContextSanitizer` ensures the cursor is in the main body before editing.
- Cursor position is restored automatically after header, footer, table and watermark operations.
- Uses the `wasInNonBody` flag to track context transitions.

### Session recovery

- `WordSession.comCall()` wraps all COM calls.
- Transient errors (`RPC_E_CALL_REJECTED`, etc.) automatically mark the session unhealthy.
- The `SessionDirector` watchdog detects hanging COM calls and triggers recovery.
- Recovery flow: kill process -> cleanup -> rebuild the COM session.

### Position cache

- `PositionMap` caches heading, table and paragraph positions.
- Uses binary search for O(1) lookups.
- Marks data dirty after document modifications and refreshes lazily.

### Security

- **Path sanitization** — prevents directory traversal attacks and restricts access to allowed directories.
- **Rate limiting** — prevents tool-call overload.
- **Zod validation** — type-safe validation of all input parameters.
- **Macro disabling** — `AutomationSecurity = 3` (high security level).
- **Audit logging** — records every tool call and error.
