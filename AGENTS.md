# Word MCP Server — Developer Quick Reference

## Build & test
```bash
npm run build       # tsc, output to build/
npm test            # vitest run --exclude 'tests/integration/**' (skips smoke tests, needs Word)
npm run dev         # tsc --watch
npx tsc --noEmit    # typecheck-only (no linter configured)
```

## Architecture constraints (enforced by tests/arch/arch.test.ts)
- `src/word/` must NOT import from `src/server/`
- `src/server/` must NOT access `.raw` on COM proxy objects

## Entry points
- `build/parent.js` — watchdog process, spawns/restarts `build/child.js`
- `build/child.js` — MCP server, creates `WordSession` + wires all tools
- MCP client config: `{ "command": "node", "args": ["path/to/build/parent.js"] }`

## Platform
- Windows-only (`"os": ["win32"]`, `"cpu": ["x64"]`)
- Requires MSVC Build Tools (winax native addon), Word installed

## COM proxy layers
- `word/com-proxy/`: DocumentProxy, SelectionProxy, RangeProxy — typed COM wrappers
- SectionProxy + HeaderFooterProxy + FieldProxy — section/header/footer operations
- CollectionProxy\<T\> with itemFactory — generic COM collection wrapper

All proxy methods wrap COM calls with try-catch + safe defaults. No bare COM access in business logic (arch test enforces this).

## Key patterns
- **Content writing**: always streaming. `word_stream_start({baseStyleProfile:{...}}) → word_stream_block(×N) → word_stream_end()`.
- **Tool registration**: `createRegTool` → `mcpCall` middleware: rate limit → precheck → ensureReady → timeout → handler → audit → circuit breaker → recovery.
- **Cursor**: ContextSanitizer ensures cursor in main body before edits. After header/footer/page-number/watermark ops, cursor auto-restored.
- **Anti-patterns** in `src/index.ts:68-98` (set font before typing, find+goTo, undo after save, word_open vs word_document).

## Mocks
- `word/com-proxy/com-proxy.mock.ts`: MockDocumentProxy, MockSelectionProxy, MockRangeProxy, mockSectionProxy, mockHeaderFooterProxy.
- `tests/unit/test-helpers.ts`: createMockDoc, createMockSel, createMockSession.
- `tests/arch/` uses ts-morph (devDep) for AST-level constraint checks.

## Config
- `.env.example` → copy to `.env`: ALLOWED_DIRECTORIES, MAX_FILE_SIZE, RATE_LIMIT_*, etc.
- CI: `.github/workflows/ci.yml` (type-check and unit tests on Ubuntu).
