# Word MCP Server

<p align="center">
  <video src="https://raw.githubusercontent.com/alejandr003/word-mcp-server/main/assets/demo/demo.mp4" autoplay loop muted playsinline width="300"></video>
</p>

<p align="center">
  <a href="https://github.com/alejandr003/word-mcp-server/blob/main/assets/demo/demo.mp4"><strong>Watch the demo video</strong></a>
</p>

<p align="center">
  <a href="https://opensource.org/licenses/MIT">
    <img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT">
  </a>
  <a href="https://nodejs.org">
    <img src="https://img.shields.io/badge/node-%3E%3D20-brightgreen" alt="Node">
  </a>
  <img src="https://img.shields.io/badge/platform-Windows-blue" alt="Platform">
</p>

> **Fork notice:** This is the maintained fork at
> [alejandr003/word-mcp-server](https://github.com/alejandr003/word-mcp-server),
> based on [HelloWorld-Open/word-mcp-server](https://github.com/HelloWorld-Open/word-mcp-server).
> It adds:
>
> - attach to an already-running Word instance instead of launching a new one;
> - reuse of documents that are already open;
> - a watchdog IPC heartbeat that prevents the child process from restarting every ~35 seconds;
> - logs on `stderr`, so the MCP JSON-RPC stream on `stdout` stays clean.
>
> Clone **this** repository if you want these fixes.

## What it is

Let your AI agent control Microsoft Word directly. 81 tools cover the full workflow: create, edit, format, tables, charts, images, bookmarks, headers, footers and comments. Word runs as a visible window, every change appears instantly, and Ctrl+Z works end to end.

> Requires a valid Microsoft Word license. This project does not include, crack or distribute Office software.

## What it can do

| Scenario | Description |
|----------|-------------|
| Reports | Agent pulls data, writes, formats and saves in one flow |
| Contracts | Fill templates with variables, generate formal documents, export PDF |
| Papers | Structured writing, insert tables and charts, format citations |
| Testing | Auto-create Word documents in pipelines for verification |

## What it works with

| Ecosystem | Agents |
|-----------|--------|
| Claude | Claude Desktop, Claude Code |
| OpenAI | ChatGPT Desktop |
| IDE | Cursor, VS Code / GitHub Copilot, Trae |
| AI terminal | OpenCode |
| More | Any stdio MCP client |

## Why you need it

- **81 tools** for text, tables, charts, images, headers, footers, comments and bookmarks.
- **Streaming Markdown** with `start -> block -> end`: instant preview and automatic style inheritance.
- **Attach and reuse** an existing Word window and the document you already have open.
- **Automatic cursor management**: header, footer, table and marker operations restore the cursor position.
- **Ctrl+Z friendly**: native COM operations, so an unwanted change can be undone.
- **Security built in**: path sanitization, macros disabled, Zod validation, rate limiting and audit logging.
- **Stable watchdog**: the child process sends an IPC heartbeat and is not restarted while idle.

## Requirements

| Requirement | Notes |
|-------------|-------|
| Windows 10/11 (x64) | COM automation is Windows-only. |
| Node.js 20 or newer | Node 20 or 22 LTS is recommended. |
| npm 9 or newer | Bundled with Node.js. |
| Microsoft Word (desktop) | Microsoft 365, Word 2019/2021 or newer, with a valid license. Word for the web is not supported. |
| Visual Studio 2022 Build Tools | Select the **Desktop development with C++** workload, MSVC v143 and the Windows 10/11 SDK. Required to compile the `winax` native addon. |
| Python 3.10 or newer | Required by `node-gyp` during `npm install`. |
| Git for Windows | Required to clone the repository. You can also download the ZIP from GitHub. |

## Quick start

### 1. Clone the repository

```powershell
git clone https://github.com/alejandr003/word-mcp-server.git
cd word-mcp-server
```

Use a short local path such as `C:\dev\word-mcp-server`. Avoid OneDrive-synced folders: OneDrive and antivirus tools can lock `node_modules` and `build` while npm or TypeScript is writing to them.

If you do not have Git, open the repository on GitHub, choose **Code > Download ZIP**, extract it to `C:\dev\word-mcp-server`, and continue with step 2.

### 2. Install dependencies

```powershell
npm ci
```

`npm ci` performs a clean, reproducible install from `package-lock.json`. If it fails because the lockfile is out of sync, use:

```powershell
npm install
```

Do not add `--production`: the build needs TypeScript and the other dev dependencies.

### 3. Build the project

```powershell
npm run build
```

The compiler is quiet on success. Verify that the entry point exists:

```powershell
Test-Path .\build\parent.js
```

The expected result is `True`.

### 4. Configure your MCP client

Use an **absolute path** to `build\parent.js`. In JSON, escape backslashes as `\\`.

**OpenCode** (`%USERPROFILE%\.config\opencode\opencode.json`):

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "word": {
      "type": "local",
      "command": ["node", "C:\\dev\\word-mcp-server\\build\\parent.js"],
      "enabled": true
    }
  }
}
```

**Claude Desktop** (`%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "word": {
      "command": "node",
      "args": ["C:\\dev\\word-mcp-server\\build\\parent.js"]
    }
  }
}
```

Other stdio MCP clients use the same `command` + `args` shape as the Claude Desktop example.

Restart the client after changing its configuration.

### 5. Test the server

```powershell
node .\build\parent.js
```

The process stays running and waits for MCP messages on stdio. No output is normal. Press `Ctrl+C` to stop it. If it exits immediately, see the troubleshooting section below.

### 6. Use it

Open Microsoft Word (or let the server open it) and ask your agent:

> "Write a weekly report in Word."

## Troubleshooting

### Clone issues

| Problem | What to do |
|---------|------------|
| `git` is not recognized | Install [Git for Windows](https://git-scm.com/download/win), then reopen the terminal. |
| `Repository not found` | Check the URL. This repository is public, so cloning over HTTPS does not require a login. For a private fork, authenticate with a personal access token or SSH. |
| SSL certificate or proxy error | Configure Git for your proxy, or download the ZIP from GitHub. For a corporate certificate, configure the trusted CA instead of disabling SSL verification. |
| `Filename too long` | Run `git config --system core.longpaths true` in an elevated terminal, then clone into `C:\dev`. |
| Clone hangs or is very slow | Try a shallow clone: `git clone --depth 1 https://github.com/alejandr003/word-mcp-server.git`. Check your network and proxy settings. |
| `Permission denied` or files locked | Clone outside OneDrive and add the folder to your antivirus exclusions. |
| Line-ending warnings | Run `git config --global core.autocrlf true` and clone again. |

### `npm install` / `npm ci` issues

| Problem | What to do |
|---------|------------|
| `gyp ERR! find VS` or `winax` build failure | Install Visual Studio 2022 Build Tools with **Desktop development with C++**, MSVC v143 and the Windows 10/11 SDK. Reboot, then run `npm config set msvs_version 2022` and reinstall. |
| Python is not found | Install Python 3.10+ and make sure `python --version` works. If needed: `npm config set python "C:\Path\To\python.exe"`. |
| `EBADENGINE` or unsupported Node version | Run `node -v`. Install Node 20 or 22 LTS; older versions are not supported. |
| `EACCES` or `EPERM` | Close Word and any running MCP server, use a short path, and exclude the folder from antivirus/OneDrive. |
| `npm ci` reports a lockfile mismatch | Run `npm install`, or restore `package-lock.json` from the repository. |
| `winax` cannot be found at runtime | Run `npm rebuild winax`. Make sure you are using 64-bit Node. |
| Registry or proxy problems | Run `npm config set registry https://registry.npmjs.org/` and configure `npm config set proxy` / `https-proxy` if your network requires it. |
| Start from a clean state | `Remove-Item -Recurse -Force node_modules` then `npm cache clean --force` and `npm ci`. |

### `npm run build` issues

| Problem | What to do |
|---------|------------|
| `'tsc' is not recognized` | Dependencies are missing. Run `npm install` first, from the repository root (the folder containing `package.json`). Check with `npx tsc --version`. |
| `Cannot find module 'winax'` | Rebuild the native addon with `npm rebuild winax`. If it still fails, reinstall with the C++ build tools present. |
| `EPERM` while writing `build/` | Close Word and the MCP server, delete `build/`, and rebuild. Exclude the folder from antivirus and OneDrive. |
| Type errors after pulling changes | Run `npm ci`, confirm `node -v` is 20+, and use `npx tsc --noEmit` to see the full diagnostics. |
| Build succeeds but `build/parent.js` is missing | Run the build from the repository root and check `tsconfig.json` (`outDir` is `./build`). List the output with `Get-ChildItem .\build`. |
| `TS2688: Cannot find type definition file for 'node'` | Dev dependencies were skipped. Do not use `npm install --production`; run a normal `npm ci`. |
| `npm run build` prints nothing | This is normal: `tsc` is quiet when it succeeds. Check the exit code and `Test-Path .\build\parent.js`. |
| Build on macOS or Linux | The build may run, but Word COM automation only works on Windows. Use Windows for the actual server. |

### Runtime and MCP client issues

| Problem | What to do |
|---------|------------|
| The client does not list the `word` server | Validate the JSON, use an absolute path with `\\`, and restart the client. Test `node .\build\parent.js` manually. |
| `Word.Application` or COM error | Open Word once manually, confirm it is licensed, and repair Office if needed. |
| A new Word window opens every time | Make sure the client points to **this fork's** `build\parent.js`, run `npm run build` again, and restart the client. The fork attaches to the existing `WINWORD.EXE`. |
| Document is locked or `EPERM` | Close the file in Word or copy it to a local folder. OneDrive can lock files while syncing. Set `ALLOWED_DIRECTORIES` in `.env` if the path is outside the allowed folders. |
| `No document is open` | Call `word_document` or `word_open` before editing. |
| Changes are not saved | Call `word_save` and verify with `word_get_info` (`saved: true`). |
| The child restarts every ~35 seconds | Update to this fork and rebuild. The IPC heartbeat fix is included here. |
| Need more logs | Set `LOG_LEVEL=debug` in `.env`. Logs are written to `stderr`, so they do not corrupt the MCP stream. Check the client's MCP logs too. |

## Configuration

Configuration is optional. To customize it:

```powershell
Copy-Item .env.example .env
```

| Variable | Description |
|----------|-------------|
| `ALLOWED_DIRECTORIES` | Semicolon-separated absolute paths the server may access. Empty means unrestricted. |
| `ALLOW_NETWORK_PATHS` | Allow UNC paths such as `\\server\share`. |
| `MAX_FILE_SIZE` | Maximum file size in bytes. |
| `MAX_TEXT_LENGTH` | Maximum text length in characters. |
| `ALLOW_MACROS` | Keep this `false` unless you really need macros. |
| `OPERATION_TIMEOUT_MS` | Timeout for a single operation, in milliseconds. |
| `RATE_LIMIT_WINDOW_MS` / `RATE_LIMIT_MAX_CALLS` | Rate-limit window and maximum calls per window. |
| `WATCHDOG_TIMEOUT_MS` / `WATCHDOG_INTERVAL_MS` | Watchdog silence timeout and health-check interval. |
| `CHART_OP_TIMEOUT` / `CHART_WORKER_IDLE_TIMEOUT` | Chart operation and chart worker idle timeouts. |
| `LOG_LEVEL` | `fatal`, `error`, `warn`, `info`, `debug` or `trace`. |

## Development

```powershell
npm run dev       # tsc --watch
npm test          # unit tests (integration tests are excluded)
npx tsc --noEmit  # type-check only
```

The integration tests require Microsoft Word. The regular `npm test` command uses mocks and does not need Word.

## Project links

- [TOOLS.md](./TOOLS.md) - tool reference
- [CONTEXT.md](./CONTEXT.md) - architecture and design
- [.env.example](./.env.example) - configuration template
- [LICENSE](./LICENSE) - MIT License
- [Upstream project](https://github.com/HelloWorld-Open/word-mcp-server)

## License

MIT. See [LICENSE](./LICENSE).
