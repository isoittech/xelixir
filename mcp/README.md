# xelixir

Local MCP server for Excel operations backed by Java tools in the xelixir repository.

This package is not intended for publishing; it exists only to provide an MCP
server entry point using `uv run` as configured in [`mcp-config.json`](mcp/mcp-config.json:1).

## Running the server with uv

The Excel MCP server is designed to be run with `uv` from the `py` directory:

```bash
cd /path/to/xelixir/py
uv run xelixir -t stdio
```

You can also use HTTP or SSE transports for local testing:

```bash
# HTTP (streamable-http) on port 8000
uv run xelixir -t http -p 8000

# Server-Sent Events (SSE) on port 8001
uv run xelixir -t sse -p 8001
```

> Note: MCP runtimes like Claude Desktop will typically use the `stdio`
> transport and will manage the server process lifecycle themselves.

## Project layout (Python side)

This `py` directory is the Python half of the repository:

- Java implementation lives under [`../java`](../java/src/xelixir/ExcelUtils.java:1)
- Python MCP server entrypoint is [`xelixir.py`](xelixir.py:1)
- Excel wrappers around the Java tools live in [`src/excel/wrapper.py`](src/excel/wrapper.py:1)
- The Python package and script entrypoint are defined in [`pyproject.toml`](pyproject.toml:1)

The wrappers in [`src/excel/wrapper.py`](src/excel/wrapper.py:1) locate the repository root,
construct a Java classpath from `../java/dist` and `../java/jars`, and invoke
Java main classes such as `xelixir.ReadExcelTool`.

## Using the server from an MCP client (e.g. Claude)

1. Make sure Java is available and the Excel Java tools are built under
   `../java` (compiled classes in `../java/dist` and JARs in `../java/jars`).
2. From `xelixir/py`, install dependencies and build the virtual
   environment (the first `uv run` will do this automatically based on
   [`pyproject.toml`](pyproject.toml:1)).
3. Configure your MCP client to use the [`mcp-config.json`](mcp-config.json:1) file in this
   directory. For example, Claude Desktop can be pointed at this
   directory so that it discovers the `xelixir` and uses:

   ```json
   {
     "mcpServers": {
       "xelixir": {
         "command": "uv",
         "args": [
           "run",
           "xelixir"
         ],
         "env": {
           "PYTHONPATH": "./src"
         }
       }
     }
   }
   ```

4. Restart the MCP client so it picks up the new server configuration.
5. In the client UI, you should now see an `xelixir` toolset
   that exposes operations such as creating workbooks, reading/writing
   ranges, managing sheets, listing sheet names, and creating charts or pivot tables.

## File sharing & environment variables (Docker / SSE)

When running under Docker (for example via [`docker-compose.yml`](docker-compose.yml:1)),
this server follows a common pattern:

- A **shared directory inside the container** (default `/mnt/data` or `WORKSPACE_DIR`) is
  used for all Excel file reads/writes.
- Tool results include a **`download_url`** that points to `/files/...` so that
  end users can download generated files from a browser.

Key environment variables:

- `WORKSPACE_DIR` (container)
  - Directory exposed under `/files` in SSE mode.
  - Defaults to `/mnt/data` in [`xelixir.py`](xelixir.py:76) if not set.
- `MCP_SHARED_DIR` (host, used by Compose)
  - Host directory mounted into `/mnt/data` inside the container.
  - See the `volumes` section in [`docker-compose.yml`](docker-compose.yml:1).
- `EXCEL_PUBLIC_BASE_URL` (container)
  - Public base URL used to build download URLs.
  - Example values: `https://your-domain.example` or `http://your-host:8585`.

If `EXCEL_PUBLIC_BASE_URL` is not configured, tools that need to return a
`download_url` (such as `tool_create_excel` or `tool_write_excel`) will raise
an error instead of returning a container-only file path.

## Development notes

- This package is **not** intended to be published to PyPI; it is tied to the
  surrounding repository layout (Java sources and JARs).
- When you add a new Java tool under [`../java/src/xelixir`](../java/src/xelixir/ExcelUtils.java:1):
  - Add a corresponding wrapper function to [`src/excel/wrapper.py`](src/excel/wrapper.py:1)
  - Register a new MCP tool in [`xelixir.py`](xelixir.py:1)
- Keep the behavior described in the top-level [`README.md`](../README.md:1)
  and [`README.JA.md`](../README.JA.md:1) in sync with the actual Python implementation.

If the server fails to start, check the client logs for errors about
Java, the classpath, or missing Python dependencies, and verify that you
can run:

```bash
cd /path/to/xelixir/py
uv run xelixir --help
```

successfully from a terminal.
