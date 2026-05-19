# excel-reader-mcp

Stdio MCP server for reading `.xlsx` files (openpyxl).

## Install

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -e .
```

Or from a Git URL (after you publish the repo):

```bash
pip install "git+https://github.com/vojtechovsky/excel-reader-mcp.git"
```

## Cursor `mcp.json`

The goal is for Cursor to run the same Python / virtual environment where you installed the package — ideally using the full path to `excel-reader-mcp.exe`.

1. Find the executable path  
In the same environment where you ran `pip install`:

```powershell
where.exe excel-reader-mcp
```

The executable is in the `Scripts` folder, e.g.:  
`...\\Python312\\Scripts\\excel-reader-mcp.exe`

2. Add the server in Cursor  

Option A: Settings → MCP → add server / Edit config  


```json
{
  "mcpServers": {
    "excel-reader": {
      "command": "C:/Users/username/AppData/Local/Programs/Python/Python312/Scripts/excel-reader-mcp.exe",
      "args": []
    }
  }
}
```
Full path is recommended.

3. Reload  
Refresh/reload MCP or restart Cursor to apply changes.


## Tools

- `read_excel` — all sheets as `sheets`: `{ sheetName: rows[][] }`
- `read_excel_by_sheet_name` — one sheet (optional `sheet_name`, default first sheet)
- `read_excel_by_sheet_index` — one sheet by index (optional `sheet_index`, default `0`)
