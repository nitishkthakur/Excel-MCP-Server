# Excel MCP Server

A [FastMCP](https://github.com/jlowin/fastmcp) server that lets any MCP-compatible AI client (Claude Desktop, Cursor, etc.) read and explore Excel (`.xlsx`) files through three focused tools.

---

## Tools

### `list_sheets`
Lists every sheet name in a workbook, in order.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_path` | `string` | No | Path to the `.xlsx` file. Falls back to the `EXCEL_FILE` env var, then `data.xlsx`. |

**Example response**
```json
{
  "file": "sales_data.xlsx",
  "sheets": ["Jan", "Feb", "Mar", "Summary"]
}
```

---

### `explore_excel`
Shows the dimensions (rows × columns) of every sheet so you can plan how to read the data.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_path` | `string` | No | Path to the `.xlsx` file. |

**Example response**
```json
{
  "file": "sales_data.xlsx",
  "sheets": [
    { "name": "Jan",     "rows": 32, "columns": 5, "data_rows": 31 },
    { "name": "Summary", "rows":  5, "columns": 3, "data_rows":  4 }
  ]
}
```

---

### `retriever`
Fetches rows from a specific sheet, or from **every** sheet at once.  
The tool **auto-detects the number of columns** from each sheet's header row.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_path` | `string` | No | Path to the `.xlsx` file. |
| `sheet_name` | `string` | No | Name of the sheet to read. Omit to read **all** sheets. |
| `n_rows` | `integer \| null` | No | Maximum data rows per sheet (default `10`). Pass `null` for all rows. |

**Example – specific sheet, first 5 rows**
```json
// Request
{ "file_path": "sales_data.xlsx", "sheet_name": "Jan", "n_rows": 5 }

// Response
{
  "file": "sales_data.xlsx",
  "sheets": [
    {
      "name": "Jan",
      "columns": 5,
      "rows_returned": 5,
      "data": [
        { "Date": "2024-01-01", "Region": "North", "Product": "Widget", "Qty": 10, "Amount": 250 },
        ...
      ]
    }
  ]
}
```

**Example – all sheets, 3 rows each**
```json
// Request
{ "file_path": "sales_data.xlsx", "n_rows": 3 }

// Response
{
  "file": "sales_data.xlsx",
  "sheets": [
    { "name": "Jan",     "columns": 5, "rows_returned": 3, "data": [...] },
    { "name": "Feb",     "columns": 5, "rows_returned": 3, "data": [...] },
    { "name": "Summary", "columns": 3, "rows_returned": 3, "data": [...] }
  ]
}
```

---

## Installation

```bash
pip install -r requirements.txt
```

## Running the server

**stdio (default – works with Claude Desktop, etc.)**
```bash
python server.py
```

**Specify the Excel file via environment variable**
```bash
EXCEL_FILE=/path/to/workbook.xlsx python server.py
```

## Running tests

```bash
pip install pytest
pytest tests/
```

## Claude Desktop configuration

Add the following block to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["/absolute/path/to/server.py"],
      "env": {
        "EXCEL_FILE": "/absolute/path/to/your/workbook.xlsx"
      }
    }
  }
}
```