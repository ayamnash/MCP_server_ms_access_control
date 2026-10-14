# MCP_server_ms_access_control pro
To get the the pro version  send me an email

# MS Access Enterprise MCP Server & Knowledge Graph Engine

> **The Definitive Model Context Protocol (MCP) Server for Microsoft Access (`.accdb` and `.mdb`).**  
> Connect Claude Desktop, Cursor, Kiro, VS Code, or custom AI agents directly to your Microsoft Access databases with native speed, token reduction, surgical UI editing, deep impact analysis, and referential integrity.


## Overview

Microsoft Access remains one of the world's most widely deployed database engines for small-to-medium businesses, internal tooling, and departmental software. However, connecting modern Large Language Models (LLMs) to Access databases has historically been difficult, unreliable, and cost-prohibitive.

This project delivers a **full-featured, 100% local Model Context Protocol (MCP) server** that bridges Microsoft Access directly to any AI assistant supporting the MCP standard. It does not simply run raw SQL queries; it acts as a comprehensive **reverse-engineering platform, knowledge graph engine, and database administrator** built specifically for autonomous AI workflows.

---

## Why This Server Exists

Traditional database MCP servers suffer from three major flaws when dealing with MS Access:

1. **Catastrophic Token Inefficiency:** When asked to inspect a database, naive tools dump 40+ tables, hundreds of columns, and thousands of metadata lines into the LLM prompt. This burns **30,000 to 50,000 tokens** (costing dollars per question) and overwhelms the AI context window.
2. **Blindness to Access UI & Logic:** In MS Access, business logic lives inside Forms, Reports, Event Procedures, and embedded VBA—not just tables. Existing tools cannot read or modify an existing form without deleting it and attempting to reconstruct thousands of lines of raw text.
3. **Database Fragility & Corruption:** Frequent edits by AI can cause `.accdb` files to bloat up to the 2GB limit or corrupt indexes. Without referential integrity, automated table generation leaves behind orphaned records.

**This server solves all three problems:**
- **Pre-indexes your database** into an entity relationship graph locally in under 1 second.
- **Reduces token consumption by 98.5%**, supplying ~400 tokens of surgical context instead of 35,000.
- **Inspects and surgically patches** existing forms and buttons without rebuilding them.
- **Enforces true referential integrity** with cascade updates and deletes via native DAO.
- **Provides database doctoring and atomic rollback snapshots** before any risky changes.

---

## Key Highlights & New Features

### 🌟 New Advanced Capabilities

1. **UI Reverse-Engineering & Surgical Patching (`inspect_form_or_report`, `patch_form_control`):**
   - Reads any pre-existing form or report via hidden Access COM `SaveAsText` without disturbing the database.
   - Extracts controls, events (`OnClick`, `BeforeUpdate`), and data bindings into a structured summary.
   - Modifies specific properties on a button or text box in memory and reloads via `LoadFromText` in fractions of a second—no form reconstruction required!

2. **Form Event Injector (`attach_control_event_vba`):**
   - Access is an event-driven RAD platform. This tool links any control's event property (`OnClick = [Event Procedure]`) and automatically injects the corresponding `Sub btn_Click()` code directly into the Form's code-behind module (`Form_<Name>`) with automatic syntax compilation checking.

3. **DAO Referential Integrity Engine (`manage_relationship`):**
   - Normal ODBC `CREATE TABLE` statements do not create visible relationships in Access. This tool uses in-process `DAO.DBEngine` to physically link tables in the Access Relationships window with full support for `Cascade Update` and `Cascade Delete`.

4. **Deep Dependency & Change Impact Analysis (`analyze_change_impact`):**
   - Answers: *"If I rename or delete field X, what breaks?"*
   - Scans foreign keys, queries (including wildcard `SELECT *`), Form/Report controls, and the exact line numbers in all VBA modules in under 0.01 seconds.

5. **Database Doctor & Atomic Snapshots (`compact_and_repair_database`, `create_restore_point`, `restore_database_snapshot`):**
   - Automatically reclaims up to 80% wasted space and repairs internal indexes using `DAO.CompactDatabase`.
   - Creates point-in-time timestamped backups in `.mcp_snapshots/` before risky operations. If an AI makes a mistake, revert the database with a single command.

6. **Migration Engine & Direct PDF Export (`generate_migration_bundle`, `export_report_to_pdf`):**
   - Generates production-ready DDL schema files for **PostgreSQL**, **SQLite**, or **DuckDB**, ordered topologically (parents before children) with converted Views.
   - Runs any Access report in the background and exports it directly to a clean PDF ready for sharing.

---

## Token Savings & Architecture

### Traditional Method vs. Knowledge Graph Subgraph

```
[Traditional MCP Approach]
User Prompt ──> AI reads 30+ tables & 600+ fields ──> 35,000 - 50,000 Tokens Consumed ($$$)
                                                      AI Context Window Polluted

[Our Knowledge Graph & Surgical Engine]
User Prompt ──> AI calls MCP Tool ──> Local BFS on graph (0.005s) ──> Returns ~400 Tokens ($0.001)
                                                                       AI Context Clean & Focused
```

### Real-World Token Benchmark

| Operation | Traditional Approach | Our Engine | Savings |
|:---|:---:|:---:|:---:|
| **Find shortest JOIN between two tables** | ~28,000 tokens | **~420 tokens** | **98.5%** |
| **Inspect target tables for insert** | ~35,000 tokens | **~540 tokens** | **98.4%** |
| **Modify a button on an existing form** | ~3,500 tokens (rebuild form) | **~120 tokens** (patch property) | **96.5%** |
| **Check impact of deleting a column** | Manual search across all objects | **~380 tokens** (instant report) | **99.0%** |
| **Credit cost in Claude / Kiro / Cursor** | 10 to 30 credits | **0.20 to 0.45 credits** | **Up to 98.8%** |

---

## Step-by-Step Installation & Setup

### Prerequisites
- **Operating System:** Windows 10 or Windows 11 (MS Access COM and ODBC drivers run on Windows).
- **Python Version:** Python 3.10, 3.11, or 3.12 (Python 3.11 32-bit recommended if using 32-bit Office).
- **Access Driver:** Microsoft Access installed OR the free [Microsoft Access Database Engine 2016 Redistributable](https://www.microsoft.com/en-us/download/details.aspx?id=54920).

---

### Understanding 32-bit vs. 64-bit Architecture
> [!IMPORTANT]
> **This is the #1 mistake new users make.** In Windows, your Python architecture **must match** your Microsoft Access / Microsoft Office architecture:
> - If your Microsoft Office / Access is **32-bit**, you **must** run this server with **32-bit Python**.
> - If your Microsoft Office / Access is **64-bit**, you **must** run this server with **64-bit Python**.

#### How to Check Your Office Architecture:
1. Open Access (or Excel/Word).
2. Go to **File → Account → About Access**.
3. Look at the top line to see whether it says **32-bit** or **64-bit**.

#### How to Check Your Python Architecture:
Open PowerShell or CMD and run:
```powershell
python -c "import platform; print(platform.architecture())"
```
If it prints `('32bit', 'WindowsPE')`, you have 32-bit Python. If it prints `('64bit', 'WindowsPE')`, you have 64-bit Python.

---

### Installation Commands

#### 1. Clone or Download This Repository
```powershell
cd C:\Projects
git clone https://github.com/ayamnash/MCP_server_ms_access_control.git mcp_access
cd mcp_access
```

#### 2. Install Required Python Packages
We recommend `fastmcp==2.11.3` for proven stability with JSON-RPC stdio:
```powershell
pip install "fastmcp==2.11.3" pyodbc pywin32
```

#### 3. Verify Your Setup
Run this one-line test to verify that the server loads all 38 tools and detects your ODBC driver:
```powershell
python -c "import server; print('Loaded tools:', len(server.mcp._tool_manager.list_tools()))"
```
You should see:
```
Loaded tools: 38
```

---

### Connecting to Claude Desktop, Cursor, and Other Clients

#### 1. Claude Desktop Setup
Open your Claude Desktop configuration file located at:  
`%APPDATA%\Claude\claude_desktop_config.json`

Add the server to the `mcpServers` section:
```json
{
  "mcpServers": {
    "msaccess": {
      "command": "C:\\Users\\YourUsername\\AppData\\Local\\Programs\\Python\\Python311-32\\python.exe",
      "args": [
        "K:\\mcp_new_feature\\server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```
*(Replace the Python path and `server.py` path with the exact paths on your computer).*

#### 2. Cursor IDE Setup
1. In Cursor, open **Settings → Features → MCP**.
2. Click **Add New MCP Server**.
3. Set **Name:** `msaccess`
4. Set **Type:** `command`
5. Set **Command:** `python K:\mcp_new_feature\server.py`

#### 3. Kiro / VS Code / Other Clients
Add to your project's `.mcp/config.json`:
```json
{
  "mcpServers": {
    "msaccess": {
      "command": "python",
      "args": ["K:/mcp_new_feature/server.py"]
    }
  }
}
```

Restart your client. You will see all 38 tools active with green indicators!

---

## Extensive AI Prompt Collection

Here is a collection of tested prompts you can copy, paste, and adapt when chatting with your AI assistant.
Group1
Let's assume you have an accounting and warehouse database
 D:\accounting_system.accdb
**********Prompt***********
"On form frmSalesInvoice in D:\accounting_system.accdb, use patch_form_control to update the control btnSave with Caption='Post & Print Invoice' and BackColor=16744448."
**********Prompt***********
"Modify the root properties of form frmSalesInvoice in D:\accounting_system.accdb using patch_form_control with control_name='Form' to set Caption='Sales Invoice Entry System'."
**********Prompt***********
"Generate a form template for record source JournalEntries with form type 'single' in D:\accounting_system.accdb using generate_form_template."
**********Prompt***********
"Inspect the existing form frmSalesInvoice in D:\accounting_system.accdb using inspect_form_or_report. Extract its RecordSource, controls, data bindings, button events, and any embedded code behind the form."
**********Prompt***********
"Create a new form named frmQuickAccount in D:\accounting_system.accdb based on table ChartOfAccounts using create_form_from_llm_text."
"Attach a Click event to button btnPostVoucher on form frmJournalVoucher in D:\accounting_system.accdb using attach_control_event_vba. Set OnClick='[Event Procedure]' and inject the VBA code:

vba
Private Sub btnPostVoucher_Click()
    MsgBox \"Journal voucher posted successfully to General Ledger.\", vbInformation, \"Accounting System\"
End Sub
```"
**********Prompt***********"
List all standard, class, and form VBA modules in D:\accounting_system.accdb using list_vba_modules."
**********Prompt***********"
Read the VBA source code of module modTaxEngine in D:\accounting_system.accdb using read_vba_module."
**********Prompt***********"
Create a new VBA standard module named modVATCalculator in D:\accounting_system.accdb using write_vba_module with a public function CalculateVAT(amount As Double, rate As Double) As Double."
**********Prompt***********"
Execute the VBA function CalculateVAT in D:\accounting_system.accdb with arguments '1000, 0.15' using run_vba_function."
**********Prompt***********"
Delete the obsolete module modLegacyPayroll from D:\accounting_system.accdb using delete_vba_module."
**********Prompt***********"
Run the report rptTrialBalance in D:\accounting_system.accdb and export it directly to D:\TrialBalance_2026.pdf using export_report_to_pdf."
Checking the existence and path of the database
**********Prompt***********
"Export the report rptInvoicePrint in D:\accounting_system.accdb to D:\Invoice_1001.pdf using export_report_to_pdf with where_condition='[InvoiceID]=1001'."
**********Prompt***********
"Create a tabular report named rptStockSummary based on table WarehouseStock in D:\accounting_system.accdb using create_report_from_source."
**********Prompt***********
"Generate a report template for record source qry_IncomeStatement with report type 'tabular' in D:\accounting_system.accdb using generate_report_template."
**********Prompt***********
"Compile and create a report named rptExpenseBreakdown in D:\accounting_system.accdb using create_report_from_template."
**********Prompt***********"
Create a new empty Access database named D:\accounting_archive.accdb using create_database."
"Create a table named PettyCash in D:\accounting_system.accdb with schema 'TransactionID AUTOINCREMENT PRIMARY KEY, TransDate DATETIME, Description TEXT(200), Amount CURRENCY, ApprovedBy TEXT(50)' using create_table."
**********Prompt***********
"List all user tables and their record counts in D:\accounting_system.accdb using list_tables."
**********Prompt***********
"Insert two rows into table ChartOfAccounts in D:\accounting_system.accdb using insert_data: [{'AccountCode': '1010', 'AccountName': 'Main Cash Box', 'AccountType': 'Asset'}, {'AccountCode': '4010', 'AccountName': 'Sales Revenue', 'AccountType': 'Revenue'}]."
**********Prompt***********
"Run a query on D:\accounting_system.accdb using run_query to calculate the total balance of debit and credit transactions: 'SELECT SUM(Debit) AS TotalDebit, SUM(Credit) AS TotalCredit FROM JournalLines'."
**********Prompt***********
"Save a query named qry_LowStockAlert in D:\accounting_system.accdb using save_query with SQL 'SELECT ItemID, ItemName, QuantityOnHand, ReorderLevel FROM WarehouseStock WHERE QuantityOnHand <= ReorderLevel'."
**********Prompt***********
"List all saved queries in D:\accounting_system.accdb using list_saved_queries."
**********Prompt***********
Create a backup snapshot for D:\accounting_system.accdb tagged 'pre_year_end_closing' using create_restore_point."
**********Prompt***********
"Perform a compact and repair operation on D:\accounting_system.accdb using compact_and_repair_database. Report the initial size, final size, and reclaimed disk space."
**********Prompt***********
"Restore D:\accounting_system.accdb from the snapshot tagged 'pre_year_end_closing' using restore_database_snapshot."
**********Prompt***********
"Generate a complete PostgreSQL migration bundle for D:\accounting_system.accdb using generate_migration_bundle with target_dialect='postgresql'. Verify that parent tables are placed before child tables."
**********Prompt***********
"Start a batch operation on D:\accounting_system.accdb using begin_batch_operation, insert 3 sample items into WarehouseStock, and then finalize and save with commit_batch_operation." 
**********Prompt*********** 
"Force close any hanging MS Access instances using force_close_access."


**********Prompt***********
"Please check if the database D:\accounting_system.accdb exists and verify its resolved path using find_database."

**********Prompt***********
Graph construction and extraction of global metrics
"Index the database D:\accounting_system.accdb using export_graphify_graph. Tell me the total number of tables, relationships, saved queries, and identify the top hub tables (god nodes) in the accounting system."
**********Prompt***********
 token-reduced 
"I need to write a double-entry journal voucher. Use query_database_context on D:\accounting_system.accdb with task description 'Post journal voucher with debit and credit lines' and tables 'ChartOfAccounts,JournalEntries,JournalLines' to give me a compact schema cheat sheet."
**********Prompt***********
"Inspect the network neighborhood of the Invoices table in D:\accounting_system.accdb using get_table_neighborhood with depth 1 to see all connected parent and child tables."
**********Prompt***********
Find the shortest relationship
"Find the shortest relationship path between Customers and WarehouseStock in D:\accounting_system.accdb using find_relationship_path. Return the connecting tables and the exact MS Access nested JOIN SQL."
**********Prompt***********
"Generate an interactive 2D Canvas HTML visualizer for D:\accounting_system.accdb using generate_interactive_graph_html and tell me where the HTML file is saved."
**********Prompt***********
"Run analyze_change_impact on table ChartOfAccounts and field AccountCode in D:\accounting_system.accdb. What queries, forms, reports, and specific lines of VBA code will break if I rename this field?"
**********Prompt***********
"We plan to deprecate the table WarehouseStock in D:\accounting_system.accdb. Use analyze_change_impact on this table to evaluate the risk level and show all dependent objects."
**********Prompt***********
"Establish a foreign key relationship in D:\accounting_system.accdb using manage_relationship between parent table Invoices (InvoiceID) and child table InvoiceDetails (InvoiceID) with cascade_update=True and cascade_delete=True."

"Link parent table ChartOfAccounts (AccountID) to child table JournalLines (AccountID) in D:\accounting_system.accdb using manage_relationship with cascade_update=True."

------------------------------------------------------
-------------------------------------------------------
Group2
### 1. Exploration & Interactive Graph Prompts

#### Build & Visualize the Entire Database
> "Please inspect `K:\mcp\Northwind.accdb` using `export_graphify_graph`. Tell me the total number of tables, relationships, and queries, and identify the top hub tables in the database."

#### Generate an Interactive HTML Viewer
> "Generate a standalone interactive 2D graph viewer for `K:\mcp\Northwind.accdb` using `generate_interactive_graph_html`. Where was the HTML file saved so I can open it in my browser?"

#### Fast Token-Saving Subgraph Query
> "I want to write a SQL query to see customer orders and shipment details in `K:\mcp\Northwind.accdb`. Use `query_database_context` with tables `Customers,Orders,Shippers` to give me only the relevant schema cheat sheet."

---

### 2. Deep Impact & "What Breaks If..." Prompts

#### Field Modification Impact
> "I am considering renaming the `UnitPrice` column in table `Order Details` in `K:\mcp\Northwind.accdb`. Run `analyze_change_impact` to tell me what queries, forms, reports, and specific lines of VBA code will break if I proceed."

#### Table Deletion Impact
> "We want to deprecate the `Shippers` table. Use `analyze_change_impact` on table `Shippers` to check if any forms, foreign keys, or report controls depend on it."

---

### 3. UI Reverse-Engineering & Surgical Form Patching Prompts

#### Inspect Existing Form Layout & Controls
> "Use `inspect_form_or_report` to inspect the form `frmOrders` in `K:\mcp\Northwind.accdb`. List all buttons, text boxes, their data sources, and any existing event handlers."

#### Change a Button Caption Without Rebuilding
> "On form `frmOrders` in `K:\mcp\Northwind.accdb`, patch the control `btnSave` using `patch_form_control` to change its `Caption` to `'Save & Print Receipt'`."

#### Change Form Background Color & Root Properties
> "Use `patch_form_control` with `control_name='Form'` on `frmCustomers` in `K:\mcp\Northwind.accdb` to set `Caption='Customer Master Registry'` and `DefaultView=0`."

---

### 4. VBA Code & Event Procedure Prompts

#### Attach a Click Event to a Button
> "Attach a Click event to button `btnExportPDF` on form `frmReports` in `K:\mcp\Northwind.accdb`. Use `attach_control_event_vba` to set its `OnClick` property to `[Event Procedure]` and inject the Sub:
> ```vba
> Private Sub btnExportPDF_Click()
>     MsgBox \"Exporting report to PDF...\", vbInformation, \"Report Generator\"
> End Sub
> ```"

#### Read and Audit a VBA Module
> "Read the VBA module `modUtilities` in `K:\mcp\Northwind.accdb` using `read_vba_module` and explain what functions it provides."

---

### 5. Referential Integrity & Relationships Prompts

#### Create a Foreign Key with Cascade Delete
> "Establish a foreign key relationship between table `Categories` (`CategoryID`) and table `Products` (`CategoryID`) in `K:\mcp\Northwind.accdb` using `manage_relationship`. Make sure referential integrity is enforced with `cascade_update=True` and `cascade_delete=False`."

#### Link Parent-Child Tables
> "Link parent table `Customers` (`CustomerID`) to child table `Orders` (`CustomerID`) in `K:\mcp\SalesDB.accdb` with Cascade Update enabled using `manage_relationship`."

---

### 6. Maintenance, Snapshots & Rollback Prompts

#### Create a Safe Restore Point
> "Before we make any changes to the database structure, create a restore point for `K:\mcp\Northwind.accdb` with tag `'before_schema_update'` using `create_restore_point`."

#### Compact & Repair
> "The database file `K:\mcp\Northwind.accdb` has grown to 150MB after extensive testing. Run `compact_and_repair_database` to recover unused space and fix internal indexes. Tell me how many megabytes were saved."

#### Restore From a Previous Snapshot
> "Our last form test introduced an error. Revert `K:\mcp\Northwind.accdb` to the latest snapshot tagged `'before_schema_update'` using `restore_database_snapshot`."

---

### 7. Migration & PDF Export Prompts

#### Export Complete Database DDL to PostgreSQL
> "Generate a complete PostgreSQL DDL migration bundle for `K:\mcp\Northwind.accdb` using `generate_migration_bundle`. Ensure tables are ordered topologically to respect foreign key constraints, and convert saved queries to SQL views."

#### Migrate Schema to SQLite
> "Create a migration schema for `K:\mcp\Inventory.accdb` targeting `sqlite` using `generate_migration_bundle`. Show me the generated SQL."

#### Export a Report to PDF
> "Run the report `rptInvoice` in `K:\mcp\Northwind.accdb` for order ID 10248 and export it as a PDF to `C:\Invoices\Invoice_10248.pdf` using `export_report_to_pdf` with `where_condition='[OrderID]=10248'`."

---

### 8. SQL Queries & Multi-Table JOIN Prompts

#### Multi-Table INNER JOIN Path
> "Find the shortest relationship path between `Customers` and `Suppliers` in `K:\mcp\Northwind.accdb` using `find_relationship_path`. Give me the exact MS Access SQL clause with proper nested parentheses."

#### Query Execution
> "Run a query on `K:\mcp\Northwind.accdb` using `run_query` to display the top 5 highest value orders from table `[Order Details]`."

---

## Complete Reference of All 38 Tools

### 1. Knowledge Graph & Token Reduction Tools (5 Tools)

#### `export_graphify_graph(db_name: str, output_path: str = None) -> dict`
Extracts the database schema, foreign keys, queries, forms, and modules into a persistent Graphify-standard node-link graph (`graph.json`).
- **`db_name`**: Path or name of the database.
- **`output_path`**: Destination JSON file path (default: `graphify-out/graph.json`).

#### `query_database_context(db_name: str, task_description: str, relevant_tables: str = "", depth: int = 1, token_budget: int = 1500) -> str`
The core token-reduction tool. Traverses the knowledge graph using keyword relevance and BFS to return a compact Markdown schema cheat sheet (~400 tokens) strictly tailored to the user's task.
- **`task_description`**: What the AI plans to accomplish.
- **`relevant_tables`**: Optional comma-separated table names to prioritize.
- **`depth`**: Hop distance for related foreign key tables (default: 1).
- **`token_budget`**: Maximum characters/tokens for schema context.

#### `find_relationship_path(db_name: str, source_table: str, target_table: str) -> dict`
Finds the shortest foreign key path between any two tables using BFS and generates the exact, Access-compliant nested `INNER JOIN` SQL clause.
- **`source_table`**: Starting table name (e.g., `'Customers'`).
- **`target_table`**: Target table name (e.g., `'Products'`).

#### `get_table_neighborhood(db_name: str, table_name: str, depth: int = 1) -> str`
Returns a compact summary of a table's immediate network, including parent foreign keys, child tables, and connected queries.
- **`table_name`**: Table to inspect.
- **`depth`**: Traversal depth (default: 1).

#### `generate_interactive_graph_html(db_name: str, output_file: str = "database_graph.html") -> str`
Generates a standalone, dependency-free interactive HTML visualization file with 2D Canvas force simulation, zoom, pan, search, and full schema drawers.
- **`output_file`**: Output HTML file path.

---

### 2. UI Reverse-Engineering & Form Patching Tools (4 Tools)

#### `inspect_form_or_report(db_name: str, object_name: str, object_type: str = "form") -> dict`
Reverse-engineers an existing Access Form or Report using hidden Access COM `SaveAsText`. Extracts controls, types, captions, data bindings (`RecordSource`, `ControlSource`, `RowSource`), events, and embedded code behind the form.
- **`object_name`**: Name of the form or report (e.g. `'frmOrders'`).
- **`object_type`**: `'form'` or `'report'` (default: `'form'`).

#### `patch_form_control(db_name: str, form_name: str, control_name: str, properties: dict) -> str`
Surgically modifies or adds properties to a specific control on an existing form (e.g., `Caption`, `BackColor`, `ControlSource`, or `OnClick`) via `SaveAsText -> Patch -> LoadFromText` without destroying the form.
- **`form_name`**: Name of the form.
- **`control_name`**: Control name (or `'Form'` to edit root form properties).
- **`properties`**: Key-value dictionary of properties to apply.

#### `generate_form_template(db_name: str, record_source: str, form_type: str, subform_object_name: str = None, link_master_field: str = None, link_child_field: str = None) -> str`
Generates an Access Form design text template for new form creation. Supports `'single'`, `'subform'`, and `'main'` master-detail forms.

#### `create_form_from_llm_text(db_name: str, form_name: str, form_text: str) -> str`
Compiles and loads a full form definition text into an Access database using `LoadFromText`. Automatically injects unique GUIDs and NameMaps.

---

### 3. Change Impact & Migration Tools (2 Tools)

#### `analyze_change_impact(db_name: str, target_object: str, target_field: str = None) -> dict`
Performs deep dependency analysis to predict what will break before you rename or delete an object or field. Checks foreign key relations, saved queries (including `SELECT *`), Form/Report controls, and exact line numbers in VBA modules.
- **`target_object`**: Table, query, form, or report name.
- **`target_field`**: Optional field name inside the target object.

#### `generate_migration_bundle(db_name: str, target_dialect: str = "postgresql") -> dict`
Generates a complete, dependency-ordered SQL DDL migration script (using Topological Sort) for migrating Access tables, primary keys, foreign keys, and saved query views to modern engines.
- **`target_dialect`**: `'postgresql'`, `'sqlite'`, or `'duckdb'`.

---

### 4. Referential Integrity & Relationship Engine (1 Tool)

#### `manage_relationship(db_name: str, primary_table: str, foreign_table: str, primary_field: str, foreign_field: str, cascade_update: bool = True, cascade_delete: bool = False, relation_name: str = None) -> str`
Creates or updates a persistent foreign key relationship with true referential integrity in MS Access using headless `DAO.DBEngine`. Appears directly in the Access Relationships window.
- **`primary_table`**: Parent table name.
- **`foreign_table`**: Child table name.
- **`primary_field`**: Primary key column name.
- **`foreign_field`**: Foreign key column name.
- **`cascade_update`**: Automatically update child records when parent key changes (default: True).
- **`cascade_delete`**: Automatically delete child records when parent record is deleted (default: False).
- **`relation_name`**: Optional custom relation name.

---

### 5. Database Doctor, Snapshots & Rollback Tools (3 Tools)

#### `compact_and_repair_database(db_name: str) -> str`
Compacts and repairs the Access database file using headless DAO. Fixes internal index corruption and reduces bloated file sizes by up to 80%.
- **`db_name`**: Database path or name.

#### `create_restore_point(db_name: str, tag: str = "manual") -> str`
Creates an atomic, point-in-time backup snapshot of the database file in a `.mcp_snapshots` folder before executing major structural changes.
- **`tag`**: Short label for the backup (e.g. `'before_bulk_import'`).

#### `restore_database_snapshot(db_name: str, snapshot_identifier: str) -> str`
Rolls back the database to a previously saved snapshot file or tag with a single command.
- **`snapshot_identifier`**: Snapshot filename, path, or tag.

---

### 6. Report Generation & PDF Export Tools (4 Tools)

#### `export_report_to_pdf(db_name: str, report_name: str, output_pdf_path: str = None, where_condition: str = "") -> str`
Runs an Access Report in the background and exports it directly to a PDF file ready for printing or emailing.
- **`report_name`**: Name of the report object.
- **`output_pdf_path`**: Destination PDF path (defaults to `<report_name>.pdf` next to DB).
- **`where_condition`**: Optional SQL WHERE filter clause (e.g. `'[CustomerID]=5'`).

#### `create_report_from_source(db_name: str, record_source: str, report_name: str, report_type: str = "tabular") -> str`
Generates a complete Access report in a single step directly from a table or saved query. Supports `'tabular'` and `'columnar'` layouts.

#### `generate_report_template(db_name: str, record_source: str, report_type: str = "tabular") -> str`
Generates an editable text template for custom report creation.

#### `create_report_from_template(db_name: str, report_name: str, report_text: str) -> str`
Compiles and loads a custom report definition into the database using `LoadFromText`.

---

### 7. VBA Module & Event Programming Tools (6 Tools)

#### `attach_control_event_vba(db_name: str, form_name: str, control_name: str, event_name: str, vba_code: str) -> str`
Sets a control's event property to `[Event Procedure]` (e.g. `OnClick`) and automatically injects the corresponding event subroutine into the form's code module (`Form_<Name>`) with syntax compilation validation.
- **`form_name`**: Form name.
- **`control_name`**: Target control (e.g. `'btnSave'`).
- **`event_name`**: Event name (e.g. `'OnClick'`, `'BeforeUpdate'`).
- **`vba_code`**: Complete Sub code block.

#### `list_vba_modules(db_name: str) -> str`
Lists all Standard, Class, and Form code modules in the database.

#### `read_vba_module(db_name: str, module_name: str) -> str`
Reads the source code from a specific VBA module.

#### `write_vba_module(db_name: str, module_name: str, code: str) -> str`
Creates or replaces a Standard VBA module with the provided code and performs compilation error checking.

#### `delete_vba_module(db_name: str, module_name: str) -> str`
Deletes a Standard VBA module from the database.

#### `run_vba_function(db_name: str, function_name: str, args: str = "") -> str`
Executes a public VBA function in the database and returns its return value.

---

### 8. Database & Table Management Tools (6 Tools)

#### `create_database(db_name: str) -> str`
Creates a brand-new, empty `.accdb` database file using ADOX.

#### `list_tables(db_name: str) -> str`
Lists all non-system user tables in the database with their record counts.

#### `create_table(db_name: str, table_name: str, schema: str) -> str`
Creates a new table with typed columns and constraints. Automatically sanitizes schema definitions for Access compatibility.

#### `insert_data(db_name: str, table: str, rows: list[dict]) -> str`
Performs bulk parameterized insert operations into a table. Automatically escapes table and field names with brackets to prevent ODBC syntax errors.

#### `find_database(db_name: str) -> str`
Diagnostic utility that searches common locations (current directory, home directory, absolute paths) to locate where a database file exists.

#### `save_and_close_access_database() -> str`
Saves all pending changes and cleanly shuts down active Access COM application instances.

---

### 9. Query Execution & Management Tools (3 Tools)

#### `run_query(db_name: str, sql: str) -> str`
Executes any SQL statement (`SELECT`, `INSERT`, `UPDATE`, `DELETE`) via ODBC. Automatically fixes quotes, `#DATE#` literals, and multiple JOIN parentheses.

#### `save_query(db_name: str, query_name: str, sql: str) -> str`
Saves or overwrites a named query definition (`QueryDef`) in the database.

#### `list_saved_queries(db_name: str) -> str`
Lists all saved query definitions in the database using the graph engine.

---

### 10. Batch Operations & Transaction Control (4 Tools)

#### `begin_batch_operation(db_name: str) -> str`
Keeps a single Access COM session open across multiple consecutive commands, dramatically speeding up multi-step operations (like building several tables and forms).

#### `commit_batch_operation() -> str`
Saves all modifications made during the active batch session and closes the database.

#### `rollback_batch_operation() -> str`
Discards all unsaved changes in the batch session and terminates the Access process.

#### `force_close_access() -> dict`
Emergency cleanup tool that forcibly terminates any hanging `MSACCESS.EXE` processes and cleans up COM locks.

---



*Engineered with precision for autonomous AI agents, enterprise database administrators, and the global MCP community.*

