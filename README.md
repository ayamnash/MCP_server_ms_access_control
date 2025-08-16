# MCP Server ms_Access 🚀

A powerful Model Context Protocol (MCP) server that provides seamless integration with Microsoft Access databases. This server enables you to create, manage, and query Access databases through MCP-compatible applications like Kiro IDE.

# **Prompt Samples**


 (1)Create a Microsoft Access database named pos.accdb in this path F:\mcp_server_ms_access_control

 for a Point of Sale (POS) system with the following structure:

Database  Name pos.accdb

📦 Items Table:

ItemID: unique ID (AutoNumber)

ItemName: name of the item

ItemPrice: price per unit

ItemDescription: optional text

🔁 Transactions Table:

TransactionID: unique ID (AutoNumber)

ItemID: link to the Items table

TransactionType: either "Purchase" or "Sales"

Quantity: number of items

TransactionDate: date of transaction

💸 Expenses Table:

ExpenseID: unique ID (AutoNumber)

ExpenseType: type/category of expense

Amount: how much was spent

ExpenseDate: date of expense

---

Create and save four queries:

1. Sales Amount Between Two Dates

Calculate the total sales (item price × quantity) filtered by a start and end date.

2. Purchase Amount Between Two Dates

Calculate total purchases (item price × quantity) between two dates.

3. Sum of Items Sold Between Two Dates

Group by item name and calculate how many of each item was sold between two dates.

4 detail expense between two dates

---

Save the queries as:

qry_SalesAmount_BetweenDates

qry_PurchaseAmount_BetweenDates

qry_SumSoldItems_BetweenDates

qry_expense_details

fix Issue may Encountered & Fixed:
The only issue was with the Items table creation - the initial ItemDescription field size (500 characters) was too large for Access. I fixed this by reducing it to 255 characters, which is the standard maximum for Access text fields.

All queries use parameter prompts [Start Date] and [End Date] so when you run them in Access, you'll be prompted to enter the date range. The database is ready for use!


============================================

(2)using mcp server  to
Create a complete Laundry Management application in Microsoft Access name laundry_managemet1.accdb in this folder path 
F:\mcp_server_ms_access_control1.

Requirements:

Database Structure

Create all necessary tables with proper field names, data types, and primary/foreign keys.

Include at least these entities:
tables:-

Customers (CustomerID, Name, Phone, Address, etc.)

LaundryItems (ItemID, Description, PricePerUnit, etc.)

Orders (OrderID, CustomerID, OrderDate, DueDate, Status, etc.)

OrderDetails (OrderDetailID, OrderID, ItemID, Quantity, Subtotal, etc.)

Payments (PaymentID, OrderID, PaymentDate, Amount, PaymentMethod, etc.)

Queries:-

Create queries for:

Orders due today

Total sales per day/month

Unpaid orders

Customer order history

Forms:-

Customer management form (add, edit, delete).

Order entry form with subform for order details.

Payment entry form.

Order tracking form (view status, mark as completed).


==========================================================================

(3)using mcp server  to
Create a complete Laundry Management application in Microsoft Access name laundry_managemet1.accdb in this folder path 
F:\mcp_server_ms_access_control1.

Requirements:

Database Structure

Create all necessary tables with proper field names, data types, and primary/foreign keys.

Include at least these entities:
tables:-

Customers (CustomerID, Name, Phone, Address, etc.)

LaundryItems (ItemID, Description, PricePerUnit, etc.)

Orders (OrderID, CustomerID, OrderDate, DueDate, Status, etc.)

OrderDetails (OrderDetailID, OrderID, ItemID, Quantity, Subtotal, etc.)

Payments (PaymentID, OrderID, PaymentDate, Amount, PaymentMethod, etc.)

Queries:-

Create queries for:

Orders due today

Total sales per day/month

Unpaid orders

Customer order history

Forms:-

Customer management form (add, edit, delete).

Order entry form with subform for order details.

Payment entry form.

Order tracking form (view status, mark as completed).

📌 Features

[v1 features vedeo](https://www.youtube.com/watch?v=TplSweAx4XU)

[v2 features vedeo](https://www.youtube.com/watch?v=vtuiIgX98t4)

[v3 features vedeo](https://www.youtube.com/watch?v=2-KPeqXjBLw)
🎨 Form Creation Tools (v3 - NEW!)
📝 generate_form_template – Generate a text template for Access forms

🏗️ create_form_from_llm_text – Create Access forms from text definitions


🗃️ Database Structure Tools
🏗️ create_database – Create an empty Access .accdb database

🧱 create_table – Create a table with specified schema

📋 list_tables – List all tables in the database

📊 Data Management Tools
➕ insert_data – Insert rows into a table

🧮 run_query – Execute SQL queries (SELECT, UPDATE, DELETE, etc.)

🔎 Query Management Tools
💾 save_query – Save a named query inside the Access database

📄 list_saved_queries – List all saved queries in the database

📜 VBA Module Tools (v2)
📚 list_vba_modules – List all VBA modules in the Access database

📖 read_vba_module – Read the code from a specific VBA module

✍️ write_vba_module – Create or replace a VBA module with provided code

❌ delete_vba_module – Delete a VBA module from the database

🚀 run_vba_function – Execute a VBA function and return the result


✨ **Form Types Supported:**
- **Single Forms** – Standalone forms for data entry and viewing
- **Subforms** – Forms designed to be embedded in other forms (datasheet view)
- **Main Forms with Subforms** – Master-detail forms with embedded subforms and automatic linking

🔧 **Enhanced Tools (v3 Improvements):**
- **Improved Error Handling** – Better error messages and feedback for all operations
- **Enhanced Query Management** – Fixed parameter handling in saved queries
- **Optimized Form Generation** – Automatic GUID and NameMap generation for robust form creation
- **Better Field Validation** – Improved data type handling and field size validation

## Prerequisites

- **Windows Operating System** (required for Access integration)
- **Python 3.13+**
- **Microsoft Access Database Engine** It is recommended to use the 2016 version.(required - see installation guide below)
- **uv** package manager (recommended)

### ⚠️ Important: Bit Architecture Compatibility

**Python and Microsoft Access Database Engine must have the same bit architecture (32-bit or 64-bit).**
Microsoft Access Database Engine 2016 Redistributable

Choose:

AccessDatabaseEngine.exe → for 32-bit systems or 32-bit Office

AccessDatabaseEngine_X64.exe → for 64-bit Office

🧪 Summary:
Feature	                     2010 Engine	         2016 Engine
Compatibility	           Office 2010–2013	       Office 2010–2021
New Excel/Access support	 ❌ Limited	                ✅ Full
Future-proof	             ❌ No	                    ✅ Yes
Stability	                 ✅ Yes                   	✅ Yes
Bitness must match Office	 ✅ Yes	                    ✅ Yes

#### Check Your Python Architecture

Open your terminal (CMD or PowerShell) and run:

```bash
python -c "import platform; print(platform.architecture())"
```

This will show either:
- `('64bit', 'WindowsPE')` - You have 64-bit Python
- `('32bit', 'WindowsPE')` - You have 32-bit Python

#### Check Your Office/Excel Architecture

1. Open **Excel**
2. Click **File** tab
3. Choose **Account**
4. Click **About Excel**

You'll see something like:
```
Microsoft® Excel® 2016 MSO (Version 2506 Build 16.0.18925.20076) 32-bit
```
or
```
Microsoft® Excel® 2019 MSO (Version 2506 Build 16.0.18925.20076) 64-bit
```

The last part shows whether you have 32-bit or 64-bit Office.

### Installing Microsoft Access Database Engine

**You must install the Access Database Engine that matches your Python architecture:**

#### For 64-bit Python:
- Download: [Microsoft Access Database Engine 2016 Redistributable (64-bit)](https://www.microsoft.com/en-us/download/details.aspx?id=54920)
- File: `AccessDatabaseEngine_X64.exe`

#### For 32-bit Python:
- Download: [Microsoft Access Database Engine 2016 Redistributable (32-bit)](https://www.microsoft.com/en-us/download/details.aspx?id=54920)
- File: `AccessDatabaseEngine.exe`

#### Installation Notes:
- If you have Office 2016/2019 installed, you may need to use the `/quiet` parameter:
  ```bash
  AccessDatabaseEngine_X64.exe /quiet
  ```
- For Office 2016/2019 users: The database engine version should match your Office version
- You may need to run the installer as Administrator

## Installation

[Watch the video on YouTube](https://www.youtube.com/watch?v=RWWANlhPjZ4)




### Option 1: Using uv (Recommended)

First, install uv if you haven't already:

```bash
# Install uv using pip
pip install uv





```bash
# Clone the repository
git clone https://github.com/ayamnash/MCP_server_ms_access_control.git
cd MCP_server_ms_access_control

# Create virtual environment and install dependencies
uv venv
uv pip install -e .
```

### Option 2: Using pip

```bash
# Clone the repository
git clone https://github.com/ayamnash/MCP_server_ms_access_control.git
cd MCP_server_ms_access_control

# Create virtual environment
python -m venv .venv

# Activate virtual environment
# On Windows:
.venv\Scripts\activate

# Install dependencies
pip install -e .
```


## Configuration

### Kiro IDE Configuration

To use this MCP server with Kiro IDE, add the following configuration to your MCP settings:

#### Workspace Configuration (`.kiro/settings/mcp.json`)
LIKE AS 
```json
{
  "mcpServers": {
    "ms_access-database": {
      "command": "python",
      "args": [
        "f:\\mcp_server_ms_access_control\\server.py"
      ],
      "env": {
        "PYTHONPATH": "f:\\mcp_server_ms_access_control"
      },
      "disabled": false,
      "autoApprove": [
        "mcp_ms_access_database1_create_database",
        "mcp_ms_access_database1_create_table",
        "mcp_ms_access_database1_insert_data",
        "mcp_ms_access_database1_run_query",
        "mcp_ms_access_database1_list_tables",
        "mcp_ms_access_database1_save_query",
        "mcp_ms_access_database1_list_vba_modules",
        "mcp_ms_access_database1_read_vba_module",
        "mcp_ms_access_database1_write_vba_module",
        "mcp_ms_access_database1_delete_vba_module",
        "mcp_ms_access_database1_run_vba_function",
        "mcp_ms_access_database1_generate_form_template",
        "mcp_ms_access_database1_create_form_from_llm_text"
      ]
    }
  }
}

```
Visual studio code 
.vscode\mcp.json
```json
{
  "servers": {
    "ms_access-database1": {
      "command": "python",
      "args": ["f:\\mcp_server_ms_access_control1\\server.py"],
      "env": {
        "PYTHONPATH": "f:\\mcp_server_ms_access_control1"
      }
    }
  }
} 
```
### Desktop Application Usage

You can also run the server as a standalone application:

```bash
# Activate your virtual environment first
.venv\Scripts\activate

# Run the server
python server.py
```

## Available Tools

The MCP server provides the following tools:

### 🗄️ Database Management
- **`create_database(db_name: str)`** - Create a new Access database
- **`list_tables(db_name: str)`** - List all tables in a database

### 🏗️ Table Operations
- **`create_table(db_name: str, table_name: str, schema: str)`** - Create a new table
  - Example schema: `"ID INT PRIMARY KEY, Name TEXT(100), Age INT"`

### 📊 Data Operations
- **`insert_data(db_name: str, table: str, rows: list[dict])`** - Insert data into tables
  - Example: `[{'ID': 1, 'Name': 'John', 'Age': 30}]`
- **`run_query(db_name: str, sql: str)`** - Execute SQL queries (SELECT, UPDATE, DELETE, etc.)

### 💾 Query Management
- **`save_query(db_name: str, query_name: str, sql: str)`** - Save named queries
- **`list_saved_queries(db_name: str)`** - List all saved queries

### 📜 VBA Module Management (v2)
- **`list_vba_modules(db_name: str)`** - List all VBA modules in the Access database
- **`read_vba_module(db_name: str, module_name: str)`** - Read the code from a specific VBA module
- **`write_vba_module(db_name: str, module_name: str, code: str)`** - Create or replace a VBA module with provided code
- **`delete_vba_module(db_name: str, module_name: str)`** - Delete a VBA module from the database
- **`run_vba_function(db_name: str, function_name: str, args: str)`** - Execute a VBA function and return the result

### 🎨 Form Creation Tools (v3 - NEW!)
- **`generate_form_template(db_name: str, record_source: str, form_type: str, ...)`** - Generate a text template for Access forms
  - **form_type options:**
    - `"single"` - Standard standalone form
    - `"subform"` - Form designed for embedding (datasheet view)
    - `"main"` - Form that contains a subform with automatic linking
- **`create_form_from_llm_text(db_name: str, form_name: str, form_text: str)`** - Create Access forms from text definitions
  - Automatically generates GUIDs and NameMaps
  - Handles form validation and error correction
  - Supports complex form layouts with subforms

## Usage Examples

### Creating a Database and Table

```python
# Create a new database
create_database("my_library")

# Create a table
create_table("my_library", "books", "ID INT PRIMARY KEY, Title TEXT(255), Author TEXT(100), Year INT")

# Insert some data
insert_data("my_library", "books", [
    {"ID": 1, "Title": "Python Programming", "Author": "John Doe", "Year": 2023},
    {"ID": 2, "Title": "Database Design", "Author": "Jane Smith", "Year": 2022}
])
```

### Querying Data

```python
# Select all books
run_query("my_library", "SELECT * FROM books")

# Filter by year
run_query("my_library", "SELECT * FROM books WHERE Year > 2022")

# Save a frequently used query
save_query("my_library", "recent_books", "SELECT * FROM books WHERE Year > 2020")
```

### Creating Forms (v3 - NEW!)

```python
# Create a single form for data entry
generate_form_template("my_library", "books", "single")
create_form_from_llm_text("my_library", "books_form", form_template_text)

# Create a subform for embedding
generate_form_template("my_library", "authors", "subform")
create_form_from_llm_text("my_library", "authors_subform", subform_template_text)

# Create a main form with embedded subform
generate_form_template(
    "my_library", 
    "books", 
    "main",
    subform_object_name="Form.authors_subform",
    link_master_field="AuthorID",
    link_child_field="AuthorID"
)
create_form_from_llm_text("my_library", "books_with_authors", main_form_template_text)
```

### VBA Integration (v2)

```python
# List all VBA modules
list_vba_modules("my_library")

# Create a VBA function
vba_code = """
Public Function CalculateLateFee(DaysLate As Integer) As Currency
    If DaysLate <= 0 Then
        CalculateLateFee = 0
    ElseIf DaysLate <= 7 Then
        CalculateLateFee = 1.00
    Else
        CalculateLateFee = 1.00 + (DaysLate - 7) * 0.50
    End If
End Function
"""
write_vba_module("my_library", "LibraryFunctions", vba_code)

# Execute the VBA function
result = run_vba_function("my_library", "CalculateLateFee", "10")
```
import win32com.client
adox = win32com.client.Dispatch("ADOX.Catalog")
conn_string = f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={db_path};"
adox.Create(conn_string)  # This creates the .accdb file

ADOX and why it's more reliable than ODBC
ADOX (ActiveX Data Objects Extensions) is a Microsoft COM library specifically designed for database schema operations like creating databases and tables.

Why ADOX is better than ODBC for creating Access databases:

ODBC (Open Database Connectivity) is a general-purpose database interface that sometimes has registry access issues on Windows
ADOX uses Windows COM (Component Object Model) which has direct access to the Access database engine
ADOX bypasses the registry issues that cause the "Unable to open registry key" errors you were seeing
ADOX is Microsoft's recommended method for programmatically creating Access databases
Here's what happens in the code:


pyodbc 

Driver detection: The code uses pyodbc.drivers() to list available database drivers
Table creation and data operations: After ADOX creates the empty database file, pyodbc is used to:
Connect to the database
Create the  table
Insert sample data
Read data for verification


ADOX: Creates the empty .accdb file
pyodbc: Handles all the SQL operations (CREATE TABLE, INSERT, SELECT)
So the combination gives you the best of both worlds:

ADOX for reliable database file creation
pyodbc for standard SQL operations
This is why your script now works - it uses the most reliable method for each task instead of trying to do everything through ODBC alone.


## Troubleshooting

### Common Issues

1. **Access Driver Not Found**
   ```
   Exception: Access ODBC driver not found
   ```
   **Solution:**
   - Install Microsoft Access Database Engine 2016 Redistributable
   - **Critical:** Ensure the database engine matches your Python architecture (32-bit or 64-bit)
   - Check available drivers: `python -c "import pyodbc; print(pyodbc.drivers())"`

2. **Architecture Mismatch Error**
   ```
   [Microsoft][ODBC Driver Manager] The specified DSN contains an architecture mismatch
   ```
   **Solution:**
   - Your Python and Access Database Engine have different architectures
   - Check Python architecture: `python -c "import platform; print(platform.architecture())"`
   - Check Office architecture: Excel → File → Account → About Excel
   - Install matching Access Database Engine version

3. **Office 2016/2019 Installation Conflicts**
   ```
   You cannot install the 64-bit version of Microsoft Access Database Engine 2016 because you currently have 32-bit Office products installed
   ```
   **Solution:**
   - Use the `/quiet` parameter: `AccessDatabaseEngine_X64.exe /quiet`
   - Or uninstall existing Office, install database engine, then reinstall Office
   - Consider using the same architecture for both Python and Office

4. **Permission Errors**
   - Run installer as Administrator
   - Check file permissions in the target directory
   - Ensure the database file location is writable

5. **Python Path Issues**
   - Ensure your virtual environment is activated
   - Verify all dependencies are installed: `pip list`
   - Check if pywin32 is properly installed: `python -c "import win32com.client"`

### Architecture Compatibility Quick Reference

| Your Setup | Python Architecture | Required Database Engine |
|------------|-------------------|-------------------------|
| 32-bit Office 2016/2019 | 32-bit Python | AccessDatabaseEngine.exe (32-bit) |
| 64-bit Office 2016/2019 | 64-bit Python | AccessDatabaseEngine_X64.exe (64-bit) |
| No Office installed | 32-bit Python | AccessDatabaseEngine.exe (32-bit) |
| No Office installed | 64-bit Python | AccessDatabaseEngine_X64.exe (64-bit) |

### System Requirements

- Windows 10/11
- Microsoft Access 2016+ or Access Database Engine
- Python 3.8 or higher
- At least 100MB free disk space

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- 📧 Email: ayamnash@gmail.com
- 🐛 Issues: [GitHub Issues](https://github.com/ayamnash/MCP_server_ms_access_control/issues)


---

Made with ❤️ for the MCP community










