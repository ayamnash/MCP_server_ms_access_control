import os
import pyodbc
from fastmcp import FastMCP
from win32com.client import Dispatch
import win32com.client

mcp = FastMCP("Flexible Access DB MCP 🚀")

def get_db_path(db_name: str) -> str:
    if not db_name.lower().endswith(".accdb"):
        db_name += ".accdb"
    return os.path.join(os.path.expanduser("~"), db_name)

def get_driver() -> str:
    drivers = pyodbc.drivers()
    for d in [
        "Microsoft Access Driver (*.mdb, *.accdb)",
        "Microsoft Access Driver (*.accdb)",
        "Microsoft Access Driver (*.mdb)"
    ]:
        if d in drivers:
            return d
    raise Exception("Access ODBC driver not found")

@mcp.tool
def create_database(db_name: str) -> str:
    """Create an empty Access .accdb database"""
    

    path = get_db_path(db_name)
    if os.path.exists(path):
        os.remove(path)

    adox = Dispatch("ADOX.Catalog")
    conn_str = f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};"
    adox.Create(conn_str)
    return f"✅ Database created at: {path}"

@mcp.tool
def create_table(db_name: str, table_name: str, schema: str) -> str:
    """
    Create a table in the specified Access DB.
    schema example: "ID INT PRIMARY KEY, Name TEXT(100), Age INT"
    """
    path = get_db_path(db_name)
    driver = get_driver()
    conn_str = f"DRIVER={{{driver}}};DBQ={path};"
    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        sql = f"CREATE TABLE {table_name} ({schema})"
        cursor.execute(sql)
        return f"✅ Table '{table_name}' created in {db_name}"

@mcp.tool
def insert_data(db_name: str, table: str, rows: list[dict]) -> str:
    """Insert rows into a table. Example: [{'ID': 1, 'Name': 'Ali'}]"""
    path = get_db_path(db_name)
    driver = get_driver()
    conn_str = f"DRIVER={{{driver}}};DBQ={path};"
    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        for row in rows:
            columns = ', '.join(row.keys())
            placeholders = ', '.join('?' for _ in row)
            values = list(row.values())
            sql = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"
            cursor.execute(sql, values)
        conn.commit()
        return f"✅ Inserted {len(rows)} rows into '{table}'"

@mcp.tool
def run_query(db_name: str, sql: str) -> str:
    """Execute SQL queries (SELECT, UPDATE, DELETE, etc.)"""
    path = get_db_path(db_name)
    driver = get_driver()
    conn_str = f"DRIVER={{{driver}}};DBQ={path};"
    
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(sql)
            
            if sql.strip().lower().startswith("select"):
                columns = [col[0] for col in cursor.description]
                rows = cursor.fetchall()
                if rows:
                    result = f"Query Results ({len(rows)} rows):\n"
                    result += " | ".join(f"{col:<15}" for col in columns) + "\n"
                    result += "-" * (len(columns) * 17) + "\n"
                    for row in rows:
                        result += " | ".join(f"{str(val):<15}" for val in row) + "\n"
                    return result
                else:
                    return "No results found"
            else:
                conn.commit()
                return f"✅ Query executed successfully"
    except Exception as e:
        return f"❌ Error: {str(e)}"

@mcp.tool
def list_tables(db_name: str) -> str:
    """List all tables in the database"""
    path = get_db_path(db_name)
    driver = get_driver()
    conn_str = f"DRIVER={{{driver}}};DBQ={path};"
    
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            tables = cursor.tables(tableType='TABLE')
            table_names = [row.table_name for row in tables if not row.table_name.startswith('MSys')]
            
            if table_names:
                return "Tables:\n" + "\n".join(f"- {name}" for name in table_names)
            else:
                return "No tables found"
    except Exception as e:
        return f"❌ Error: {str(e)}"
@mcp.tool
def save_query(db_name: str, query_name: str, sql: str) -> str:
    """
    Save a named query inside the Access .accdb database.
    This acts like saving a query in the Access UI.
    """
    

    try:
        path = get_db_path(db_name)

        # Launch Access application
        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False  # run in background
        access.OpenCurrentDatabase(path)

        # Get the QueryDefs collection through DAO
        dao = access.CurrentDb()

        # If query already exists, delete it first
        try:
            dao.QueryDefs.Delete(query_name)
        except Exception:
            pass  # ignore if query doesn't exist

        # Create new query
        dao.CreateQueryDef(query_name, sql)

        # Optional: Close Access
        access.CloseCurrentDatabase()
        access.Quit()

        return f"✅ Saved query '{query_name}' in {db_name}"
    except Exception as e:
        return f"❌ Error saving query: {str(e)}"       
@mcp.tool
def list_saved_queries(db_name: str) -> str:
    """List all saved queries in the Access DB."""
    

    try:
        path = get_db_path(db_name)
        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(path)
        dao = access.CurrentDb()

        queries = [q.Name for q in dao.QueryDefs if not q.Name.startswith("~")]
        access.CloseCurrentDatabase()
        access.Quit()

        if queries:
            return "Saved Queries:\n" + "\n".join(f"- {q}" for q in queries)
        else:
            return "No saved queries found."
    except Exception as e:
        return f"❌ Error listing saved queries: {str(e)}"


if __name__ == "__main__":
    mcp.run()
 