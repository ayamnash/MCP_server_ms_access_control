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

@mcp.tool
def list_vba_modules(db_name: str) -> str:
    """List all VBA modules in the Access database"""
    try:
        path = get_db_path(db_name)
        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(path)
        
        # Access the VBA project
        project = access.VBE.VBProjects(1)
        
        modules = []
        for i in range(1, project.VBComponents.Count + 1):
            component = project.VBComponents(i)
            module_type = {
                1: "Standard Module",
                2: "Class Module", 
                3: "Form Module",
                100: "Document Module"
            }.get(component.Type, f"Type {component.Type}")
            
            modules.append(f"- {component.Name} ({module_type})")
        
        access.Quit()
        
        if modules:
            return "VBA Modules:\n" + "\n".join(modules)
        else:
            return "No VBA modules found"
            
    except Exception as e:
        return f"❌ Error listing VBA modules: {str(e)}"

@mcp.tool
def read_vba_module(db_name: str, module_name: str) -> str:
    """Read the code from a specific VBA module"""
    try:
        path = get_db_path(db_name)
        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(path)
        
        # Access the VBA project
        project = access.VBE.VBProjects(1)
        
        # Find the specific module
        module_found = False
        for i in range(1, project.VBComponents.Count + 1):
            component = project.VBComponents(i)
            if component.Name.lower() == module_name.lower():
                code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                module_found = True
                break
        
        access.Quit()
        
        if module_found:
            return f"VBA Code from module '{module_name}':\n\n{code}"
        else:
            return f"❌ Module '{module_name}' not found"
            
    except Exception as e:
        return f"❌ Error reading VBA module: {str(e)}"

@mcp.tool
def write_vba_module(db_name: str, module_name: str, code: str) -> str:
    """Create or replace a VBA module with the provided code"""
    try:
        path = get_db_path(db_name)
        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(path)
        
        # Access the VBA project
        project = access.VBE.VBProjects(1)
        
        # Check if module already exists
        module_exists = False
        for i in range(1, project.VBComponents.Count + 1):
            component = project.VBComponents(i)
            if component.Name.lower() == module_name.lower():
                # Clear existing code
                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                # Add new code
                component.CodeModule.AddFromString(code)
                module_exists = True
                break
        
        if not module_exists:
            # Create new standard module
            new_module = project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
            new_module.Name = module_name
            new_module.CodeModule.AddFromString(code)
        
        access.Quit()
        
        action = "updated" if module_exists else "created"
        return f"✅ VBA module '{module_name}' {action} successfully"
        
    except Exception as e:
        return f"❌ Error writing VBA module: {str(e)}"

@mcp.tool
def delete_vba_module(db_name: str, module_name: str) -> str:
    """Delete a VBA module from the Access database"""
    try:
        path = get_db_path(db_name)
        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(path)
        
        # Access the VBA project
        project = access.VBE.VBProjects(1)
        
        # Find and delete the module
        module_found = False
        for i in range(1, project.VBComponents.Count + 1):
            component = project.VBComponents(i)
            if component.Name.lower() == module_name.lower():
                project.VBComponents.Remove(component)
                module_found = True
                break
        
        access.Quit()
        
        if module_found:
            return f"✅ VBA module '{module_name}' deleted successfully"
        else:
            return f"❌ Module '{module_name}' not found"
            
    except Exception as e:
        return f"❌ Error deleting VBA module: {str(e)}"

@mcp.tool
def run_vba_function(db_name: str, function_name: str, args: str = "") -> str:
    """Execute a VBA function in the Access database and return the result. 
    Args should be comma-separated values like: 'arg1,arg2,arg3'"""
    try:
        path = get_db_path(db_name)
        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(path)
        
        # Parse arguments if provided
        if args.strip():
            arg_list = [arg.strip() for arg in args.split(',')]
            result = access.Run(function_name, *arg_list)
        else:
            result = access.Run(function_name)
        
        access.Quit()
        
        return f"✅ Function '{function_name}' executed successfully. Result: {result}"
        
    except Exception as e:
        return f"❌ Error running VBA function: {str(e)}"


if __name__ == "__main__":
    mcp.run()
 
