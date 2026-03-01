
import os
import pyodbc
from fastmcp import FastMCP
from win32com.client import Dispatch
import win32com.client
import uuid
import random
import tempfile
import re
import gc
import pythoncom
import time
import logging
from typing import Callable, Tuple, Optional, List, Dict, Any

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

mcp = FastMCP("Flexible Access DB MCP")

# --- Configuration ---
class Config:
    """Configuration settings for the MCP server"""
    LOCK_TIMEOUT = 10  # seconds to wait for lock release
    CLEANUP_DELAY = 0.5  # seconds to wait after cleanup
    MAX_RETRIES = 3  # maximum retry attempts for transient errors
    RETRY_DELAY = 1.0  # seconds between retries
    POLL_INTERVAL = 0.5  # seconds between lock file checks

# --- State Tracking ---
_template_generated = False
_last_template_type = None
_batch_mode_db = None
_batch_mode_access = None

# --- Helper Functions ---

def _ensure_access_closed():
    """Force close all Access instances and clean up COM objects"""
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        try:
            access.Quit(1)  # acQuitSaveAll
            logger.debug("Successfully closed active Access instance")
        except win32com.client.pywintypes.com_error as e:
            logger.warning(f"COM error while closing Access: {e}")
        except Exception as e:
            logger.warning(f"Unexpected error while closing Access: {e}")
        finally:
            del access
    except win32com.client.pywintypes.com_error:
        # No active Access instance - this is expected
        logger.debug("No active Access instance to close")
    except Exception as e:
        logger.warning(f"Unexpected error in _ensure_access_closed: {e}")
    
    # Force COM cleanup
    try:
        pythoncom.CoUninitialize()
    except Exception as e:
        logger.debug(f"CoUninitialize error (may be expected): {e}")
    
    try:
        pythoncom.CoInitialize()
    except Exception as e:
        logger.debug(f"CoInitialize error (may be expected): {e}")
    
    gc.collect()
    time.sleep(Config.CLEANUP_DELAY)

def _with_access_database(db_name: str, operation_func: Callable) -> Any:
    """Context manager pattern for Access operations with automatic cleanup
    
    Args:
        db_name: Database name or path
        operation_func: Function that takes access object and returns result
        
    Returns:
        Result from operation_func
        
    Raises:
        Exception: If operation fails after retries
    """
    path = get_db_path(db_name)
    access = None
    
    try:
        # Check if in batch mode
        global _batch_mode_access, _batch_mode_db
        if _batch_mode_access and _batch_mode_db == db_name:
            logger.debug(f"Using existing batch connection for {db_name}")
            return operation_func(_batch_mode_access)
        
        # Normal mode - open, execute, close
        logger.info(f"Opening database: {path}")
        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(path)
        
        result = operation_func(access)
        
        # Save and close
        try:
            access.DoCmd.Save()
            logger.debug("Database saved successfully")
        except Exception as e:
            logger.debug(f"Save not needed or failed (may be expected): {e}")
        
        access.CloseCurrentDatabase()
        access.Quit(1)
        logger.info(f"Database closed successfully: {path}")
        
        return result
        
    except win32com.client.pywintypes.com_error as e:
        logger.error(f"COM error in database operation: {e}")
        raise Exception(f"COM error: {str(e)}")
    except Exception as e:
        logger.error(f"Error in database operation: {e}")
        raise
    finally:
        if access and not _batch_mode_access:
            try:
                access.Quit(1)
            except Exception as e:
                logger.debug(f"Error during final quit (may be expected): {e}")
            del access
            _ensure_access_closed()

def is_database_locked(db_path: str) -> bool:
    """Check if database has an active lock file
    
    Args:
        db_path: Full path to database file
        
    Returns:
        True if lock file exists, False otherwise
    """
    lock_file = db_path.replace('.accdb', '.laccdb')
    locked = os.path.exists(lock_file)
    if locked:
        logger.warning(f"Database is locked: {lock_file}")
    return locked

def wait_for_lock_release(db_path: str, timeout: Optional[int] = None) -> Tuple[bool, str]:
    """Wait for lock file to be released
    
    Args:
        db_path: Full path to database file
        timeout: Maximum seconds to wait (default: Config.LOCK_TIMEOUT)
        
    Returns:
        Tuple of (success: bool, message: str)
    """
    if timeout is None:
        timeout = Config.LOCK_TIMEOUT
        
    lock_file = db_path.replace('.accdb', '.laccdb')
    
    if not os.path.exists(lock_file):
        return True, "Database is not locked"
    
    logger.info(f"Waiting for lock release: {lock_file} (timeout: {timeout}s)")
    start_time = time.time()
    
    while os.path.exists(lock_file):
        elapsed = time.time() - start_time
        if elapsed > timeout:
            msg = f"Timeout: Database still locked after {timeout} seconds. Please close MS Access manually."
            logger.error(msg)
            return False, msg
        time.sleep(Config.POLL_INTERVAL)
    
    logger.info(f"Lock released after {time.time() - start_time:.1f} seconds")
    return True, "Lock released"

def _validate_module_name(module_name: str) -> Tuple[bool, str]:
    """Validate VBA module name
    
    Args:
        module_name: Name to validate
        
    Returns:
        Tuple of (is_valid: bool, error_message: str)
    """
    if not module_name or not module_name.strip():
        return False, "Module name cannot be empty"
    
    if not re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*$', module_name):
        return False, "Module name must be a valid VBA identifier (letters, numbers, underscore; cannot start with number)"
    
    if len(module_name) > 64:
        return False, "Module name too long (max 64 characters)"
    
    # VBA reserved words
    reserved = ['Sub', 'Function', 'End', 'If', 'Then', 'Else', 'For', 'Next', 'Do', 'Loop', 
                'While', 'Select', 'Case', 'Dim', 'As', 'Integer', 'String', 'Boolean']
    if module_name in reserved:
        return False, f"Module name '{module_name}' is a VBA reserved word"
    
    return True, ""

def _validate_database_name(db_name: str) -> Tuple[bool, str]:
    """Validate database name/path
    
    Args:
        db_name: Database name or path to validate
        
    Returns:
        Tuple of (is_valid: bool, error_message: str)
    """
    if not db_name or not db_name.strip():
        return False, "Database name cannot be empty"
    
    # Check for path traversal attempts
    if '..' in db_name:
        return False, "Database name cannot contain '..' (path traversal)"
    
    return True, ""

# IMPROVED get_db_path function with better path detection
def get_db_path(db_name: str) -> str:
    """Gets the full path for the database. Handles both absolute and relative paths.
    Now includes better path detection and validation."""
    
    # If the path is already absolute (e.g., "F:\...") use it directly.
    if os.path.isabs(db_name):
        if not db_name.lower().endswith(".accdb"):
            db_name += ".accdb"
        return db_name
    
    # For relative paths, try multiple locations in order of preference:
    if not db_name.lower().endswith(".accdb"):
        db_name += ".accdb"
    
    # 1. Current working directory (most common for development)
    current_dir_path = os.path.join(os.getcwd(), db_name)
    if os.path.exists(current_dir_path):
        return current_dir_path
    
    # 2. User's home directory (original behavior)
    home_dir_path = os.path.join(os.path.expanduser("~"), db_name)
    if os.path.exists(home_dir_path):
        return home_dir_path
    
    # 3. If neither exists, default to current directory (for new database creation)
    return current_dir_path

def get_driver() -> str:
    """Finds a suitable Microsoft Access ODBC driver."""
    drivers = pyodbc.drivers()
    for d in [
        "Microsoft Access Driver (*.mdb, *.accdb)",
        "Microsoft Access Driver (*.accdb)",
        "Microsoft Access Driver (*.mdb)"
    ]:
        if d in drivers:
            return d
    raise Exception("Access ODBC driver not found")



def _run_query_internal(db_name: str, sql: str) -> str:
    """Internal helper to run any SQL query."""
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
                return "Query executed successfully"
    except Exception as e:
        return f"Error: {str(e)}"

def _get_table_schema(db_name: str, table_name: str) -> list[str]:
    """Internal helper to get column names for a table or query."""
    path = get_db_path(db_name)
    driver = get_driver()
    conn_str = f"DRIVER={{{driver}}};DBQ={path};"
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            # Try to get schema by running a SELECT query, which works for both tables and queries
            cursor.execute(f"SELECT * FROM [{table_name}] WHERE 1=0")
            columns = [col[0] for col in cursor.description]
            if not columns:
                raise ValueError(f"Table or query '{table_name}' not found or has no columns.")
            return columns
    except Exception as e:
        raise ValueError(f"Could not retrieve schema for table or query '{table_name}'. Error: {e}")
def sanitize_access_schema(schema: str) -> str:
    replacements = {
        r"\bAUTOINCREMENT\b": "COUNTER",
        r"\bINTEGER\b": "LONG",
        r"\bINT\b": "LONG",
        r"\bBIGINT\b": "LONG",
        r"\bBOOLEAN\b": "YESNO",
        r"\bBIT\b": "YESNO",
        r"\bLONGTEXT\b": "MEMO",
        r"\bTEXT\(MAX\)": "MEMO",
        r"\bDECIMAL\([^)]+\)": "CURRENCY",
        r"\bNUMERIC\([^)]+\)": "CURRENCY",
    }
    for pattern, repl in replacements.items():
        schema = re.sub(pattern, repl, schema, flags=re.IGNORECASE)
    
    # Remove DEFAULT clauses that Access doesn't handle well in CREATE TABLE
    schema = re.sub(r"DEFAULT\s+NOW\(\)", "", schema, flags=re.IGNORECASE)
    schema = re.sub(r"DEFAULT\s+CURRENT_TIMESTAMP", "", schema, flags=re.IGNORECASE)
    schema = re.sub(r"DEFAULT\s+TRUE", "", schema, flags=re.IGNORECASE)
    schema = re.sub(r"DEFAULT\s+-1", "", schema, flags=re.IGNORECASE)
    schema = re.sub(r"DEFAULT\s+0", "", schema, flags=re.IGNORECASE)
    schema = re.sub(r"DEFAULT\s+'[^']*'", "", schema, flags=re.IGNORECASE)
    
    # Wrap reserved words in brackets
    reserved_words = ["Status", "Notes", "Description", "Name", "Date", "User"]
    for word in reserved_words:
        schema = re.sub(rf"\b{word}\b(?!\])", f"[{word}]", schema, flags=re.IGNORECASE)
    
    # Clean up extra spaces and fix malformed parentheses
    schema = re.sub(r"\s{2,}", " ", schema)
    schema = re.sub(r",\s*\)", ")", schema)
    schema = re.sub(r"\(\s*,", "(", schema)
    
    return schema.strip()
@mcp.tool()
def save_and_close_access_database(db_name: str) -> dict:
    """
    Save all changes and close the MS Access database.
    If Access is not running, returns a safe success message.
    """
    try:
        access_app = win32com.client.GetActiveObject("Access.Application")
        current_db = access_app.CurrentDb()

        if current_db is None:
            return {"success": False, "message": "No database is currently open in Access."}

        current_path = current_db.Name

        if db_name.lower() not in current_path.lower():
            return {
                "success": False,
                "message": f"The open database '{current_path}' does not match '{db_name}'."
            }

        access_app.DoCmd.Save()
        access_app.Quit(1)  # acQuitSaveAll = 1

        lock_file = current_path.replace('.accdb', '.laccdb')

        return {
            "success": True,
            "message": f"'{current_path}' saved and closed successfully.",
            "lock_file_released": not os.path.exists(lock_file)
        }

    except win32com.client.pywintypes.com_error:
        return {"success": True, "message": "MS Access was not running. Nothing to close."}
    except Exception as e:
        return {"success": False, "message": f"Unexpected error: {str(e)}"}

@mcp.tool
def create_database(db_name: str) -> str:
    """Create an empty Access .accdb database"""
    path = get_db_path(db_name)
    if os.path.exists(path):
        os.remove(path)
    adox = Dispatch("ADOX.Catalog")
    conn_str = f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};"
    adox.Create(conn_str)
    return f"Database created at: {path}"

@mcp.tool()
def create_table(db_name: str, table_name: str, schema: str) -> str:
    """Creates a table in the Access database."""
    db_path = get_db_path(db_name)
    sanitized_schema = sanitize_access_schema(schema)
    sql = f"CREATE TABLE [{table_name}] ({sanitized_schema})"
    
    # Debug output to see what's happening
    print(f"Original schema: {schema}")
    print(f"Sanitized schema: {sanitized_schema}")
    print(f"Final SQL: {sql}")
    
    try:
        driver = get_driver()
        conn_str = f"DRIVER={{{driver}}};DBQ={db_path};"
        conn = pyodbc.connect(conn_str)
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        cur.close()
        conn.close()
        return f"Table '{table_name}' created successfully."
    except Exception as e:
        return f"Error creating table '{table_name}': {str(e)}"
    

@mcp.tool
def insert_data(db_name: str, table: str, rows: list[dict]) -> str:
    """Insert rows into a table. Example: [{'ID': 1, 'Name': 'Ali'}]"""
    path = get_db_path(db_name)
    driver = get_driver()
    conn_str = f"DRIVER={{{driver}}};DBQ={path};"
    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        for row in rows:
            columns = ', '.join(f"[{c}]" for c in row.keys())
            placeholders = ', '.join('?' for _ in row)
            values = list(row.values())
            sql = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"
            cursor.execute(sql, values)
        conn.commit()
        return f"Inserted {len(rows)} rows into '{table}'"

@mcp.tool
def run_query(db_name: str, sql: str) -> str:
    """Run a SELECT or action query (INSERT, UPDATE, DELETE)."""
    return _run_query_internal(db_name, sql)

@mcp.tool
def find_database(db_name: str) -> str:
    """Debug tool to find where a database file actually exists"""
    possible_paths = []
    
    # Add the resolved path from get_db_path
    resolved_path = get_db_path(db_name)
    possible_paths.append(("get_db_path() result", resolved_path, os.path.exists(resolved_path)))
    
    # Add current directory
    if not db_name.lower().endswith('.accdb'):
        db_name_with_ext = db_name + '.accdb'
    else:
        db_name_with_ext = db_name
    
    current_dir = os.path.join(os.getcwd(), db_name_with_ext)
    possible_paths.append(("Current directory", current_dir, os.path.exists(current_dir)))
    
    # Add home directory
    home_dir = os.path.join(os.path.expanduser("~"), db_name_with_ext)
    possible_paths.append(("Home directory", home_dir, os.path.exists(home_dir)))
    
    # If db_name looks like an absolute path, check it
    if os.path.isabs(db_name):
        possible_paths.append(("Absolute path (as-is)", db_name, os.path.exists(db_name)))
        if not db_name.lower().endswith('.accdb'):
            abs_with_ext = db_name + '.accdb'
            possible_paths.append(("Absolute path + .accdb", abs_with_ext, os.path.exists(abs_with_ext)))
    
    result = f"Database search results for '{db_name}':\n"
    result += f"Current working directory: {os.getcwd()}\n\n"
    
    found_any = False
    for description, path, exists in possible_paths:
        status = "✓ EXISTS" if exists else "✗ Not found"
        result += f"{description}: {status}\n  {path}\n\n"
        if exists:
            found_any = True
    
    if found_any:
        result += "✓ Database found in at least one location."
    else:
        result += "✗ Database not found in any checked location."
    
    return result

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
        return f"Error: {str(e)}"
def fix_access_sql_syntax(sql: str) -> str:
    """
    Automatically fix common Access SQL syntax issues:
    1. Convert double quotes to single quotes for string literals
    2. Keep double quotes only for special cases like Format functions
    3. Fix multiple JOIN syntax by adding proper parentheses
    """
    # Pattern to match string literals that should use single quotes
    # This matches double quotes that are NOT part of function calls like Format("yyyy-mm-dd")
    
    # First, protect Format function quotes and similar cases
    protected_patterns = []
    
    # Find and temporarily replace Format function quotes
    format_pattern = r'(Format\s*\([^,]+,\s*)"([^"]+)"'
    def protect_format(match):
        placeholder = f"__PROTECTED_QUOTE_{len(protected_patterns)}__"
        protected_patterns.append(f'"{match.group(2)}"')
        return f'{match.group(1)}{placeholder}'
    
    sql = re.sub(format_pattern, protect_format, sql, flags=re.IGNORECASE)
    
    # Now convert remaining double quotes to single quotes for string literals
    # This pattern matches double quotes around values (not in function contexts)
    sql = re.sub(r'=\s*"([^"]*)"', r"= '\1'", sql)  # = "value" -> = 'value'
    sql = re.sub(r'<>\s*"([^"]*)"', r"<> '\1'", sql)  # <> "value" -> <> 'value'
    sql = re.sub(r'IN\s*\(\s*"([^"]*)"', r"IN ('\1'", sql, flags=re.IGNORECASE)  # IN ("value" -> IN ('value'
    sql = re.sub(r'LIKE\s*"([^"]*)"', r"LIKE '\1'", sql, flags=re.IGNORECASE)  # LIKE "value" -> LIKE 'value'
    
    # Fix multiple JOIN syntax for Access
    # Access requires parentheses around multiple JOINs
    # Pattern: FROM table1 INNER JOIN table2 ON ... INNER JOIN table3 ON ...
    # Should become: FROM (table1 INNER JOIN table2 ON ...) INNER JOIN table3 ON ...
    
    # Find FROM clause with multiple INNER JOINs
    from_pattern = r'FROM\s+([^()]+?)\s+INNER\s+JOIN\s+([^()]+?)\s+ON\s+([^()]+?)\s+INNER\s+JOIN'
    if re.search(from_pattern, sql, re.IGNORECASE):
        # Replace the pattern to add parentheses around the first JOIN
        sql = re.sub(
            from_pattern,
            r'FROM (\1 INNER JOIN \2 ON \3) INNER JOIN',
            sql,
            flags=re.IGNORECASE
        )
    
    # Handle LEFT JOIN cases too
    from_pattern_left = r'FROM\s+([^()]+?)\s+LEFT\s+JOIN\s+([^()]+?)\s+ON\s+([^()]+?)\s+(?:INNER|LEFT)\s+JOIN'
    if re.search(from_pattern_left, sql, re.IGNORECASE):
        sql = re.sub(
            from_pattern_left,
            r'FROM (\1 LEFT JOIN \2 ON \3) INNER JOIN' if 'INNER JOIN' in sql.upper() else r'FROM (\1 LEFT JOIN \2 ON \3) LEFT JOIN',
            sql,
            flags=re.IGNORECASE
        )
    
    # Restore protected quotes
    for i, protected in enumerate(protected_patterns):
        sql = sql.replace(f"__PROTECTED_QUOTE_{i}__", protected)
    
    return sql

@mcp.tool
def save_query(db_name: str, query_name: str, sql: str) -> str:
    """Save or overwrite a named query in an Access database.
    Automatically fixes common Access SQL syntax issues like double quotes.
    
    Args:
        db_name: Database name or path
        query_name: Name for the saved query
        sql: SQL query text
        
    Returns:
        Success or error message
    """
    # Validate inputs
    is_valid, error_msg = _validate_database_name(db_name)
    if not is_valid:
        logger.error(f"Invalid database name: {error_msg}")
        return f"Error: {error_msg}"
    
    if not query_name or not query_name.strip():
        logger.error("Query name cannot be empty")
        return "Error: Query name cannot be empty"
    
    if not sql or not sql.strip():
        logger.error("SQL cannot be empty")
        return "Error: SQL cannot be empty"
    
    def operation(access):
        """Inner function to save the query"""
        logger.info(f"Saving query: {query_name}")
        
        # Fix Access SQL syntax issues
        sql_fixed = fix_access_sql_syntax(sql)
        
        # For COM interface, escape double quotes
        sql_escaped = sql_fixed.replace('"', '""')
        
        dao = access.CurrentDb()
        
        # Delete existing query if it exists
        try:
            dao.QueryDefs.Delete(query_name)
            logger.debug(f"Deleted existing query: {query_name}")
        except Exception:
            logger.debug(f"Query {query_name} doesn't exist (creating new)")
        
        # Create new query with escaped SQL
        dao.CreateQueryDef(query_name, sql_escaped)
        logger.info(f"Query '{query_name}' saved successfully")
        
        return f"Query '{query_name}' saved successfully"
    
    try:
        path = get_db_path(db_name)
        
        # Check if database exists
        if not os.path.exists(path):
            logger.error(f"Database not found: {path}")
            return f"Error: Database not found at {path}"
        
        # Check for lock
        if is_database_locked(path):
            success, message = wait_for_lock_release(path)
            if not success:
                return f"Error: {message}"
        
        result = _with_access_database(db_name, operation)
        return result
        
    except Exception as e:
        logger.error(f"Error saving query '{query_name}': {e}")
        return f"Error saving query '{query_name}': {str(e)}"






@mcp.tool
def generate_form_template(
    db_name: str, 
    record_source: str, 
    form_type: str, 
    subform_object_name: str = None, 
    link_master_field: str = None, 
    link_child_field: str = None
) -> str:
    """
    STEP 1/2 for creating a form. Generates a text template for an Access form.
    The LLM must complete this template and pass it to 'create_form_from_llm_text'.
    
    Workflow for a single form:
    1. Call this tool with form_type='single' or 'subform'.
    
    Workflow for a form with a subform:
    1. First, create the subform object (e.g., 'movements_subform') using the full two-step process.
    2. Then, call this tool with form_type='main', providing the main form's record_source, the subform_object_name, and the linking fields.

    Args:
        db_name: The name of the database file (e.g., 'inventory.accdb'). Can be an absolute path.
        record_source: The name of the table or saved query the form is based on.
        form_type: The type of form. Must be one of: 'single', 'subform', 'main'.
                   - 'single': A standard, standalone form.
                   - 'subform': A form intended to be embedded, usually in Datasheet view.
                   - 'main': A form that will contain a subform.
        subform_object_name: (Required for 'main' type) The name of the already-created form object to use as the subform. e.g. 'Form.movements_subform'
        link_master_field: (Required for 'main' type) The linking field from the main form's record source. e.g. 'ProductID'
        link_child_field: (Required for 'main' type) The linking field from the subform's record source. e.g. 'ProductID'
    """
    global _template_generated, _last_template_type
    
    if form_type not in ['single', 'subform', 'main']:
        return "Error: form_type must be 'single', 'subform', or 'main'."
    if form_type == 'main' and not (subform_object_name and link_master_field and link_child_field):
        return "Error: For 'main' form_type, you must provide subform_object_name, link_master_field, and link_child_field."

    try:
        # This check also validates that the record_source exists.
        fields = _get_table_schema(db_name, record_source)
    except Exception as e:
        return f"Error getting schema for record source '{record_source}': {e}"

    form_guid = str(uuid.uuid4()).replace('-', '')
    
    # --- Generate Controls and NameMap ---
    controls_text = ""
    namemap_entries = []
    y_pos = 200 # Starting Y position for controls
    
    # For a main form, we only want specific fields as per the user request.
    # This logic can be enhanced, but for this specific request, we'll customize it.
    # A more advanced version might take a list of fields as an argument.
    fields_to_show = fields
    if form_type == 'main' and record_source == 'movements':
        fields_to_show = ['ProductID', 'ProductName']


    for i, field in enumerate(fields_to_show):
        controls_text += f"""
                Begin TextBox
                    OverlapFlags =85
                    Left =2500
                    Top ={y_pos}
                    Height =315
                    Width = 3000
                    TabIndex ={i}
                    Name ="{field}"
                    ControlSource ="{field}"
                    GUID = Begin
                        0x{uuid.uuid4().hex}
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =500
                            Top ={y_pos}
                            Width =1800
                            Height =315
                            Name ="{field}_Label"
                            Caption ="{field}"
                            GUID = Begin
                                0x{uuid.uuid4().hex}
                            End
                        End
                    End
                End"""
        rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
        field_hex = field.encode('utf-16le').hex()
        namemap_entries.append(f"0x{rand_hex}{len(field):02x}000000{field_hex}")
        y_pos += 400

    namemap_text = ",\n        ".join(namemap_entries) + ",\n        0x000000000000000000000000000000000c000000050000000000000000000000000000000000"

    if form_type == 'main':
        clean_subform_name = re.sub(r'^Form\.', '', subform_object_name)
        subform_guid = uuid.uuid4().hex

        controls_text += f"""
                Begin Subform
                    OverlapFlags =85
                    Left =500
                    Top ={y_pos + 200}
                    Width =10000
                    Height =4000
                    TabIndex ={len(fields_to_show)}
                    Name ="{clean_subform_name}"
                    SourceObject ="{subform_object_name}"
                    LinkChildFields ="{link_child_field}"
                    LinkMasterFields ="{link_master_field}"
                    GUID = Begin
                        0x{subform_guid}
                    End
                End"""

    view_type = "2" if form_type == 'subform' else "0"
    
    template = f"""Version =21
VersionRequired =20
PublishOption =1
Checksum ={random.randint(-2000000000, 2000000000)}
Begin Form
    DefaultView ={view_type}
    Width =11500
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    GUID = Begin
        0x{form_guid}
    End
    NameMap = Begin
        {namemap_text}
    End
    RecordSource ="{record_source}"
    Caption ="__FORM_NAME_PLACEHOLDER__"
    Begin
        Begin Section
            Height ={y_pos + (4500 if form_type == 'main' else 500)}
            Name ="Detail"
            AutoHeight= -1
            Begin
                {controls_text}
            End
        End
    End
End
"""
    _template_generated = True
    _last_template_type = form_type
    
    return f"""Template generated successfully.
IMPORTANT: 
1. Replace '__FORM_NAME_PLACEHOLDER__' with the desired form name.
2. Review the template below. You can adjust properties like layout (Left, Top, Width, Height) if needed.
3. Pass the **entire, final text content** to the 'create_form_from_llm_text' tool.

--- TEMPLATE BEGIN ---
{template}
--- TEMPLATE END ---
"""



@mcp.tool
def create_form_from_llm_text(db_name: str, form_name: str, form_text: str) -> str:
    """STEP 2/2 for creating a form. Creates an Access form from its text definition.
    
    This tool will automatically correct/generate the NameMap and GUIDs based on the
    controls found in the form_text, making it robust against LLM-generated errors.
    
    Args:
        db_name: The name of the database file (e.g., 'inventory.accdb'). Can be an absolute path.
        form_name: The name to save the form as (e.g., 'ProductsForm').
        form_text: The complete text definition of the form.
        
    Returns:
        Success or error message
    """
    # Validate inputs
    is_valid, error_msg = _validate_database_name(db_name)
    if not is_valid:
        logger.error(f"Invalid database name: {error_msg}")
        return f"Error: {error_msg}"
    
    if not form_name or not form_name.strip():
        logger.error("Form name cannot be empty")
        return "Error: Form name cannot be empty"
    
    if not form_text or not form_text.strip():
        logger.error("Form text cannot be empty")
        return "Error: Form text cannot be empty"
    
    # --- PRE-PROCESSING AND VALIDATION ---
    try:
        logger.info(f"Pre-processing form: {form_name}")
        
        # 1. Replace placeholder if it exists
        if "__FORM_NAME_PLACEHOLDER__" in form_text:
             form_text = form_text.replace("__FORM_NAME_PLACEHOLDER__", form_name)

        # 2. Find all controls with a 'Name' property
        control_names = re.findall(r'^\s*Name\s*=\s*"([^"]+)"', form_text, re.MULTILINE)
        if not control_names:
            logger.error("No named controls found in form text")
            return "Error: Could not find any named controls in the form text to build a NameMap."

        # 3. Generate a fresh, correct NameMap
        namemap_entries = []
        for name in control_names:
            rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
            field_hex = name.encode('utf-16le').hex()
            namemap_entries.append(f"0x{rand_hex}{len(name):02x}000000{field_hex}")
        
        # Add the required terminator for the NameMap
        namemap_terminator = "0x000000000000000000000000000000000c000000050000000000000000000000000000000000"
        namemap_entries.append(namemap_terminator)
        
        new_namemap_text = "NameMap = Begin\n        " + ",\n        ".join(namemap_entries) + "\n    End"

        # 4. Replace the old NameMap in the text with our new, correct one
        form_text = re.sub(r'NameMap\s*=\s*Begin.*?End', new_namemap_text, form_text, flags=re.DOTALL)

        # 5. Find and fix all GUIDs
        def replace_guid(match):
            guid_content = match.group(1).strip().replace('0x', '')
            if len(guid_content) == 32 and all(c in '0123456789abcdefABCDEF' for c in guid_content):
                return match.group(0)
            else:
                return f"GUID = Begin\n            0x{uuid.uuid4().hex}\n        End"

        form_text = re.sub(r'GUID\s*=\s*Begin(.*?)End', replace_guid, form_text, flags=re.DOTALL)
        
        logger.debug("Form text pre-processing completed")

    except Exception as e:
        logger.error(f"Error during form pre-processing: {e}")
        return f"An unexpected error occurred during pre-processing: {e}"

    def operation(access):
        """Inner function to create the form"""
        logger.info(f"Creating form: {form_name}")
        
        # Write to temp file
        temp_file_path = None
        try:
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix=".txt", encoding='utf-8') as tf:
                tf.write(form_text)
                temp_file_path = tf.name
            
            AC_FORM = 2
            
            # Delete existing form if it exists
            try:
                access.DoCmd.DeleteObject(AC_FORM, form_name)
                logger.debug(f"Deleted existing form: {form_name}")
            except Exception:
                logger.debug(f"Form {form_name} doesn't exist (creating new)")

            # Load form from text file
            access.LoadFromText(AC_FORM, form_name, temp_file_path)
            logger.info(f"Form '{form_name}' created successfully")
            
            global _template_generated, _last_template_type
            _template_generated = False
            _last_template_type = None

            return f"Form '{form_name}' created successfully in database '{db_name}'."
            
        finally:
            if temp_file_path and os.path.exists(temp_file_path):
                os.remove(temp_file_path)
                logger.debug("Temp file cleaned up")

    try:
        path = get_db_path(db_name)
        
        # Check for lock
        if is_database_locked(path):
            success, message = wait_for_lock_release(path)
            if not success:
                return f"Error: {message}"
        
        result = _with_access_database(db_name, operation)
        return result
        
    except Exception as e:
        logger.error(f"Error creating form '{form_name}': {e}")
        return f"Error creating form from text: {str(e)}"

@mcp.tool
def list_vba_modules(db_name: str) -> str:
    """List all VBA modules in the Access database"""
    
    def operation(access):
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
        
        if modules:
            return "VBA Modules:\n" + "\n".join(modules)
        else:
            return "No VBA modules found"
    
    try:
        path = get_db_path(db_name)
        if is_database_locked(path):
            success, message = wait_for_lock_release(path, timeout=10)
            if not success:
                return f"Error: {message}"
        
        result = _with_access_database(db_name, operation)
        return result
        
    except Exception as e:
        return f"Error listing VBA modules: {str(e)}"

@mcp.tool
def read_vba_module(db_name: str, module_name: str) -> str:
    """Read the code from a specific VBA module"""
    
    def operation(access):
        project = access.VBE.VBProjects(1)
        
        # Find the specific module
        for i in range(1, project.VBComponents.Count + 1):
            component = project.VBComponents(i)
            if component.Name.lower() == module_name.lower():
                if component.CodeModule.CountOfLines > 0:
                    code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                    return f"VBA Code from module '{module_name}':\n\n{code}"
                else:
                    return f"Module '{module_name}' exists but is empty"
        
        return f"Module '{module_name}' not found"
    
    try:
        path = get_db_path(db_name)
        if is_database_locked(path):
            success, message = wait_for_lock_release(path, timeout=10)
            if not success:
                return f"Error: {message}"
        
        result = _with_access_database(db_name, operation)
        return result
        
    except Exception as e:
        return f"Error reading VBA module '{module_name}': {str(e)}"

@mcp.tool
def write_vba_module(db_name: str, module_name: str, code: str) -> str:
    """Create or replace a VBA module with the provided code.
    
    Automatically saves and closes the database to prevent lock issues.
    
    Args:
        db_name: Database name or path
        module_name: Name for the VBA module (must be valid VBA identifier)
        code: VBA code to write
        
    Returns:
        Success or error message
    """
    # Validate inputs
    is_valid, error_msg = _validate_database_name(db_name)
    if not is_valid:
        logger.error(f"Invalid database name: {error_msg}")
        return f"Error: {error_msg}"
    
    is_valid, error_msg = _validate_module_name(module_name)
    if not is_valid:
        logger.error(f"Invalid module name: {error_msg}")
        return f"Error: {error_msg}"
    
    if not code or not code.strip():
        logger.error("VBA code cannot be empty")
        return "Error: VBA code cannot be empty"
    
    def operation(access):
        """Inner function that does the actual work"""
        logger.info(f"Writing VBA module: {module_name}")
        
        # Access the VBA project
        project = access.VBE.VBProjects(1)
        
        # Check if module already exists
        module_exists = False
        for i in range(1, project.VBComponents.Count + 1):
            component = project.VBComponents(i)
            if component.Name.lower() == module_name.lower():
                # Clear existing code
                if component.CodeModule.CountOfLines > 0:
                    component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                # Add new code
                component.CodeModule.AddFromString(code)
                module_exists = True
                logger.info(f"Updated existing module: {module_name}")
                break
        
        if not module_exists:
            # Create new standard module
            new_module = project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
            new_module.Name = module_name
            new_module.CodeModule.AddFromString(code)
            logger.info(f"Created new module: {module_name}")
        
        action = "updated" if module_exists else "created"
        return f"VBA module '{module_name}' {action} successfully"
    
    try:
        # Check for lock before starting
        path = get_db_path(db_name)
        if is_database_locked(path):
            success, message = wait_for_lock_release(path)
            if not success:
                return f"Error: {message}"
        
        # Execute operation with automatic cleanup
        result = _with_access_database(db_name, operation)
        logger.info(f"Successfully wrote VBA module: {module_name}")
        return result
        
    except Exception as e:
        logger.error(f"Error writing VBA module '{module_name}': {e}")
        return f"Error writing VBA module '{module_name}': {str(e)}"

@mcp.tool
def delete_vba_module(db_name: str, module_name: str) -> str:
    """Delete a VBA module from the Access database"""
    
    def operation(access):
        project = access.VBE.VBProjects(1)
        
        # Find and delete the module
        for i in range(1, project.VBComponents.Count + 1):
            component = project.VBComponents(i)
            if component.Name.lower() == module_name.lower():
                project.VBComponents.Remove(component)
                return f"VBA module '{module_name}' deleted successfully"
        
        return f"Module '{module_name}' not found"
    
    try:
        path = get_db_path(db_name)
        if is_database_locked(path):
            success, message = wait_for_lock_release(path, timeout=10)
            if not success:
                return f"Error: {message}"
        
        result = _with_access_database(db_name, operation)
        return result
        
    except Exception as e:
        return f"Error deleting VBA module '{module_name}': {str(e)}"

@mcp.tool
def run_vba_function(db_name: str, function_name: str, args: str = "") -> str:
    """Execute a VBA function in the Access database and return the result. 
    Args should be comma-separated values like: 'arg1,arg2,arg3'"""
    
    def operation(access):
        # Parse arguments if provided
        if args.strip():
            arg_list = [arg.strip() for arg in args.split(',')]
            result = access.Run(function_name, *arg_list)
        else:
            result = access.Run(function_name)
        
        return f"Function '{function_name}' executed successfully. Result: {result}"
    
    try:
        path = get_db_path(db_name)
        if is_database_locked(path):
            success, message = wait_for_lock_release(path, timeout=10)
            if not success:
                return f"Error: {message}"
        
        result = _with_access_database(db_name, operation)
        return result
        
    except Exception as e:
        return f"Error running VBA function '{function_name}': {str(e)}"

@mcp.tool
def begin_batch_operation(db_name: str) -> str:
    """Start a batch operation - keeps database open for multiple commands.
    
    Use this when you need to perform multiple operations (create tables, forms, VBA modules)
    in sequence. This is much faster than individual operations.
    
    IMPORTANT: You MUST call commit_batch_operation() when done!
    """
    global _batch_mode_db, _batch_mode_access
    
    if _batch_mode_access:
        return f"Error: Batch operation already in progress for '{_batch_mode_db}'"
    
    try:
        path = get_db_path(db_name)
        
        # Check for lock
        if is_database_locked(path):
            success, message = wait_for_lock_release(path, timeout=10)
            if not success:
                return f"Error: {message}"
        
        _batch_mode_access = win32com.client.Dispatch("Access.Application")
        _batch_mode_access.Visible = False
        _batch_mode_access.OpenCurrentDatabase(path)
        _batch_mode_db = db_name
        
        return f"✓ Batch operation started for '{db_name}'. Database will stay open until you call commit_batch_operation()."
    
    except Exception as e:
        _batch_mode_access = None
        _batch_mode_db = None
        return f"Error starting batch operation: {str(e)}"

@mcp.tool
def commit_batch_operation() -> str:
    """End batch operation, save all changes, and close database.
    
    Call this after you've completed all operations in a batch.
    """
    global _batch_mode_db, _batch_mode_access
    
    if not _batch_mode_access:
        return "Error: No batch operation in progress"
    
    db_name = _batch_mode_db
    
    try:
        # Save all changes
        _batch_mode_access.DoCmd.Save()
        
        # Close database
        _batch_mode_access.CloseCurrentDatabase()
        _batch_mode_access.Quit(1)
        
        # Clear state
        _batch_mode_db = None
        _batch_mode_access = None
        
        # Force cleanup
        _ensure_access_closed()
        
        return f"✓ Batch operation committed successfully for '{db_name}'. Database closed and saved."
    
    except Exception as e:
        # Try to cleanup even on error
        try:
            if _batch_mode_access:
                _batch_mode_access.Quit(1)
        except:
            pass
        
        _batch_mode_db = None
        _batch_mode_access = None
        _ensure_access_closed()
        
        return f"Error committing batch operation: {str(e)}"

@mcp.tool
def rollback_batch_operation() -> str:
    """Cancel batch operation without saving changes and close database.
    
    Use this if something went wrong and you want to discard all changes.
    """
    global _batch_mode_db, _batch_mode_access
    
    if not _batch_mode_access:
        return "Error: No batch operation in progress"
    
    db_name = _batch_mode_db
    
    try:
        # Close without saving
        _batch_mode_access.CloseCurrentDatabase()
        _batch_mode_access.Quit(0)  # acQuitSaveNone
        
        _batch_mode_db = None
        _batch_mode_access = None
        
        _ensure_access_closed()
        
        return f"✓ Batch operation rolled back for '{db_name}'. Changes discarded."
    
    except Exception as e:
        _batch_mode_db = None
        _batch_mode_access = None
        _ensure_access_closed()
        
        return f"Error rolling back batch operation: {str(e)}"

def _generate_report_template_internal(db_name: str, record_source: str, report_type: str = "tabular") -> str:
    """Internal helper function to generate report template without MCP tool wrapper."""
    try:
        # Validate record source and get fields
        fields = _get_table_schema(db_name, record_source)
        
        report_guid = str(uuid.uuid4()).replace('-', '')
        
        # Generate controls based on report type
        if report_type.lower() == "columnar":
            # Columnar layout - fields stacked vertically
            controls_text = ""
            namemap_entries = []
            y_pos = 500
            
            for i, field in enumerate(fields):
                controls_text += f"""
                Begin Label
                    OverlapFlags =85
                    Left =500
                    Top ={y_pos}
                    Width =2000
                    Height =315
                    Name ="{field}_Label"
                    Caption ="{field}:"
                    GUID = Begin
                        0x{uuid.uuid4().hex}
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2700
                    Top ={y_pos}
                    Width =4000
                    Height =315
                    Name ="{field}"
                    ControlSource ="{field}"
                    GUID = Begin
                        0x{uuid.uuid4().hex}
                    End
                End"""
                
                # Add to NameMap
                rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
                field_hex = f"{field}_Label".encode('utf-16le').hex()
                namemap_entries.append(f"0x{rand_hex}{len(f'{field}_Label'):02x}000000{field_hex}")
                
                rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
                field_hex = field.encode('utf-16le').hex()
                namemap_entries.append(f"0x{rand_hex}{len(field):02x}000000{field_hex}")
                
                y_pos += 400
                
        else:  # tabular layout (default)
            # Header controls
            header_controls = ""
            detail_controls = ""
            namemap_entries = []
            x_pos = 500
            
            for i, field in enumerate(fields):
                # Header label
                header_controls += f"""
                Begin Label
                    OverlapFlags =85
                    Left ={x_pos}
                    Top =200
                    Width =1500
                    Height =315
                    Name ="{field}_Header"
                    Caption ="{field}"
                    GUID = Begin
                        0x{uuid.uuid4().hex}
                    End
                End"""
                
                # Detail textbox
                detail_controls += f"""
                Begin TextBox
                    OverlapFlags =85
                    Left ={x_pos}
                    Top =200
                    Width =1500
                    Height =315
                    Name ="{field}"
                    ControlSource ="{field}"
                    GUID = Begin
                        0x{uuid.uuid4().hex}
                    End
                End"""
                
                # Add to NameMap
                rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
                field_hex = f"{field}_Header".encode('utf-16le').hex()
                namemap_entries.append(f"0x{rand_hex}{len(f'{field}_Header'):02x}000000{field_hex}")
                
                rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
                field_hex = field.encode('utf-16le').hex()
                namemap_entries.append(f"0x{rand_hex}{len(field):02x}000000{field_hex}")
                
                x_pos += 1600
            
            controls_text = f"""
        Begin Section
            Height =600
            Name ="ReportHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =500
                    Top =200
                    Width =6000
                    Height =400
                    Name ="Title"
                    Caption ="__REPORT_NAME_PLACEHOLDER__"
                    FontSize =14
                    FontWeight =700
                    GUID = Begin
                        0x{uuid.uuid4().hex}
                    End
                End
            End
        End
        Begin Section
            Height =600
            Name ="PageHeader"
            Begin
                {header_controls}
            End
        End
        Begin Section
            Height =400
            Name ="Detail"
            Begin
                {detail_controls}
            End
        End"""
        
        # Add Title to NameMap
        rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
        field_hex = "Title".encode('utf-16le').hex()
        namemap_entries.append(f"0x{rand_hex}05000000{field_hex}")
        
        # NameMap
        namemap_text = ",\n        ".join(namemap_entries) + ",\n        0x000000000000000000000000000000000c000000050000000000000000000000000000000000"
        
        if report_type.lower() == "columnar":
            template = f"""Version =21
VersionRequired =20
PublishOption =1
Checksum ={random.randint(-2000000000, 2000000000)}
Begin Report
    Width =7400
    PictureAlignment =2
    GUID = Begin
        0x{report_guid}
    End
    NameMap = Begin
        {namemap_text}
    End
    RecordSource ="{record_source}"
    Caption ="__REPORT_NAME_PLACEHOLDER__"
    Begin
        Begin Section
            Height ={y_pos + 200}
            Name ="Detail"
            Begin
                {controls_text}
            End
        End
    End
End"""
        else:  # tabular
            template = f"""Version =21
VersionRequired =20
PublishOption =1
Checksum ={random.randint(-2000000000, 2000000000)}
Begin Report
    Width =7400
    PictureAlignment =2
    GUID = Begin
        0x{report_guid}
    End
    NameMap = Begin
        {namemap_text}
    End
    RecordSource ="{record_source}"
    Caption ="__REPORT_NAME_PLACEHOLDER__"
    Begin
        {controls_text}
    End
End"""
        
        return template
        
    except Exception as e:
        raise Exception(f"Error generating report template: {e}")

def _create_report_from_template_internal(db_name: str, report_name: str, report_text: str) -> str:
    """Internal helper function to create report from template using new pattern."""
    # Replace placeholder if it exists
    if "__REPORT_NAME_PLACEHOLDER__" in report_text:
        report_text = report_text.replace("__REPORT_NAME_PLACEHOLDER__", report_name)
    
    def operation(access):
        """Inner function to create the report"""
        logger.info(f"Creating report: {report_name}")
        
        # Write to temp file
        temp_file_path = None
        try:
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix=".txt", encoding='utf-8') as tf:
                tf.write(report_text)
                temp_file_path = tf.name

            AC_REPORT = 3
            
            # Delete existing report if it exists
            try:
                access.DoCmd.DeleteObject(AC_REPORT, report_name)
                logger.debug(f"Deleted existing report: {report_name}")
            except Exception:
                logger.debug(f"Report {report_name} doesn't exist (creating new)")

            # Load report from text file
            access.LoadFromText(AC_REPORT, report_name, temp_file_path)
            logger.info(f"Report '{report_name}' created successfully")

            return f"Report '{report_name}' created successfully in database '{db_name}'."
            
        finally:
            if temp_file_path and os.path.exists(temp_file_path):
                os.remove(temp_file_path)
                logger.debug("Temp file cleaned up")
    
    try:
        path = get_db_path(db_name)
        
        # Check for lock
        if is_database_locked(path):
            success, message = wait_for_lock_release(path)
            if not success:
                raise Exception(message)
        
        result = _with_access_database(db_name, operation)
        return result
        
    except Exception as e:
        logger.error(f"Error creating report '{report_name}': {e}")
        raise Exception(f"Error creating report from template: {e}")

@mcp.tool
def create_report_from_source(db_name: str, report_name: str, record_source: str, report_type: str = "tabular") -> str:
    """Creates a complete Access report from a table or query in a single step.

    This tool combines template generation and creation, making it more reliable.

    Args:
        db_name: The name of the database file (e.g., 'inventory.accdb').
        report_name: The name to save the report as (e.g., 'ProductsReport').
        record_source: The name of the table or saved query the report is based on.
        report_type: Type of report layout - 'tabular' (default) or 'columnar'.
    """
    try:
        # Step 1: Generate the report template using internal helper
        report_text = _generate_report_template_internal(db_name, record_source, report_type)
        
        # Step 2: Create the report using internal helper
        result = _create_report_from_template_internal(db_name, report_name, report_text)
        
        return result

    except Exception as e:
        return f"An unexpected error occurred in create_report_from_source: {e}"

@mcp.tool
def generate_report_template(db_name: str, record_source: str, report_type: str = "tabular") -> str:
    """Generate a text template for an Access report that can be customized and created.
    
    Args:
        db_name: The name of the database file
        record_source: The name of the table or saved query the report is based on
        report_type: Type of report layout - 'tabular' or 'columnar'
    """
    try:
        template = _generate_report_template_internal(db_name, record_source, report_type)
        
        return f"""Report template generated successfully for {report_type} layout.
IMPORTANT: 
1. Replace '__REPORT_NAME_PLACEHOLDER__' with the desired report name.
2. Review and customize the template below as needed.
3. Pass the entire final text content to the 'create_report_from_template' tool.

--- TEMPLATE BEGIN ---
{template}
--- TEMPLATE END ---"""
        
    except Exception as e:
        return f"Error generating report template: {e}"

@mcp.tool
def create_report_from_template(db_name: str, report_name: str, report_text: str) -> str:
    """Create an Access report from a text template definition.
    
    Args:
        db_name: The name of the database file
        report_name: The name to save the report as
        report_text: The complete text definition of the report
    """
    try:
        return _create_report_from_template_internal(db_name, report_name, report_text)
    except Exception as e:
        return f"Error creating report from template: {str(e)}"
            
if __name__ == "__main__":
    mcp.run()
