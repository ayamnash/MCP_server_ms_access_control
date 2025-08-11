
import os
import pyodbc
from fastmcp import FastMCP
from win32com.client import Dispatch
import win32com.client
import uuid
import random
import tempfile
import re # <--- FIX #2

mcp = FastMCP("Flexible Access DB MCP")

# --- State Tracking ---
_template_generated = False
_last_template_type = None

# --- Helper Functions ---

# FIXED get_db_path function
def get_db_path(db_name: str) -> str:
    """Gets the full path for the database. Handles both absolute and relative paths."""
    # If the path is already absolute (e.g., "F:\...") use it directly.
    if os.path.isabs(db_name):
        if not db_name.lower().endswith(".accdb"):
            db_name += ".accdb"
        return db_name
    
    # Otherwise, build the path in the user's home directory.
    if not db_name.lower().endswith(".accdb"):
        db_name += ".accdb"
    return os.path.join(os.path.expanduser("~"), db_name)

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

# ... (the rest of your code from the previous correct answer is fine)
# Make sure to copy the full script provided in the previous response,
# as this snippet only shows the changed parts. The key is the new get_db_path
# and the added `import re`.
# I will paste the full code here again for clarity.

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
    
    # Restore protected quotes
    for i, protected in enumerate(protected_patterns):
        sql = sql.replace(f"__PROTECTED_QUOTE_{i}__", protected)
    
    return sql

@mcp.tool
def save_query(db_name: str, query_name: str, sql: str) -> str:
    """Save or overwrite a named query in an Access database.
    Automatically fixes common Access SQL syntax issues like double quotes."""
    try:
        path = get_db_path(db_name)
        
        # Fix Access SQL syntax issues
        sql_fixed = fix_access_sql_syntax(sql)
        
        # For COM interface, we need to escape double quotes in the SQL
        # This is specifically for cases like Format(field, "yyyy-mm-dd")
        sql_escaped = sql_fixed.replace('"', '""')
        
        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(path)
        dao = access.CurrentDb()
        
        # Delete existing query if it exists
        try:
            dao.QueryDefs.Delete(query_name)
        except Exception:
            pass  # Query doesn't exist, that's fine
        
        # Create new query with escaped SQL
        dao.CreateQueryDef(query_name, sql_escaped)
        
        access.CloseCurrentDatabase()
        access.Quit()
        return f"Query '{query_name}' saved successfully."
    except Exception as e:
        return f"Error saving query: {str(e)}"



# @mcp.tool
# def save_query(db_name: str, query_name: str, sql: str) -> str:
#     """Saves a named query (a 'View') inside the Access database."""
#     try:
#         path = get_db_path(db_name)
#         access = win32com.client.Dispatch("Access.Application")
#         access.Visible = False
#         access.OpenCurrentDatabase(path)
#         dao = access.CurrentDb()
#         try:
#             dao.QueryDefs.Delete(query_name)
#         except Exception:
#             pass
#         dao.CreateQueryDef(query_name, sql)
#         access.CloseCurrentDatabase()
#         access.Quit()
#         return f"Saved query '{query_name}' in {db_name}"
#     except Exception as e:
#         return f"Error saving query: {str(e)}"
# Replace your old save_query function with this new one

# @mcp.tool
# def save_query(db_name: str, query_name: str, sql: str) -> str:
#     """
#     Saves a named query (a 'View') inside the Access database using the direct ADOX method.
#     This is more reliable than automating the full Access application.
#     """
#     path = get_db_path(db_name)
#     conn = None
    
#     try:
#         # 1. Create a direct connection to the database engine
#         conn = win32com.client.Dispatch("ADODB.Connection")
#         conn_str = f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};"
#         conn.Open(conn_str)

#         # 2. Create a Catalog object to manage database schema
#         cat = win32com.client.Dispatch("ADOX.Catalog")
#         cat.ActiveConnection = conn

#         # 3. If the query already exists, delete it first.
#         # This prevents an error if you try to save over an existing query.
#         try:
#             cat.Views.Delete(query_name)
#         except Exception:
#             pass  # Ignore if the query doesn't exist

#         # 4. Create the new query (View) using a Command object
#         cmd = win32com.client.Dispatch("ADODB.Command")
#         cmd.CommandText = sql
        
#         # 5. Append the command as a new View to the catalog
#         cat.Views.Append(query_name, cmd)

#         return f"Saved query '{query_name}' in {db_name}"

#     except Exception as e:
#         return f"Error saving query: {str(e)}"
        
#     finally:
#         # 6. Clean up the connection
#         if conn and conn.State == 1: # 1 means the connection is open
#             conn.Close()


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
        controls_text += f"""
                Begin Subform
                    OverlapFlags =85
                    Left =500
                    Top ={y_pos + 200}
                    Width =10000
                    Height =4000
                    TabIndex = {len(fields_to_show)}
                    Name ="{re.sub(r'^Form\.', '', subform_object_name)}"
                    SourceObject ="{subform_object_name}"
                    LinkChildFields ="{link_child_field}"
                    LinkMasterFields ="{link_master_field}"
                    GUID = Begin
                        0x{uuid.uuid4().hex}
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
    """
    STEP 2/2 for creating a form. Creates an Access form from its text definition.
    This tool will automatically correct/generate the NameMap and GUIDs based on the
    controls found in the form_text, making it robust against LLM-generated errors.
    
    Args:
        db_name: The name of the database file (e.g., 'inventory.accdb'). Can be an absolute path.
        form_name: The name to save the form as (e.g., 'ProductsForm').
        form_text: The complete text definition of the form.
    """
    
    # --- PRE-PROCESSING AND VALIDATION ---
    try:
        # 1. Replace placeholder if it exists
        if "__FORM_NAME_PLACEHOLDER__" in form_text:
             form_text = form_text.replace("__FORM_NAME_PLACEHOLDER__", form_name)

        # 2. Find all controls with a 'Name' property
        # This regex finds 'Name ="ControlName"' and captures 'ControlName'
        control_names = re.findall(r'^\s*Name\s*=\s*"([^"]+)"', form_text, re.MULTILINE)
        if not control_names:
            return "Error: Could not find any named controls in the form text to build a NameMap."

        # 3. Generate a fresh, correct NameMap
        namemap_entries = []
        for name in control_names:
            rand_hex = ''.join(random.choices('0123456789abcdef', k=32))
            field_hex = name.encode('utf-16le').hex()
            # The format is: {random_guid}{hex_len_of_name}{padding}{hex_encoded_name}
            namemap_entries.append(f"0x{rand_hex}{len(name):02x}000000{field_hex}")
        
        # Add the required terminator for the NameMap
        namemap_terminator = "0x000000000000000000000000000000000c000000050000000000000000000000000000000000"
        namemap_entries.append(namemap_terminator)
        
        new_namemap_text = "NameMap = Begin\n        " + ",\n        ".join(namemap_entries) + "\n    End"

        # 4. Replace the old NameMap in the text with our new, correct one
        form_text = re.sub(r'NameMap\s*=\s*Begin.*?End', new_namemap_text, form_text, flags=re.DOTALL)

        # 5. Find and fix all GUIDs. Replace any invalid ones.
        def replace_guid(match):
            guid_content = match.group(1).strip().replace('0x', '')
            if len(guid_content) == 32 and all(c in '0123456789abcdefABCDEF' for c in guid_content):
                return match.group(0) # It's a valid GUID, leave it alone
            else:
                # It's invalid (wrong length, bad chars), so replace it
                return f"GUID = Begin\n            0x{uuid.uuid4().hex}\n        End"

        form_text = re.sub(r'GUID\s*=\s*Begin(.*?)End', replace_guid, form_text, flags=re.DOTALL)

    except Exception as e:
        return f"An unexpected error occurred during pre-processing: {e}"


    # --- THE REST OF THE FUNCTION IS THE SAME ---
    path = get_db_path(db_name)
    temp_file_path = None
    
    try:
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix=".txt", encoding='utf-8') as tf:
            tf.write(form_text)
            temp_file_path = tf.name

        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(path)
        
        AC_FORM = 2
        
        try:
            access.DoCmd.DeleteObject(AC_FORM, form_name)
        except Exception:
            pass

        access.LoadFromText(AC_FORM, form_name, temp_file_path)

        access.CloseCurrentDatabase()
        access.Quit()

        global _template_generated, _last_template_type
        _template_generated = False
        _last_template_type = None

        return f"Form '{form_name}' created successfully in database '{db_name}'."

    except Exception as e:
        return f"Error creating form from text: {str(e)}"
    finally:
        if temp_file_path and os.path.exists(temp_file_path):
            os.remove(temp_file_path)

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
        return f"Error listing VBA modules: {str(e)}"

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
            return f"Module '{module_name}' not found"
            
    except Exception as e:
        return f"Error reading VBA module: {str(e)}"

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
        return f"VBA module '{module_name}' {action} successfully"
        
    except Exception as e:
        return f"Error writing VBA module: {str(e)}"

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
            return f"VBA module '{module_name}' deleted successfully"
        else:
            return f"Module '{module_name}' not found"
            
    except Exception as e:
        return f"Error deleting VBA module: {str(e)}"

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
        
        return f"Function '{function_name}' executed successfully. Result: {result}"
        
    except Exception as e:
        return f"Error running VBA function: {str(e)}"
            
if __name__ == "__main__":
    mcp.run()
