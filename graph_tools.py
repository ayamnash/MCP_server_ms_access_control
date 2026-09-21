"""
Database Knowledge Graph and Subgraph Context Engine for Microsoft Access.
Provides Graphify-like entity relationship graph extraction, token-reducing
subgraph context querying for AI agents, and interactive HTML visualization.
Zero external graph library dependencies (pure Python).
"""

import os
import re
import json
import math
import logging
from collections import deque, defaultdict
from typing import Dict, List, Any, Optional, Set, Tuple

logger = logging.getLogger(__name__)

# DAO Field Data Type Mapping constants
DAO_DATA_TYPES = {
    1: "Boolean (Yes/No)",
    2: "Byte",
    3: "Integer",
    4: "Long Integer",
    5: "Currency",
    6: "Single",
    7: "Double",
    8: "Date/Time",
    9: "Binary",
    10: "Text",
    11: "Long Binary (OLE Object)",
    12: "Memo (Long Text)",
    15: "GUID",
    16: "BigInt",
    17: "VarNumeric",
    18: "Char",
    19: "Numeric",
    20: "Decimal",
    22: "Time",
    23: "TimeStamp",
}

# DAO Field Attribute Flags
DB_AUTO_INCR_FIELD = 16


def _clean_sql_for_tables(sql: str) -> List[str]:
    """Extract table and view names referenced in a SQL query using regex."""
    if not sql:
        return []
    tables = set()
    
    # Pattern to find tables in FROM and JOIN clauses
    # Supports [Table Name] or TableName
    pattern = r'\b(?:FROM|JOIN)\s+(?:\[([^\]]+)\]|([a-zA-Z0-9_]+))'
    matches = re.findall(pattern, sql, re.IGNORECASE)
    for m in matches:
        t = m[0] if m[0] else m[1]
        if t and not t.upper().startswith("MSYS"):
            tables.add(t)
            
    return sorted(list(tables))


class AccessGraphEngine:
    """Extracts, indexes, analyzes, and visualizes MS Access database structures."""

    def __init__(self, db_path: str):
        self.db_path = db_path
        self.nodes: Dict[str, Dict[str, Any]] = {}
        self.edges: List[Dict[str, Any]] = []
        self.adjacency: Dict[str, Set[str]] = defaultdict(set)
        self.table_schemas: Dict[str, Dict[str, Any]] = {}

    @classmethod
    def from_graphify_json(cls, json_path: str) -> "AccessGraphEngine":
        """Instantly load a Knowledge Graph from an existing graph.json file.
        Enables 100% offline graph querying in under 0.01 seconds without connecting to MS Access."""
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        db_name = data.get("database", os.path.basename(json_path))
        engine = cls(db_path=db_name)

        for node in data.get("nodes", []):
            nid = node["id"]
            engine.nodes[nid] = node
            if node.get("type") == "table":
                tname = node.get("object_name") or node.get("label")
                engine.table_schemas[tname] = {
                    "name": tname,
                    "fields": node.get("fields", []),
                    "pk": node.get("pk", [])
                }

        for edge in data.get("edges", []):
            engine.edges.append(edge)
            src = edge["source"]
            dst = edge["target"]
            engine.adjacency[src].add(dst)
            engine.adjacency[dst].add(src)

        return engine

    def extract_from_dao(self, dao_source) -> None:
        """Extract full schema, relations, queries, forms, reports, and modules via COM / DAO.
        Supports both direct DAO Database (headless, fast, no GUI modals) and Access.Application."""
        if hasattr(dao_source, "CurrentDb"):
            dao = dao_source.CurrentDb()
        else:
            dao = dao_source
        
        # 1. Extract Tables and Columns
        for tbl in dao.TableDefs:
            tname = tbl.Name
            if tname.startswith("MSys") or tname.startswith("~"):
                continue

            fields_list = []
            pk_fields = set()

            # Find Primary Key indexes
            try:
                for idx in tbl.Indexes:
                    if idx.Primary:
                        for f in idx.Fields:
                            pk_fields.add(f.Name)
            except Exception as e:
                logger.debug(f"Could not read indexes for table {tname}: {e}")

            # Read Fields
            for fld in tbl.Fields:
                is_autonumber = False
                try:
                    is_autonumber = bool(fld.Attributes & DB_AUTO_INCR_FIELD)
                except Exception:
                    pass

                type_code = fld.Type
                type_name = DAO_DATA_TYPES.get(type_code, f"Type({type_code})")
                if is_autonumber:
                    type_name = "AutoNumber (Counter)"

                is_pk = fld.Name in pk_fields
                is_required = False
                try:
                    is_required = bool(fld.Required)
                except Exception:
                    pass

                fields_list.append({
                    "name": fld.Name,
                    "type": type_name,
                    "type_code": type_code,
                    "size": getattr(fld, "Size", 0),
                    "is_pk": is_pk,
                    "is_autonumber": is_autonumber,
                    "required": is_required
                })

            self.table_schemas[tname] = {
                "name": tname,
                "fields": fields_list,
                "record_count": getattr(tbl, "RecordCount", 0),
                "pk": list(pk_fields)
            }

            node_id = f"table:{tname}"
            self.nodes[node_id] = {
                "id": node_id,
                "label": tname,
                "type": "table",
                "object_name": tname,
                "fields": fields_list,
                "field_count": len(fields_list),
                "pk": list(pk_fields),
                "record_count": getattr(tbl, "RecordCount", 0)
            }

        # 2. Extract Relationships (Relations collection)
        for rel in dao.Relations:
            rname = rel.Name
            primary_tbl = rel.Table
            foreign_tbl = rel.ForeignTable

            if primary_tbl.startswith("MSys") or foreign_tbl.startswith("MSys"):
                continue

            field_pairs = []
            try:
                for rfld in rel.Fields:
                    field_pairs.append({
                        "primary_field": rfld.Name,
                        "foreign_field": rfld.ForeignName
                    })
            except Exception as e:
                logger.debug(f"Error reading relation fields for {rname}: {e}")

            attr = getattr(rel, "Attributes", 0)
            cascade_update = bool(attr & 256)
            cascade_delete = bool(attr & 4096)
            unique = bool(attr & 1)

            src_node = f"table:{primary_tbl}"
            dst_node = f"table:{foreign_tbl}"

            # Make sure both tables exist in nodes
            if src_node in self.nodes and dst_node in self.nodes:
                edge_label = ", ".join(f"{p['primary_field']} = {p['foreign_field']}" for p in field_pairs)
                edge = {
                    "source": src_node,
                    "target": dst_node,
                    "type": "FOREIGN_KEY",
                    "label": edge_label or "Relationship",
                    "relation_name": rname,
                    "field_pairs": field_pairs,
                    "cascade_update": cascade_update,
                    "cascade_delete": cascade_delete,
                    "unique": unique,
                    "cardinality": "1:1" if unique else "1:N"
                }
                self.edges.append(edge)
                self.adjacency[src_node].add(dst_node)
                self.adjacency[dst_node].add(src_node)

        # 3. Inferred Relationships (Foreign Key heuristics if relations not defined)
        self._infer_missing_foreign_keys()

        # 4. Extract Saved Queries (QueryDefs)
        for qdf in dao.QueryDefs:
            qname = qdf.Name
            if qname.startswith("~"):
                continue

            sql_text = getattr(qdf, "SQL", "").strip()
            referenced_tables = _clean_sql_for_tables(sql_text)

            node_id = f"query:{qname}"
            self.nodes[node_id] = {
                "id": node_id,
                "label": qname,
                "type": "query",
                "object_name": qname,
                "sql": sql_text,
                "referenced_tables": referenced_tables
            }

            for rtable in referenced_tables:
                tbl_node = f"table:{rtable}"
                if tbl_node in self.nodes:
                    self.edges.append({
                        "source": node_id,
                        "target": tbl_node,
                        "type": "QUERIES_TABLE",
                        "label": "Reads From"
                    })
                    self.adjacency[node_id].add(tbl_node)
                    self.adjacency[tbl_node].add(node_id)

        # 5. Extract Forms (Supports both headless DAO Containers and Access.Application)
        try:
            if hasattr(dao, "Containers"):
                for doc in dao.Containers("Forms").Documents:
                    fname = doc.Name
                    node_id = f"form:{fname}"
                    self.nodes[node_id] = {
                        "id": node_id,
                        "label": fname,
                        "type": "form",
                        "object_name": fname
                    }
            elif hasattr(dao_source, "CurrentProject"):
                for form in dao_source.CurrentProject.AllForms:
                    fname = form.Name
                    node_id = f"form:{fname}"
                    self.nodes[node_id] = {
                        "id": node_id,
                        "label": fname,
                        "type": "form",
                        "object_name": fname
                    }
        except Exception as e:
            logger.debug(f"Could not read forms collection: {e}")

        # 6. Extract Reports (Supports both headless DAO Containers and Access.Application)
        try:
            if hasattr(dao, "Containers"):
                for doc in dao.Containers("Reports").Documents:
                    rname = doc.Name
                    node_id = f"report:{rname}"
                    self.nodes[node_id] = {
                        "id": node_id,
                        "label": rname,
                        "type": "report",
                        "object_name": rname
                    }
            elif hasattr(dao_source, "CurrentProject"):
                for rep in dao_source.CurrentProject.AllReports:
                    rname = rep.Name
                    node_id = f"report:{rname}"
                    self.nodes[node_id] = {
                        "id": node_id,
                        "label": rname,
                        "type": "report",
                        "object_name": rname
                    }
        except Exception as e:
            logger.debug(f"Could not read reports collection: {e}")

        # 7. Extract VBA Modules (Supports both headless DAO Containers and Access.Application)
        try:
            if hasattr(dao, "Containers"):
                for doc in dao.Containers("Modules").Documents:
                    mname = doc.Name
                    node_id = f"module:{mname}"
                    self.nodes[node_id] = {
                        "id": node_id,
                        "label": mname,
                        "type": "module",
                        "object_name": mname,
                        "module_type": "Standard",
                        "line_count": 0
                    }
            elif hasattr(dao_source, "VBE"):
                vbe_proj = dao_source.VBE.VBProjects(1)
                for i in range(1, vbe_proj.VBComponents.Count + 1):
                    comp = vbe_proj.VBComponents(i)
                    mname = comp.Name
                    mtype = {1: "Standard", 2: "Class", 3: "Form", 100: "Document"}.get(comp.Type, "Module")
                    node_id = f"module:{mname}"
                    line_count = 0
                    if comp.CodeModule:
                        line_count = comp.CodeModule.CountOfLines
                    self.nodes[node_id] = {
                        "id": node_id,
                        "label": mname,
                        "type": "module",
                        "object_name": mname,
                        "module_type": mtype,
                        "line_count": line_count
                    }
        except Exception as e:
            logger.debug(f"Could not read VBA components: {e}")

    def extract_from_pyodbc(self, pyodbc_conn) -> None:
        """Fallback extraction using ODBC driver when COM is not active."""
        cursor = pyodbc_conn.cursor()

        # 1. Get Tables
        tables = [row.table_name for row in cursor.tables(tableType='TABLE') if not row.table_name.startswith("MSys")]

        for tname in tables:
            fields_list = []
            pk_fields = set()

            # Primary keys
            try:
                for pk_row in cursor.primaryKeys(table=tname):
                    pk_fields.add(pk_row.column_name)
            except Exception:
                pass

            # Columns
            try:
                for col_row in cursor.columns(table=tname):
                    cname = col_row.column_name
                    ctype = col_row.type_name
                    is_pk = cname in pk_fields
                    is_autonumber = "COUNTER" in str(ctype).upper() or "AUTOINCREMENT" in str(ctype).upper()
                    fields_list.append({
                        "name": cname,
                        "type": ctype,
                        "is_pk": is_pk,
                        "is_autonumber": is_autonumber,
                        "required": bool(col_row.nullable == 0)
                    })
            except Exception as e:
                logger.debug(f"Error reading columns for {tname}: {e}")

            self.table_schemas[tname] = {
                "name": tname,
                "fields": fields_list,
                "pk": list(pk_fields)
            }

            node_id = f"table:{tname}"
            self.nodes[node_id] = {
                "id": node_id,
                "label": tname,
                "type": "table",
                "object_name": tname,
                "fields": fields_list,
                "field_count": len(fields_list),
                "pk": list(pk_fields)
            }

        # 2. Foreign keys via ODBC
        for tname in tables:
            try:
                for fk_row in cursor.foreignKeys(foreignTable=tname):
                    ptbl = fk_row.pktable_name
                    pcol = fk_row.pkcolumn_name
                    fcol = fk_row.fkcolumn_name

                    src_node = f"table:{ptbl}"
                    dst_node = f"table:{tname}"
                    if src_node in self.nodes and dst_node in self.nodes:
                        edge = {
                            "source": src_node,
                            "target": dst_node,
                            "type": "FOREIGN_KEY",
                            "label": f"{pcol} = {fcol}",
                            "field_pairs": [{"primary_field": pcol, "foreign_field": fcol}],
                            "cardinality": "1:N"
                        }
                        self.edges.append(edge)
                        self.adjacency[src_node].add(dst_node)
                        self.adjacency[dst_node].add(src_node)
            except Exception:
                pass

        # 3. Infer missing foreign keys
        self._infer_missing_foreign_keys()

    def _infer_missing_foreign_keys(self) -> None:
        """Heuristic discovery of relationships between tables based on naming conventions."""
        existing_pairs = set()
        for edge in self.edges:
            if edge.get("type") == "FOREIGN_KEY":
                existing_pairs.add((edge["source"], edge["target"]))
                existing_pairs.add((edge["target"], edge["source"]))

        # Check each table against every other table
        table_names = list(self.table_schemas.keys())
        for i, t1 in enumerate(table_names):
            t1_schema = self.table_schemas[t1]
            t1_pks = t1_schema.get("pk", [])
            t1_node = f"table:{t1}"

            for j, t2 in enumerate(table_names):
                if i == j:
                    continue
                t2_node = f"table:{t2}"
                if (t1_node, t2_node) in existing_pairs:
                    continue

                t2_schema = self.table_schemas[t2]
                t2_field_names = [f["name"] for f in t2_schema.get("fields", [])]

                # Match patterns:
                # 1. t1 is 'Customers' and t2 has 'CustomerID' or 'Customer_ID'
                # 2. t1 has PK 'ID' and t2 has 't1ID'
                singular_t1 = t1.rstrip("s").rstrip("S")
                potential_fk_names = [
                    f"{t1}ID", f"{t1}_ID", f"{singular_t1}ID", f"{singular_t1}_ID",
                    f"{t1}Id", f"{singular_t1}Id"
                ]

                for fk_candidate in potential_fk_names:
                    for f in t2_schema.get("fields", []):
                        if f["name"].lower() == fk_candidate.lower():
                            # Determine matched PK in t1
                            matched_pk = t1_pks[0] if t1_pks else "ID"
                            for t1_f in t1_schema.get("fields", []):
                                if t1_f["name"].lower() in ["id", f"{t1.lower()}id", f"{singular_t1.lower()}id"]:
                                    matched_pk = t1_f["name"]
                                    break

                            edge = {
                                "source": t1_node,
                                "target": t2_node,
                                "type": "FOREIGN_KEY",
                                "label": f"{matched_pk} = {f['name']} (Inferred)",
                                "field_pairs": [{"primary_field": matched_pk, "foreign_field": f["name"]}],
                                "inferred": True,
                                "cardinality": "1:N"
                            }
                            self.edges.append(edge)
                            self.adjacency[t1_node].add(t2_node)
                            self.adjacency[t2_node].add(t1_node)
                            existing_pairs.add((t1_node, t2_node))
                            break

    def compute_metrics_and_communities(self) -> Dict[str, Any]:
        """Compute degrees, identify god/hub nodes, and detect connected communities."""
        # Calculate degrees
        degrees = {}
        for nid in self.nodes:
            degrees[nid] = len(self.adjacency.get(nid, set()))
            self.nodes[nid]["degree"] = degrees[nid]

        # God nodes: Top 20% highest degree table nodes (min degree >= 2)
        table_nodes = [n for n in self.nodes.values() if n["type"] == "table"]
        sorted_tables = sorted(table_nodes, key=lambda x: x.get("degree", 0), reverse=True)
        god_nodes = [t["object_name"] for t in sorted_tables if t.get("degree", 0) >= 2][:5]

        # Connected components / Communities via BFS
        visited = set()
        communities = {}
        comm_id = 0

        for nid in self.nodes:
            if nid not in visited:
                queue = deque([nid])
                visited.add(nid)
                comp = []
                while queue:
                    curr = queue.popleft()
                    comp.append(curr)
                    for neighbor in self.adjacency.get(curr, set()):
                        if neighbor not in visited:
                            visited.add(neighbor)
                            queue.append(neighbor)
                communities[comm_id] = comp
                for member in comp:
                    self.nodes[member]["community"] = comm_id
                comm_id += 1

        return {
            "god_nodes": god_nodes,
            "communities_count": len(communities),
            "total_nodes": len(self.nodes),
            "total_edges": len(self.edges)
        }

    def get_subgraph_context(
        self,
        task_description: str,
        relevant_tables: Optional[str] = None,
        depth: int = 1,
        token_budget: int = 1500
    ) -> str:
        """
        Produce a compact, high-density schema cheat sheet for AI operations.
        Massively reduces token consumption by extracting only targeted tables,
        their immediate 1-hop FK relationships, constraints, and data types.
        """
        # Parse search keywords from task description and explicit tables
        keywords = set()
        if relevant_tables:
            for t in relevant_tables.split(","):
                t_clean = t.strip().lower()
                if t_clean:
                    keywords.add(t_clean)

        # Tokenize task description
        clean_words = re.findall(r'[a-zA-Z0-9_]+', task_description.lower())
        for w in clean_words:
            if len(w) > 2 and w not in ["the", "and", "for", "with", "add", "edit", "insert", "update", "delete", "table", "from", "into", "records", "data"]:
                keywords.add(w)

        # Score nodes for relevance
        scores = {}
        for nid, node in self.nodes.items():
            name = node.get("object_name", "").lower()
            score = 0
            # Direct table name match
            for kw in keywords:
                if kw == name:
                    score += 50
                elif kw in name:
                    score += 20

            # Match on column names
            if node["type"] == "table":
                for fld in node.get("fields", []):
                    fname = fld["name"].lower()
                    for kw in keywords:
                        if kw == fname:
                            score += 15
                        elif kw in fname:
                            score += 5

            if score > 0:
                scores[nid] = score

        # If no keywords matched, fall back to all tables up to budget
        if not scores:
            for nid, node in self.nodes.items():
                if node["type"] == "table":
                    scores[nid] = 1

        # Pick top seed nodes
        sorted_seeds = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        core_nodes = set([item[0] for item in sorted_seeds[:5]])

        # Expand neighborhood up to specified depth
        expanded_nodes = set(core_nodes)
        for _ in range(depth):
            current_layer = list(expanded_nodes)
            for curr in current_layer:
                for neighbor in self.adjacency.get(curr, set()):
                    expanded_nodes.add(neighbor)

        # Separate into tables and other objects
        selected_tables = [nid for nid in expanded_nodes if self.nodes.get(nid, {}).get("type") == "table"]
        selected_queries = [nid for nid in expanded_nodes if self.nodes.get(nid, {}).get("type") == "query"]

        # Gather relevant edges between selected nodes
        relevant_edges = []
        for edge in self.edges:
            if edge["source"] in expanded_nodes and edge["target"] in expanded_nodes:
                relevant_edges.append(edge)

        # Format compact Markdown cheat sheet
        lines = []
        lines.append("### Targeted Database Schema Context (Graph Subgraph)")
        lines.append(f"**Focused Tables ({len(selected_tables)})** | Target Task: *{task_description}*\n")

        # 1. Foreign Key Relationships Table
        if relevant_edges:
            lines.append("#### Active Relationships & Foreign Keys:")
            for edge in relevant_edges:
                src_tbl = self.nodes.get(edge["source"], {}).get("object_name", edge["source"])
                dst_tbl = self.nodes.get(edge["target"], {}).get("object_name", edge["target"])
                lbl = edge.get("label", "")
                card = edge.get("cardinality", "1:N")
                inferred = " *(Inferred)*" if edge.get("inferred") else ""
                lines.append(f"- **{src_tbl}** --[{card}]--> **{dst_tbl}**: `{lbl}`{inferred}")
            lines.append("")

        # 2. Table Column Definitions
        lines.append("#### Table Definitions & Constraints:")
        for t_nid in selected_tables:
            node = self.nodes[t_nid]
            tname = node["object_name"]
            fields = node.get("fields", [])
            pk = node.get("pk", [])

            lines.append(f"**Table: `[{tname}]`** (Primary Key: `{', '.join(pk) if pk else 'None'}`)")
            field_defs = []
            for f in fields:
                flags = []
                if f.get("is_pk"):
                    flags.append("PK")
                if f.get("is_autonumber"):
                    flags.append("AutoNumber - DO NOT INSERT")
                if f.get("required") and not f.get("is_autonumber"):
                    flags.append("Required")

                flag_str = f" [{', '.join(flags)}]" if flags else ""
                field_defs.append(f"  - `{f['name']}`: {f['type']}{flag_str}")
            lines.append("\n".join(field_defs))
            lines.append("")

        # 3. Relevant Queries (if any)
        if selected_queries:
            lines.append("#### Connected Saved Queries:")
            for q_nid in selected_queries[:3]:
                qnode = self.nodes[q_nid]
                qname = qnode["object_name"]
                sql = qnode.get("sql", "").strip()
                # Truncate long SQL
                if len(sql) > 150:
                    sql = sql[:150] + "..."
                lines.append(f"- `[{qname}]`: `{sql}`")
            lines.append("")

        # 4. Critical Constraints & Tips
        lines.append("#### Execution Guidance:")
        lines.append("- Never insert into columns flagged as `AutoNumber`.")
        lines.append("- For date literals in SQL WHERE clauses, format as `#YYYY-MM-DD#`.")
        lines.append("- For multi-table JOINs in Access, wrap initial JOINs in parentheses: `FROM (TableA INNER JOIN TableB ON ...) INNER JOIN TableC ON ...`.")

        result = "\n".join(lines)
        
        # Approximate token budget check (1 token ~= 4 chars)
        max_chars = token_budget * 4
        if len(result) > max_chars:
            result = result[:max_chars] + "\n\n...[Truncated to strictly adhere to token budget]..."

        return result

    def find_path(self, source_table: str, target_table: str) -> Dict[str, Any]:
        """Find the shortest relationship path between two tables using BFS."""
        src_node = f"table:{source_table}"
        dst_node = f"table:{target_table}"

        if src_node not in self.nodes:
            return {"success": False, "error": f"Source table '{source_table}' not found in database."}
        if dst_node not in self.nodes:
            return {"success": False, "error": f"Target table '{target_table}' not found in database."}

        if src_node == dst_node:
            return {
                "success": True,
                "path": [source_table],
                "distance": 0,
                "join_sql": f"FROM [{source_table}]"
            }

        queue = deque([(src_node, [src_node], [])])
        visited = {src_node}

        while queue:
            curr, path, edge_path = queue.popleft()

            if curr == dst_node:
                # Reconstruct path
                table_names = [self.nodes[n]["object_name"] for n in path]
                
                # Build suggested SQL JOIN clause
                join_clauses = []
                current_from = f"[{table_names[0]}]"
                
                for i, edge in enumerate(edge_path):
                    next_table = table_names[i + 1]
                    pairs = edge.get("field_pairs", [])
                    if pairs:
                        on_clause = " AND ".join(
                            f"[{self.nodes[edge['source']]['object_name']}].[{p['primary_field']}] = [{self.nodes[edge['target']]['object_name']}].[{p['foreign_field']}]"
                            for p in pairs
                        )
                    else:
                        on_clause = f"[{self.nodes[edge['source']]['object_name']}].ID = [{next_table}].ID"
                    
                    if i == 0:
                        current_from = f"[{table_names[0]}] INNER JOIN [{next_table}] ON {on_clause}"
                    else:
                        current_from = f"({current_from}) INNER JOIN [{next_table}] ON {on_clause}"

                return {
                    "success": True,
                    "path": table_names,
                    "distance": len(edge_path),
                    "edges": edge_path,
                    "join_sql": f"FROM {current_from}"
                }

            for neighbor in self.adjacency.get(curr, set()):
                if neighbor not in visited:
                    visited.add(neighbor)
                    # Find matching edge
                    matched_edge = None
                    for edge in self.edges:
                        if (edge["source"] == curr and edge["target"] == neighbor) or \
                           (edge["source"] == neighbor and edge["target"] == curr):
                            matched_edge = edge
                            break
                    queue.append((neighbor, path + [neighbor], edge_path + [matched_edge or {}]))

        return {
            "success": False,
            "error": f"No relationship path connects '{source_table}' and '{target_table}'."
        }

    def to_graphify_json(self) -> Dict[str, Any]:
        """Convert graph to Graphify-standard node-link format."""
        self.compute_metrics_and_communities()
        return {
            "version": "1.0.0",
            "database": os.path.basename(self.db_path),
            "nodes": list(self.nodes.values()),
            "edges": self.edges,
            "metadata": {
                "total_tables": sum(1 for n in self.nodes.values() if n["type"] == "table"),
                "total_queries": sum(1 for n in self.nodes.values() if n["type"] == "query"),
                "total_forms": sum(1 for n in self.nodes.values() if n["type"] == "form"),
                "total_reports": sum(1 for n in self.nodes.values() if n["type"] == "report"),
                "total_modules": sum(1 for n in self.nodes.values() if n["type"] == "module"),
                "total_relationships": len([e for e in self.edges if e.get("type") == "FOREIGN_KEY"]),
            }
        }

    def generate_interactive_html(self, output_file: str) -> str:
        """
        Generate a rich, standalone interactive HTML visualizer with zero external dependencies.
        High-performance 2D Canvas force-directed graph with pan, zoom, search, filtering,
        physics toggles, and full-schema slide-out drawer.
        """
        graph_data = self.to_graphify_json()
        json_data_str = json.dumps(graph_data, ensure_ascii=False)
        db_filename = os.path.basename(self.db_path)

        html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Database Knowledge Graph - {db_filename}</title>
    <style>
        :root {{
            --bg-primary: #0b0f19;
            --bg-secondary: #131b2e;
            --bg-glass: rgba(19, 27, 46, 0.85);
            --border-color: rgba(255, 255, 255, 0.1);
            --text-primary: #f1f5f9;
            --text-secondary: #94a3b8;
            --accent-table: #10b981;
            --accent-query: #3b82f6;
            --accent-form: #8b5cf6;
            --accent-report: #f59e0b;
            --accent-module: #ec4899;
            --accent-pk: #fbbf24;
            --accent-fk: #06b6d4;
            --accent-glow: rgba(59, 130, 246, 0.5);
        }}

        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        }}

        body {{
            background-color: var(--bg-primary);
            color: var(--text-primary);
            overflow: hidden;
            width: 100vw;
            height: 100vh;
        }}

        #graph-canvas {{
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            display: block;
            cursor: grab;
        }}

        #graph-canvas:active {{
            cursor: grabbing;
        }}

        /* Top Header Glassmorphism Bar */
        .header-bar {{
            position: absolute;
            top: 16px;
            left: 16px;
            right: 16px;
            height: 60px;
            background: var(--bg-glass);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 24px;
            z-index: 10;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.4);
        }}

        .brand-section {{
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .brand-badge {{
            background: linear-gradient(135deg, #3b82f6, #8b5cf6);
            padding: 6px 12px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 700;
            letter-spacing: 0.5px;
            text-transform: uppercase;
        }}

        .brand-title {{
            font-size: 17px;
            font-weight: 600;
            color: var(--text-primary);
        }}

        .stats-group {{
            display: flex;
            gap: 12px;
        }}

        .stat-pill {{
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-color);
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        .stat-dot {{
            width: 8px;
            height: 8px;
            border-radius: 50%;
        }}

        /* Controls floating toolbar */
        .controls-panel {{
            position: absolute;
            bottom: 24px;
            left: 24px;
            background: var(--bg-glass);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 16px;
            z-index: 10;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.4);
            display: flex;
            flex-direction: column;
            gap: 12px;
            width: 280px;
        }}

        .search-box {{
            width: 100%;
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 8px 12px;
            color: var(--text-primary);
            font-size: 13px;
            outline: none;
            transition: border-color 0.2s;
        }}

        .search-box:focus {{
            border-color: #3b82f6;
        }}

        .filter-row {{
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }}

        .filter-btn {{
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-color);
            color: var(--text-secondary);
            font-size: 11px;
            padding: 4px 10px;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s;
            user-select: none;
        }}

        .filter-btn.active {{
            background: rgba(59, 130, 246, 0.2);
            color: #60a5fa;
            border-color: #3b82f6;
        }}

        .action-buttons {{
            display: flex;
            gap: 8px;
        }}

        .btn-action {{
            flex: 1;
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid var(--border-color);
            color: var(--text-primary);
            font-size: 12px;
            padding: 6px;
            border-radius: 6px;
            cursor: pointer;
            text-align: center;
        }}

        .btn-action:hover {{
            background: rgba(255, 255, 255, 0.15);
        }}

        /* Slide-over Inspector Drawer */
        .inspector-drawer {{
            position: absolute;
            top: 92px;
            right: 16px;
            bottom: 24px;
            width: 420px;
            background: var(--bg-glass);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 24px;
            z-index: 10;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.5);
            display: flex;
            flex-direction: column;
            gap: 16px;
            transform: translateX(460px);
            transition: transform 0.3s cubic-bezier(0.16, 1, 0.3, 1);
            overflow-y: auto;
        }}

        .inspector-drawer.open {{
            transform: translateX(0);
        }}

        .drawer-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-bottom: 12px;
            border-bottom: 1px solid var(--border-color);
        }}

        .drawer-title {{
            font-size: 18px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .btn-close {{
            background: none;
            border: none;
            color: var(--text-secondary);
            font-size: 20px;
            cursor: pointer;
            padding: 4px;
        }}

        .badge-pill {{
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 10px;
            font-weight: 700;
            text-transform: uppercase;
        }}

        .fields-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            margin-top: 8px;
        }}

        .fields-table th, .fields-table td {{
            padding: 8px 6px;
            text-align: left;
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
        }}

        .fields-table th {{
            color: var(--text-secondary);
            font-weight: 600;
        }}

        .field-tag {{
            display: inline-block;
            padding: 1px 4px;
            border-radius: 3px;
            font-size: 9px;
            font-weight: 700;
            margin-right: 4px;
        }}

        .tag-pk {{ background: var(--accent-pk); color: #000; }}
        .tag-fk {{ background: var(--accent-fk); color: #000; }}
        .tag-auto {{ background: #c084fc; color: #000; }}

        .relation-link {{
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-color);
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 12px;
            cursor: pointer;
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 6px;
            transition: background 0.2s;
        }}

        .relation-link:hover {{
            background: rgba(255, 255, 255, 0.1);
        }}

        .sql-snippet {{
            background: #060911;
            padding: 10px;
            border-radius: 6px;
            font-family: monospace;
            font-size: 11px;
            color: #38bdf8;
            white-space: pre-wrap;
            word-break: break-all;
            max-height: 200px;
            overflow-y: auto;
        }}

        /* Legend */
        .legend {{
            position: absolute;
            bottom: 24px;
            right: 24px;
            background: var(--bg-glass);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 12px 16px;
            z-index: 9;
            display: flex;
            gap: 16px;
            font-size: 12px;
        }}

        .legend-item {{
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        .legend-circle {{
            width: 10px;
            height: 10px;
            border-radius: 50%;
        }}
    </style>
</head>
<body>
    <!-- Top Header Bar -->
    <div class="header-bar">
        <div class="brand-section">
            <div class="brand-badge">Schema Graph</div>
            <div class="brand-title">{db_filename}</div>
        </div>
        <div class="stats-group">
            <div class="stat-pill"><div class="stat-dot" style="background: var(--accent-table);"></div><span id="stat-tables">0</span> Tables</div>
            <div class="stat-pill"><div class="stat-dot" style="background: var(--accent-fk);"></div><span id="stat-relations">0</span> Relations</div>
            <div class="stat-pill"><div class="stat-dot" style="background: var(--accent-query);"></div><span id="stat-queries">0</span> Queries</div>
            <div class="stat-pill"><div class="stat-dot" style="background: var(--accent-form);"></div><span id="stat-forms">0</span> Forms</div>
            <div class="stat-pill"><div class="stat-dot" style="background: var(--accent-report);"></div><span id="stat-reports">0</span> Reports</div>
        </div>
    </div>

    <!-- Controls Panel -->
    <div class="controls-panel">
        <input type="text" class="search-box" id="search-input" placeholder="Search tables, fields, queries...">
        <div class="filter-row">
            <div class="filter-btn active" data-type="table">Tables</div>
            <div class="filter-btn active" data-type="query">Queries</div>
            <div class="filter-btn active" data-type="form">Forms</div>
            <div class="filter-btn active" data-type="report">Reports</div>
            <div class="filter-btn active" data-type="module">Modules</div>
        </div>
        <div class="action-buttons">
            <div class="btn-action" id="btn-fit">Fit View</div>
            <div class="btn-action" id="btn-physics">Pause Physics</div>
        </div>
    </div>

    <!-- Details Inspector Drawer -->
    <div class="inspector-drawer" id="inspector">
        <div class="drawer-header">
            <div class="drawer-title" id="drawer-title">Table Details</div>
            <button class="btn-close" id="btn-close-drawer">&times;</button>
        </div>
        <div id="drawer-content"></div>
    </div>

    <!-- Legend -->
    <div class="legend">
        <div class="legend-item"><div class="legend-circle" style="background: var(--accent-table);"></div>Table</div>
        <div class="legend-item"><div class="legend-circle" style="background: var(--accent-query);"></div>Query</div>
        <div class="legend-item"><div class="legend-circle" style="background: var(--accent-form);"></div>Form</div>
        <div class="legend-item"><div class="legend-circle" style="background: var(--accent-report);"></div>Report</div>
        <div class="legend-item"><div class="legend-circle" style="background: var(--accent-module);"></div>Module</div>
    </div>

    <!-- Graph Canvas -->
    <canvas id="graph-canvas"></canvas>

    <script>
        const graphData = {json_data_str};

        // Populate header stats
        document.getElementById('stat-tables').innerText = graphData.metadata.total_tables || 0;
        document.getElementById('stat-relations').innerText = graphData.metadata.total_relationships || 0;
        document.getElementById('stat-queries').innerText = graphData.metadata.total_queries || 0;
        document.getElementById('stat-forms').innerText = graphData.metadata.total_forms || 0;
        document.getElementById('stat-reports').innerText = graphData.metadata.total_reports || 0;

        const canvas = document.getElementById('graph-canvas');
        const ctx = canvas.getContext('2d');

        let width = canvas.width = window.innerWidth;
        let height = canvas.height = window.innerHeight;

        window.addEventListener('resize', () => {{
            width = canvas.width = window.innerWidth;
            height = canvas.height = window.innerHeight;
        }});

        // Color mapper
        const TYPE_COLORS = {{
            'table': '#10b981',
            'query': '#3b82f6',
            'form': '#8b5cf6',
            'report': '#f59e0b',
            'module': '#ec4899'
        }};

        // Active filters
        const activeTypes = new Set(['table', 'query', 'form', 'report', 'module']);

        // Initialize node positions & velocities
        const nodes = graphData.nodes.map((n, i) => {{
            const angle = (i / graphData.nodes.length) * Math.PI * 2;
            const radius = 250 + (i % 3) * 80;
            return {{
                ...n,
                x: width / 2 + Math.cos(angle) * radius + (Math.random() - 0.5) * 50,
                y: height / 2 + Math.sin(angle) * radius + (Math.random() - 0.5) * 50,
                vx: 0,
                vy: 0,
                radius: n.type === 'table' ? Math.max(16, Math.min(32, 14 + (n.degree || 0) * 2)) : 14
            }};
        }});

        const nodeMap = new Map();
        nodes.forEach(n => nodeMap.set(n.id, n));

        const edges = graphData.edges.map(e => ({{
            ...e,
            sourceNode: nodeMap.get(e.source),
            targetNode: nodeMap.get(e.target)
        }})).filter(e => e.sourceNode && e.targetNode);

        // View Transform
        let transform = {{ x: 0, y: 0, scale: 1 }};
        let isDragging = false;
        let dragNode = null;
        let startPan = {{ x: 0, y: 0 }};
        let physicsEnabled = true;
        let selectedNode = null;
        let hoveredNode = null;
        let searchQuery = "";

        // Simulation parameters
        const REPULSION = 1200;
        const SPRING_LENGTH = 140;
        const SPRING_STRENGTH = 0.05;
        const DAMPING = 0.85;
        const CENTER_GRAVITY = 0.01;

        function simulate() {{
            if (!physicsEnabled) return;

            const visibleNodes = nodes.filter(n => activeTypes.has(n.type));

            // Repulsion between visible nodes
            for (let i = 0; i < visibleNodes.length; i++) {{
                const na = visibleNodes[i];
                for (let j = i + 1; j < visibleNodes.length; j++) {{
                    const nb = visibleNodes[j];
                    const dx = nb.x - na.x;
                    const dy = nb.y - na.y;
                    const dist = Math.sqrt(dx * dx + dy * dy) || 1;
                    if (dist < 400) {{
                        const force = (REPULSION / (dist * dist));
                        const fx = (dx / dist) * force;
                        const fy = (dy / dist) * force;
                        if (na !== dragNode) {{ na.vx -= fx; na.vy -= fy; }}
                        if (nb !== dragNode) {{ nb.vx += fx; nb.vy += fy; }}
                    }}
                }}
            }}

            // Spring attraction along edges
            for (const edge of edges) {{
                if (!activeTypes.has(edge.sourceNode.type) || !activeTypes.has(edge.targetNode.type)) continue;
                const na = edge.sourceNode;
                const nb = edge.targetNode;
                const dx = nb.x - na.x;
                const dy = nb.y - na.y;
                const dist = Math.sqrt(dx * dx + dy * dy) || 1;
                const force = (dist - SPRING_LENGTH) * SPRING_STRENGTH;
                const fx = (dx / dist) * force;
                const fy = (dy / dist) * force;

                if (na !== dragNode) {{ na.vx += fx; na.vy += fy; }}
                if (nb !== dragNode) {{ nb.vx -= fx; nb.vy -= fy; }}
            }}

            // Gentle center gravity & update positions
            for (const n of visibleNodes) {{
                if (n === dragNode) continue;
                n.vx += (width / 2 - n.x) * CENTER_GRAVITY;
                n.vy += (height / 2 - n.y) * CENTER_GRAVITY;

                n.vx *= DAMPING;
                n.vy *= DAMPING;

                n.x += n.vx;
                n.y += n.vy;
            }}
        }}

        function draw() {{
            ctx.clearRect(0, 0, width, height);

            ctx.save();
            ctx.translate(transform.x, transform.y);
            ctx.scale(transform.scale, transform.scale);

            // Draw grid dots in background
            const gridSize = 40;
            ctx.fillStyle = 'rgba(255, 255, 255, 0.03)';
            const startX = -transform.x / transform.scale - 200;
            const endX = startX + width / transform.scale + 400;
            const startY = -transform.y / transform.scale - 200;
            const endY = startY + height / transform.scale + 400;

            for (let x = Math.floor(startX / gridSize) * gridSize; x < endX; x += gridSize) {{
                for (let y = Math.floor(startY / gridSize) * gridSize; y < endY; y += gridSize) {{
                    ctx.fillRect(x - 1, y - 1, 2, 2);
                }}
            }}

            // Draw Edges
            for (const edge of edges) {{
                if (!activeTypes.has(edge.sourceNode.type) || !activeTypes.has(edge.targetNode.type)) continue;

                const isConnectedToSelected = selectedNode && (edge.sourceNode === selectedNode || edge.targetNode === selectedNode);
                const isHovered = hoveredNode && (edge.sourceNode === hoveredNode || edge.targetNode === hoveredNode);

                ctx.beginPath();
                ctx.moveTo(edge.sourceNode.x, edge.sourceNode.y);
                ctx.lineTo(edge.targetNode.x, edge.targetNode.y);

                if (isConnectedToSelected || isHovered) {{
                    ctx.strokeStyle = '#38bdf8';
                    ctx.lineWidth = 3;
                    ctx.shadowColor = '#0284c7';
                    ctx.shadowBlur = 8;
                }} else {{
                    ctx.strokeStyle = edge.type === 'FOREIGN_KEY' ? 'rgba(16, 185, 129, 0.35)' : 'rgba(255, 255, 255, 0.15)';
                    ctx.lineWidth = edge.type === 'FOREIGN_KEY' ? 2 : 1;
                    ctx.shadowBlur = 0;
                }}
                ctx.stroke();
                ctx.shadowBlur = 0;

                // Draw edge label if selected or hovered
                if ((isConnectedToSelected || isHovered) && edge.label) {{
                    const midX = (edge.sourceNode.x + edge.targetNode.x) / 2;
                    const midY = (edge.sourceNode.y + edge.targetNode.y) / 2;
                    ctx.fillStyle = '#94a3b8';
                    ctx.font = '11px sans-serif';
                    ctx.textAlign = 'center';
                    ctx.fillText(edge.label, midX, midY - 6);
                }}
            }}

            // Draw Nodes
            for (const node of nodes) {{
                if (!activeTypes.has(node.type)) continue;

                const isSelected = selectedNode === node;
                const isHovered = hoveredNode === node;
                const matchesSearch = searchQuery && node.label.toLowerCase().includes(searchQuery);

                const color = TYPE_COLORS[node.type] || '#94a3b8';

                // Outer glow if active
                if (isSelected || isHovered || matchesSearch) {{
                    ctx.beginPath();
                    ctx.arc(node.x, node.y, node.radius + 8, 0, Math.PI * 2);
                    ctx.fillStyle = isSelected ? 'rgba(59, 130, 246, 0.4)' : (matchesSearch ? 'rgba(251, 191, 36, 0.4)' : 'rgba(255, 255, 255, 0.15)');
                    ctx.fill();
                }}

                // Main node circle
                ctx.beginPath();
                ctx.arc(node.x, node.y, node.radius, 0, Math.PI * 2);
                ctx.fillStyle = color;
                ctx.fill();
                ctx.strokeStyle = isSelected ? '#ffffff' : 'rgba(255, 255, 255, 0.3)';
                ctx.lineWidth = isSelected ? 3 : 1.5;
                ctx.stroke();

                // Inner icon / letter
                ctx.fillStyle = '#0f172a';
                ctx.font = `bold ${{Math.max(10, node.radius * 0.8)}}px sans-serif`;
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                const initial = node.type === 'table' ? 'T' : (node.type === 'query' ? 'Q' : (node.type === 'form' ? 'F' : (node.type === 'report' ? 'R' : 'M')));
                ctx.fillText(initial, node.x, node.y);

                // Node Label below
                ctx.font = isSelected ? 'bold 13px sans-serif' : '12px sans-serif';
                ctx.fillStyle = isSelected ? '#ffffff' : (matchesSearch ? '#fbbf24' : '#e2e8f0');
                ctx.textAlign = 'center';
                ctx.textBaseline = 'top';
                ctx.fillText(node.label, node.x, node.y + node.radius + 6);
            }}

            ctx.restore();
        }}

        function loop() {{
            simulate();
            draw();
            requestAnimationFrame(loop);
        }}
        loop();

        // Screen coordinate to world coordinate
        function screenToWorld(sx, sy) {{
            return {{
                x: (sx - transform.x) / transform.scale,
                y: (sy - transform.y) / transform.scale
            }};
        }}

        // Mouse Events for Pan, Zoom, Drag, Inspect
        canvas.addEventListener('mousedown', e => {{
            const w = screenToWorld(e.clientX, e.clientY);
            // Check if clicked a node
            const clicked = nodes.slice().reverse().find(n => {{
                if (!activeTypes.has(n.type)) return false;
                const dx = n.x - w.x;
                const dy = n.y - w.y;
                return Math.sqrt(dx * dx + dy * dy) <= n.radius + 5;
            }});

            if (clicked) {{
                dragNode = clicked;
                selectNode(clicked);
            }} else {{
                isDragging = true;
                startPan = {{ x: e.clientX - transform.x, y: e.clientY - transform.y }};
            }}
        }});

        window.addEventListener('mousemove', e => {{
            const w = screenToWorld(e.clientX, e.clientY);

            if (dragNode) {{
                dragNode.x = w.x;
                dragNode.y = w.y;
                dragNode.vx = 0;
                dragNode.vy = 0;
            }} else if (isDragging) {{
                transform.x = e.clientX - startPan.x;
                transform.y = e.clientY - startPan.y;
            }} else {{
                // Hover detection
                const hovered = nodes.slice().reverse().find(n => {{
                    if (!activeTypes.has(n.type)) return false;
                    const dx = n.x - w.x;
                    const dy = n.y - w.y;
                    return Math.sqrt(dx * dx + dy * dy) <= n.radius + 5;
                }});
                hoveredNode = hovered || null;
            }}
        }});

        window.addEventListener('mouseup', () => {{
            dragNode = null;
            isDragging = false;
        }});

        // Zoom with wheel
        canvas.addEventListener('wheel', e => {{
            e.preventDefault();
            const zoomFactor = e.deltaY < 0 ? 1.1 : 0.9;
            const mouseWorld = screenToWorld(e.clientX, e.clientY);

            transform.scale *= zoomFactor;
            transform.scale = Math.max(0.2, Math.min(4, transform.scale));

            transform.x = e.clientX - mouseWorld.x * transform.scale;
            transform.y = e.clientY - mouseWorld.y * transform.scale;
        }});

        // Inspector drawer logic
        const inspector = document.getElementById('inspector');
        const drawerTitle = document.getElementById('drawer-title');
        const drawerContent = document.getElementById('drawer-content');
        const btnClose = document.getElementById('btn-close-drawer');

        btnClose.addEventListener('click', () => {{
            inspector.classList.remove('open');
            selectedNode = null;
        }});

        function selectNode(node) {{
            selectedNode = node;
            inspector.classList.add('open');
            const color = TYPE_COLORS[node.type] || '#94a3b8';

            drawerTitle.innerHTML = `<span style="color: ${{color}};">●</span> ${{node.label}} <span class="badge-pill" style="background: ${{color}}22; color: ${{color}};">${{node.type}}</span>`;

            let html = "";

            if (node.type === 'table') {{
                html += `<div><strong>Columns (${{node.fields ? node.fields.length : 0}}):</strong></div>`;
                html += `<table class="fields-table"><thead><tr><th>Name</th><th>Type</th><th>Key</th></tr></thead><tbody>`;

                if (node.fields) {{
                    for (const f of node.fields) {{
                        let badges = "";
                        if (f.is_pk) badges += `<span class="field-tag tag-pk">PK</span>`;
                        if (f.is_autonumber) badges += `<span class="field-tag tag-auto">Auto</span>`;
                        html += `<tr><td>${{f.name}}</td><td style="color: var(--text-secondary);">${{f.type}}</td><td>${{badges}}</td></tr>`;
                    }}
                }}
                html += `</tbody></table>`;

                // Show related tables
                const connectedEdges = edges.filter(e => e.sourceNode === node || e.targetNode === node);
                if (connectedEdges.length > 0) {{
                    html += `<div style="margin-top: 16px;"><strong>Connected Relationships:</strong></div><div style="margin-top: 8px;">`;
                    for (const ce of connectedEdges) {{
                        const otherNode = ce.sourceNode === node ? ce.targetNode : ce.sourceNode;
                        const direction = ce.sourceNode === node ? "--> (Child)" : "<-- (Parent)";
                        html += `<div class="relation-link" onclick="focusNode('${{otherNode.id}}')">
                            <div><strong>${{otherNode.label}}</strong> <span style="font-size: 10px; color: var(--text-secondary);">${{direction}}</span></div>
                            <div style="font-size: 11px; color: #38bdf8;">${{ce.label || ''}}</div>
                        </div>`;
                    }}
                    html += `</div>`;
                }}
            }} else if (node.type === 'query') {{
                html += `<div><strong>Query SQL:</strong></div>`;
                html += `<div class="sql-snippet">${{node.sql || 'No SQL available'}}</div>`;
                if (node.referenced_tables && node.referenced_tables.length > 0) {{
                    html += `<div style="margin-top: 16px;"><strong>Source Tables:</strong></div><div style="margin-top: 8px;">`;
                    for (const t of node.referenced_tables) {{
                        html += `<div class="relation-link" onclick="focusNode('table:${{t}}')">
                            <div>Table: <strong>${{t}}</strong></div>
                            <div style="font-size: 11px; color: #10b981;">Inspect</div>
                        </div>`;
                    }}
                    html += `</div>`;
                }}
            }} else {{
                html += `<div style="color: var(--text-secondary);">Object Name: <strong>${{node.object_name}}</strong></div>`;
                if (node.module_type) {{
                    html += `<div style="margin-top: 8px;">Type: ${{node.module_type}}</div>`;
                    html += `<div>Lines of code: ${{node.line_count || 0}}</div>`;
                }}
            }}

            drawerContent.innerHTML = html;
        }}

        window.focusNode = function(nodeId) {{
            const target = nodeMap.get(nodeId);
            if (target) {{
                selectNode(target);
                transform.x = width / 2 - target.x * transform.scale;
                transform.y = height / 2 - target.y * transform.scale;
            }}
        }};

        // Fit View
        document.getElementById('btn-fit').addEventListener('click', () => {{
            const visibleNodes = nodes.filter(n => activeTypes.has(n.type));
            if (visibleNodes.length === 0) return;

            let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
            for (const n of visibleNodes) {{
                minX = Math.min(minX, n.x);
                maxX = Math.max(maxX, n.x);
                minY = Math.min(minY, n.y);
                maxY = Math.max(maxY, n.y);
            }}

            const boxWidth = maxX - minX + 100;
            const boxHeight = maxY - minY + 100;
            const scaleX = (width - 100) / boxWidth;
            const scaleY = (height - 100) / boxHeight;
            transform.scale = Math.min(1.5, Math.max(0.3, Math.min(scaleX, scaleY)));

            transform.x = width / 2 - ((minX + maxX) / 2) * transform.scale;
            transform.y = height / 2 - ((minY + maxY) / 2) * transform.scale;
        }});

        // Toggle Physics
        const btnPhysics = document.getElementById('btn-physics');
        btnPhysics.addEventListener('click', () => {{
            physicsEnabled = !physicsEnabled;
            btnPhysics.innerText = physicsEnabled ? "Pause Physics" : "Resume Physics";
        }});

        // Filter toggles
        document.querySelectorAll('.filter-btn').forEach(btn => {{
            btn.addEventListener('click', () => {{
                const type = btn.dataset.type;
                if (activeTypes.has(type)) {{
                    activeTypes.delete(type);
                    btn.classList.remove('active');
                }} else {{
                    activeTypes.add(type);
                    btn.classList.add('active');
                }}
            }});
        }});

        // Search box input
        const searchInput = document.getElementById('search-input');
        searchInput.addEventListener('input', e => {{
            searchQuery = e.target.value.trim().toLowerCase();
            if (searchQuery) {{
                const match = nodes.find(n => activeTypes.has(n.type) && n.label.toLowerCase().includes(searchQuery));
                if (match) {{
                    transform.x = width / 2 - match.x * transform.scale;
                    transform.y = height / 2 - match.y * transform.scale;
                }}
            }}
        }});
    </script>
</body>
</html>"""

        # Write output file
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        with open(output_file, "w", encoding="utf-8") as f:
            f.write(html_template)

        return os.path.abspath(output_file)
