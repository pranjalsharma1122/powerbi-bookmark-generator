import json
import os
import re
import time
import uuid
from typing import Dict, List, Optional, Tuple

try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False
    print("⚠️ google-generativeai not installed. Install: pip install google-generativeai")
    exit(1)

# ============================================================================
# CONFIGURATION
# ============================================================================
BASE_DIR = r"C:\Bookmark\output"
REFERENCE_FILE = r"C:\Bookmark\output\Reference Chart Configurations.txt"
OUTPUT_FILE = os.path.join(BASE_DIR, "visuals_output.json")

GEMINI_API_KEY = "AIzaSyD3I9LHDjDc6quM67o1xdFiZgHZi9qEtj4"
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-2.5-flash")

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def gen_guid() -> str:
    """Generate 20-character hex GUID for Power BI visual IDs"""
    return uuid.uuid4().hex[:20]


def load_json(filepath: str) -> dict:
    """Load JSON file with error handling"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"❌ File not found: {filepath}")
        return {}
    except json.JSONDecodeError as e:
        print(f"❌ Invalid JSON in {filepath}: {e}")
        return {}


def get_column_type(table: str, column: str, schema: dict) -> str:
    """Get column data type from schema_output.json"""
    if table not in schema:
        return "string"
    for col in schema[table]:
        if col.get("name") == column:
            return col.get("type", "string")
    return "string"


def get_agg_function_code(func_name: str) -> int:
    """
    Map aggregation function name to Power BI code
    0=Sum, 1=Average, 2=Min, 3=Max, 4=Count, 5=DistinctCount
    """
    mapping = {
        "sum": 0, "average": 1, "avg": 1, "mean": 1,
        "min": 2, "max": 3, "count": 4, 
        "distinctcount": 5, "countd": 5
    }
    return mapping.get(func_name.lower(), 0)


def map_chart_type(tableau_type: str, visual_data: dict) -> str:
    """
    🔥 FIXED: Map Tableau chart type to Power BI visualType
    
    Key Improvements:
    - Checks visual title for hints (e.g., "Donut M" → donutChart)
    - Checks labels flag (donuts often have labels=true)
    - Checks if Legend exists (pies/donuts typically have legends)
    """
    if not tableau_type:
        return "clusteredColumnChart"
    
    # Check title for explicit hints
    title = (visual_data.get("title") or visual_data.get("Source") or "").lower()
    
    # Explicit donut/pie detection
    if "donut" in title:
        return "donutChart"
    if "pie" in title:
        return "pieChart"
    
    # Check chart_type field
    chart_lower = tableau_type.lower()
    
    # Direct mappings
    mapping = {
        "donutchart": "donutChart",
        "donut": "donutChart",
        "piechart": "pieChart",
        "pie": "pieChart",
        "linechart": "lineChart",
        "line": "lineChart",
        "bar": "clusteredBarChart",
        "clusteredbarchart": "clusteredBarChart",
        "column": "clusteredColumnChart",
        "columnchart": "clusteredColumnChart",
        "clusteredcolumnchart": "clusteredColumnChart",
        "histogram": "clusteredColumnChart",
        "area": "areaChart",
        "scatter": "scatterChart",
        "map": "shapeMap",
        "shapemap": "shapeMap"
    }
    
    matched_type = mapping.get(chart_lower, None)
    if matched_type:
        return matched_type
    
    # Heuristic: If has Legend + labels=true + no rows → likely donut/pie
    has_legend = bool(visual_data.get("Legend"))
    has_labels = visual_data.get("labels", False)
    rows = visual_data.get("Rows", {})
    
    if has_legend and has_labels and not rows:
        return "donutChart"
    
    # Default
    return "clusteredColumnChart"


def parse_aggregation(agg_str: str) -> Tuple[str, Optional[str], str]:
    """
    Parse aggregation string like "Sum(Table.Column)" or "Average(Column)"
    Returns: (function, table, column)
    """
    agg_str = (agg_str or "").strip()
    
    # Match "Func(Table.Col)" or "Func(Col)"
    match = re.search(r'([A-Za-z]+)\s*\(\s*(?:([A-Za-z0-9_\s]+)\.)?([^)]+)\s*\)', agg_str)
    if match:
        func, table, col = match.groups()
        return func.strip(), table.strip() if table else None, col.strip()
    
    return "Sum", None, agg_str


def gemini_safe_generate(prompt: str, max_retries: int = 3) -> str:
    """Call Gemini API with retry logic"""
    for attempt in range(max_retries):
        try:
            response = model.generate_content(prompt)
            if response and response.candidates:
                candidate = response.candidates[0]
                if candidate.content and hasattr(candidate.content, "parts"):
                    text = "".join(part.text for part in candidate.content.parts 
                                   if hasattr(part, "text"))
                    return text
        except Exception as e:
            print(f"  ⚠️ Gemini attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                time.sleep(2)
    return ""


def validate_config_with_gemini(config_str: str, visual_title: str) -> str:
    """
    Validate Power BI config with Gemini (3 retries)
    Returns corrected config or original if validation passes
    """
    for attempt in range(3):
        prompt = f"""Validate this Power BI visual config:

```json
{config_str}
```

Rules:
1. Keep ALL table/column names EXACTLY as written
2. Verify JSON syntax is valid
3. Check required keys: name, layouts, singleVisual
4. Verify projections have queryRef fields
5. Check objects structure if present

If valid: respond "VALID"
If invalid: provide corrected JSON in ```json``` code block"""
        
        response = gemini_safe_generate(prompt)
        
        if "VALID" in response and "corrected" not in response.lower():
            return config_str
        
        # Extract corrected JSON
        match = re.search(r'```json\s*(.*?)\s*```', response, re.DOTALL)
        if match:
            corrected = match.group(1).strip()
            try:
                json.loads(corrected)  # Verify it parses
                print(f"  ✅ Gemini corrected '{visual_title}' (attempt {attempt + 1})")
                return corrected
            except:
                config_str = corrected  # Try again with corrected version
        
        time.sleep(2)
    
    print(f"  ⚠️ Using unvalidated config for '{visual_title}'")
    return config_str


# ============================================================================
# VISUAL BUILDING FUNCTIONS
# ============================================================================

def build_projections(tableau_visual: dict, schema: dict) -> dict:
    """Build Power BI projections from Tableau visual data"""
    projections = {}
    
    columns = tableau_visual.get("Columns", {}) or {}
    rows = tableau_visual.get("Rows", {}) or {}
    agg_row = tableau_visual.get("Aggregation_row", [])
    agg_col = tableau_visual.get("Aggregation_columns", [])
    hierarchy = tableau_visual.get("Hierarchy", [])
    
    # Normalize to lists
    if not isinstance(agg_row, list):
        agg_row = [agg_row] if agg_row else []
    if not isinstance(agg_col, list):
        agg_col = [agg_col] if agg_col else []
    
    # Build list of aggregated field names to exclude from Category
    aggregated_fields = set()
    for agg_str in (agg_row + agg_col):
        func, table, col = parse_aggregation(agg_str)
        aggregated_fields.add(col.lower())
    
    # Build Category (dimensions only - exclude aggregated measures)
    category = []
    for col_name, table in columns.items():
        # Skip if this column is used in aggregation
        if col_name.lower() in aggregated_fields:
            continue
            
        col_type = get_column_type(table, col_name, schema)
        
        # Check if it's a date with hierarchy
        if col_type == "datetime" and hierarchy:
            for hier_str in hierarchy:
                if "Year" in hier_str:
                    category.append({
                        "queryRef": f"{table}.{col_name}.Variation.Date Hierarchy.Year",
                        "active": True
                    })
                if "Month" in hier_str:
                    category.append({
                        "queryRef": f"{table}.{col_name}.Variation.Date Hierarchy.Month",
                        "active": True
                    })
                if "Day" in hier_str:
                    category.append({
                        "queryRef": f"{table}.{col_name}.Variation.Date Hierarchy.Day",
                        "active": True
                    })
        else:
            category.append({
                "queryRef": f"{table}.{col_name}",
                "active": True
            })
    
    # Also check rows for dimensions
    for row_name, table in rows.items():
        if row_name.lower() in aggregated_fields:
            continue
        if any(row_name in c.get("queryRef", "") for c in category):
            continue
            
        category.append({
            "queryRef": f"{table}.{row_name}",
            "active": True
        })
    
    if category:
        projections["Category"] = category
    
    # Build Y (measures)
    y_items = []
    for agg_str in (agg_row + agg_col):
        func, table, col = parse_aggregation(agg_str)
        
        if not table:
            table = columns.get(col) or rows.get(col)
            if not table:
                all_tables = tableau_visual.get("tables", [])
                if all_tables:
                    table = all_tables[0]
                elif columns:
                    table = list(columns.values())[0]
                elif rows:
                    table = list(rows.values())[0]
                else:
                    print(f"  ⚠️ Cannot determine table for column '{col}', skipping")
                    continue
        
        y_items.append({"queryRef": f"{func}({table}.{col})"})
    
    if y_items:
        projections["Y"] = y_items
    
    # Add Series if Legend exists
    legend = tableau_visual.get("Legend")
    if isinstance(legend, dict) and legend:
        for legend_field, legend_table in legend.items():
            legend_ref = f"{legend_table}.{legend_field}"
            if not any(legend_field in str(y) for y in y_items):
                projections["Series"] = [{"queryRef": legend_ref}]
                break
    
    return projections


def build_prototype_query(tableau_visual: dict, schema: dict) -> dict:
    """Build prototypeQuery matching Power BI structure"""
    query = {"Version": 2, "From": [], "Select": []}
    
    tables_used = set()
    columns = tableau_visual.get("Columns", {}) or {}
    rows = tableau_visual.get("Rows", {}) or {}
    
    for table in list(columns.values()) + list(rows.values()):
        if table and table not in tables_used:
            alias = table[0].lower()
            query["From"].append({
                "Name": alias,
                "Entity": table,
                "Type": 0
            })
            tables_used.add(table)
    
    # Add dimension selects
    hierarchy = tableau_visual.get("Hierarchy", [])
    for col_name, table in columns.items():
        alias = table[0].lower()
        col_type = get_column_type(table, col_name, schema)
        
        if col_type == "datetime" and hierarchy:
            query["Select"].append({
                "HierarchyLevel": {
                    "Expression": {
                        "Hierarchy": {
                            "Expression": {
                                "PropertyVariationSource": {
                                    "Expression": {"SourceRef": {"Source": alias}},
                                    "Name": "Variation",
                                    "Property": col_name
                                }
                            },
                            "Hierarchy": "Date Hierarchy"
                        }
                    },
                    "Level": "Year"
                },
                "Name": f"{table}.{col_name}.Variation.Date Hierarchy.Year",
                "NativeReferenceName": f"{col_name} Year"
            })
            query["Select"].append({
                "HierarchyLevel": {
                    "Expression": {
                        "Hierarchy": {
                            "Expression": {
                                "PropertyVariationSource": {
                                    "Expression": {"SourceRef": {"Source": alias}},
                                    "Name": "Variation",
                                    "Property": col_name
                                }
                            },
                            "Hierarchy": "Date Hierarchy"
                        }
                    },
                    "Level": "Month"
                },
                "Name": f"{table}.{col_name}.Variation.Date Hierarchy.Month",
                "NativeReferenceName": f"{col_name} Month"
            })
        else:
            query["Select"].append({
                "Column": {
                    "Expression": {"SourceRef": {"Source": alias}},
                    "Property": col_name
                },
                "Name": f"{table}.{col_name}",
                "NativeReferenceName": col_name
            })
    
    # Add measure selects
    agg_row = tableau_visual.get("Aggregation_row", [])
    agg_col = tableau_visual.get("Aggregation_columns", [])
    
    if not isinstance(agg_row, list):
        agg_row = [agg_row] if agg_row else []
    if not isinstance(agg_col, list):
        agg_col = [agg_col] if agg_col else []
    
    for agg_str in (agg_row + agg_col):
        func, table, col = parse_aggregation(agg_str)
        
        if not table:
            table = columns.get(col) or rows.get(col)
            if not table and tables_used:
                table = list(tables_used)[0]
            elif not table:
                print(f"  ⚠️ Cannot determine table for measure '{col}', skipping")
                continue
        
        alias = table[0].lower()
        
        query["Select"].append({
            "Aggregation": {
                "Expression": {
                    "Column": {
                        "Expression": {"SourceRef": {"Source": alias}},
                        "Property": col
                    }
                },
                "Function": get_agg_function_code(func)
            },
            "Name": f"{func}({table}.{col})",
            "NativeReferenceName": f"{func} of {col}"
        })
    
    return query


def build_objects(tableau_visual: dict, colors_data: list) -> dict:
    """Build objects (conditional formatting, labels, etc.)"""
    objects = {}
    
    source = tableau_visual.get("Source")
    cf_config = next((c for c in colors_data if c.get("Source_name") == source), None)
    
    if cf_config and cf_config.get("mark"):
        marks = cf_config["mark"]
        
        if len(marks) >= 2:
            data_points = []
            for field_name, field_data in marks.items():
                palette = field_data.get("palette", {})
                color = None
                
                if isinstance(palette, dict):
                    color = palette.get('starting_value') or palette.get('middle_value') or palette.get('ending_value')
                elif isinstance(palette, list) and palette:
                    color = palette[0]
                
                if color:
                    rows = tableau_visual.get("Rows", {}) or {}
                    cols = tableau_visual.get("Columns", {}) or {}
                    table = rows.get(field_name) or cols.get(field_name)
                    
                    if table:
                        data_points.append({
                            "properties": {
                                "fill": {
                                    "solid": {
                                        "color": {"expr": {"Literal": {"Value": f"'{color}'"}}}
                                    }
                                }
                            },
                            "selector": {"metadata": f"{table}.{field_name}"}
                        })
            
            if data_points:
                objects["dataPoint"] = data_points
        
        elif len(marks) == 1:
            measure_name = list(marks.keys())[0]
            palette = marks[measure_name].get("palette", {})
            
            if isinstance(palette, dict):
                start = palette.get("starting_value")
                middle = palette.get("middle_value")
                end = palette.get("ending_value")
                
                if start and end:
                    rows = tableau_visual.get("Rows", {}) or {}
                    cols = tableau_visual.get("Columns", {}) or {}
                    table = rows.get(measure_name) or cols.get(measure_name) or "Sheet1"
                    
                    agg_row = tableau_visual.get("Aggregation_row", [])
                    if not isinstance(agg_row, list):
                        agg_row = [agg_row] if agg_row else []
                    
                    func_code = 0
                    for agg_str in agg_row:
                        if measure_name in agg_str:
                            func, _, _ = parse_aggregation(agg_str)
                            func_code = get_agg_function_code(func)
                            break
                    
                    if middle:
                        fill_rule = {
                            "linearGradient3": {
                                "min": {"color": {"Literal": {"Value": f"'{start}'"}}},
                                "mid": {"color": {"Literal": {"Value": f"'{middle}'"}}},
                                "max": {"color": {"Literal": {"Value": f"'{end}'"}}},
                                "nullColoringStrategy": {"strategy": {"Literal": {"Value": "'asZero'"}}}
                            }
                        }
                    else:
                        fill_rule = {
                            "linearGradient2": {
                                "min": {"color": {"Literal": {"Value": f"'{start}'"}}},
                                "max": {"color": {"Literal": {"Value": f"'{end}'"}}},
                                "nullColoringStrategy": {"strategy": {"Literal": {"Value": "'asZero'"}}}
                            }
                        }
                    
                    objects["dataPoint"] = [{
                        "properties": {
                            "fill": {
                                "solid": {
                                    "color": {
                                        "expr": {
                                            "FillRule": {
                                                "Input": {
                                                    "Aggregation": {
                                                        "Expression": {
                                                            "Column": {
                                                                "Expression": {"SourceRef": {"Entity": table}},
                                                                "Property": measure_name
                                                            }
                                                        },
                                                        "Function": func_code
                                                    }
                                                },
                                                "FillRule": fill_rule
                                            }
                                        }
                                    }
                                }
                            }
                        },
                        "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}
                    }]
    
    show_labels = tableau_visual.get("labels", False)
    objects["labels"] = [{
        "properties": {
            "show": {"expr": {"Literal": {"Value": "true" if show_labels else "false"}}}
        }
    }]
    
    return objects


# ============================================================================
# MAIN PROCESSING
# ============================================================================

def generate_visuals():
    """Main function to generate Power BI visual configurations"""
    print("\n" + "="*70)
    print("STAGE 1: Generating Power BI Visual Configurations")
    print("="*70 + "\n")
    
    print("📂 Loading input files...")
    final_data = load_json(os.path.join(BASE_DIR, "final.json"))
    schema_data = load_json(os.path.join(BASE_DIR, "schema_output.json"))
    colors_data = load_json(os.path.join(BASE_DIR, "extracted_colors.json"))
    positions_data = load_json(os.path.join(BASE_DIR, "powerbi_chart_positions.json"))
    
    if not final_data:
        print("❌ Cannot proceed without final.json")
        return
    
    print(f"✅ Loaded {len(final_data)} visuals from final.json")
    print(f"✅ Loaded schema for {len(schema_data)} tables")
    print(f"✅ Loaded {len(colors_data)} color configs")
    print(f"✅ Loaded {len(positions_data)} position configs\n")
    
    generated_visuals = []
    
    for i, tableau_visual in enumerate(final_data):
        source = tableau_visual.get("Source", f"Visual_{i}")
        title = tableau_visual.get("title") or source
        chart_type = tableau_visual.get("chart_type", "")
        
        if source == "Buttons":
            print(f"⏭️ Skipping '{source}' (metadata only)")
            continue
        
        if chart_type in ["treemap", "cardVisual", "symbolMap", "bulletChart"]:
            print(f"⏭️ Skipping '{source}' ({chart_type} not supported)")
            continue
        
        print(f"\n🛠️ Processing Visual {i + 1}: {title}")
        
        visual_guid = gen_guid()
        
        pos = next((p for p in positions_data if p.get("chart") == source), None)
        if not pos:
            print(f"   ⚠️ No position found, using default")
            pos = {"x": 0, "y": 200 + i*100, "width": 640, "height": 400, "z": i}
        
        # 🔥 FIXED: Use improved chart type detection
        detected_type = map_chart_type(chart_type, tableau_visual)
        print(f"   Type: {chart_type or 'auto'} → {detected_type}")
        
        projections = build_projections(tableau_visual, schema_data)
        prototype_query = build_prototype_query(tableau_visual, schema_data)
        objects = build_objects(tableau_visual, colors_data)
        
        single_visual = {
            "visualType": detected_type,
            "projections": projections,
            "prototypeQuery": prototype_query,
            "drillFilterOtherVisuals": True
        }
        
        if objects:
            single_visual["objects"] = objects
        
        single_visual["vcObjects"] = {
            "title": [{
                "properties": {
                    "text": {"expr": {"Literal": {"Value": f"'{title}'"}}}
                }
            }]
        }
        
        visual_config = {
            "name": visual_guid,
            "layouts": [{
                "id": 0,
                "position": {
                    "x": pos["x"],
                    "y": pos["y"],
                    "z": pos.get("z", i),
                    "width": pos["width"],
                    "height": pos["height"],
                    "tabOrder": 0
                }
            }],
            "singleVisual": single_visual
        }
        
        config_str = json.dumps(visual_config, separators=(',', ':'))
        
        print(f"   🤖 Validating with Gemini...")
        validated_config = validate_config_with_gemini(config_str, title)
        
        visual_container = {
            "name": visual_guid,
            "title": title,
            "source": source,
            "config": validated_config,
            "filters": "[]",
            "height": pos["height"],
            "width": pos["width"],
            "x": pos["x"],
            "y": pos["y"],
            "z": pos.get("z", i)
        }
        
        generated_visuals.append(visual_container)
        print(f"   ✅ Visual generated successfully")
    
    print(f"\n💾 Saving {len(generated_visuals)} visuals to {OUTPUT_FILE}...")
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(generated_visuals, f, indent=2, ensure_ascii=False)
    
    print("\n" + "="*70)
    print(f"✅ STAGE 1 COMPLETE: Generated {len(generated_visuals)} visuals")
    print(f"📄 Output: {OUTPUT_FILE}")
    print("="*70 + "\n")


if __name__ == "__main__":
    generate_visuals()