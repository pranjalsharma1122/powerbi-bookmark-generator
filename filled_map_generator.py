import json
import os
import re
import time
from typing import Dict, List, Optional

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
BASE_DIR = r"C:\map\output"
REFERENCE_FILE = os.path.join(BASE_DIR, "Reference Chart Configurations.txt")
OUTPUT_FILE = os.path.join(BASE_DIR, "visual_output.json")

GEMINI_API_KEY = "AIzaSyD3I9LHDjDc6quM67o1xdFiZgHZi9qEtj4"
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-2.5-flash")

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

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


def extract_reference_structure(reference_file: str) -> Optional[Dict]:
    """Extract the complete reference structure including config, filters, positioning"""
    try:
        with open(reference_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Parse the entire reference file as JSON
        reference_data = json.loads(content)
        
        # Extract and parse the nested config string
        config_str = reference_data.get("config", "")
        config_str = config_str.replace('\\"', '"')
        config = json.loads(config_str)
        
        print("✅ Extracted complete reference structure")
        print(f"   Visual Type: {config.get('singleVisual', {}).get('visualType')}")
        
        return {
            "config": config,
            "filters": reference_data.get("filters", "[]"),
            "height": reference_data.get("height"),
            "width": reference_data.get("width"),
            "x": reference_data.get("x"),
            "y": reference_data.get("y"),
            "z": reference_data.get("z")
        }
        
    except Exception as e:
        print(f"❌ Error extracting reference structure: {e}")
        return None


def get_column_info(table: str, column: str, schema: dict) -> Dict:
    """Get column information from schema"""
    if table not in schema:
        return {"type": "string", "is_categorical": True, "is_numeric": False, "is_integer": False}
    
    for col in schema[table]:
        if col.get("name") == column:
            col_type = col.get("type", "string").lower()
            is_numeric = col_type in ['number', 'integer', 'float', 'datetime']
            is_integer = col_type in ['integer', 'number']
            is_categorical = col_type in ['string', 'text']
            return {
                "type": col_type,
                "is_categorical": is_categorical,
                "is_numeric": is_numeric,
                "is_integer": is_integer
            }
    return {"type": "string", "is_categorical": True, "is_numeric": False, "is_integer": False}


def classify_columns_for_gemini(visual_data: dict, schema: dict) -> Dict:
    """Classify columns and prepare data for Gemini"""
    columns = visual_data.get("Columns", {})
    legend = visual_data.get("Legend", {})
    tooltips = visual_data.get("tooltips", {})
    agg_columns = visual_data.get("Aggregation_columns", [])
    tooltips_agg = visual_data.get("tooltips_agg", [])
    
    # Build column classification
    categorical_fields = []
    numeric_fields = []
    
    for col_name, table_name in columns.items():
        col_info = get_column_info(table_name, col_name, schema)
        
        field_data = {
            "name": col_name,
            "table": table_name,
            "type": col_info["type"],
            "queryRef": f"{table_name}.{col_name}"
        }
        
        if col_info["is_categorical"]:
            categorical_fields.append(field_data)
        elif col_info["is_numeric"]:
            numeric_fields.append(field_data)
    
    # Build tooltips with aggregations
    tooltip_fields = []
    for tooltip_field, tooltip_table in tooltips.items():
        agg = next((a for a in tooltips_agg if tooltip_field.lower() in a.lower()), None)
        tooltip_fields.append({
            "name": tooltip_field,
            "table": tooltip_table,
            "aggregation": agg,
            "queryRef": f"{tooltip_table}.{tooltip_field}"
        })
    
    return {
        "categorical": categorical_fields,
        "numeric": numeric_fields,
        "legend": legend,
        "tooltips": tooltip_fields,
        "aggregation_columns": agg_columns
    }


def build_dynamic_config(visual_data: dict, schema: dict, reference: Dict) -> Dict:
    """Build configuration dynamically following reference structure"""

    classified = classify_columns_for_gemini(visual_data, schema)
    ref_config = reference.get("config", {})

    # Get fields from final.json
    columns = visual_data.get("Columns", {})
    legend = visual_data.get("Legend", {})
    tooltips = visual_data.get("tooltips", {})
    tooltips_agg = visual_data.get("tooltips_agg", [])

    # Determine Category field (first categorical)
    category_field = None
    for col_name, table_name in columns.items():
        col_info = get_column_info(table_name, col_name, schema)
        if col_info.get("is_categorical"):
            category_field = {"name": col_name, "table": table_name}
            break

    # Determine Y (first legend field)
    y_field = None
    if legend:
        first_legend_name = list(legend.keys())[0]
        first_legend_table = legend[first_legend_name]
        y_field = {"name": first_legend_name, "table": first_legend_table}

    # Build projections
    projections = {
        "Category": [],
        "Y": [],
        "Tooltips": []
    }

    if category_field:
        projections["Category"].append({
            "queryRef": f"{category_field['table']}.{category_field['name']}",
            "active": True
        })

    if y_field:
        projections["Y"].append({
            "queryRef": f"{y_field['table']}.{y_field['name']}"
        })

    # Add tooltip fields
    for tooltip_name, tooltip_table in tooltips.items():
        projections["Tooltips"].append({
            "queryRef": f"{tooltip_table}.{tooltip_name}"
        })

    # Build prototypeQuery parts
    all_tables = set()
    select_fields = []
    column_properties = {}
    added_fields = set()

    # Collect all tables used
    if category_field:
        all_tables.add(category_field["table"])
    if y_field:
        all_tables.add(y_field["table"])
    for tooltip_table in tooltips.values():
        all_tables.add(tooltip_table)

    # Table aliases
    table_aliases = {}
    alias_counter = 0
    alias_names = ["bp", "b", "c", "d", "e"]

    for table in sorted(all_tables):
        table_aliases[table] = alias_names[alias_counter] if alias_counter < len(alias_names) else f"t{alias_counter}"
        alias_counter += 1

    # FROM clause
    from_clause = []
    for table, alias in table_aliases.items():
        from_clause.append({
            "Name": alias,
            "Entity": table,
            "Type": 0
        })

    # Helper for aggregation mapping
    def find_aggregation(field_name, agg_list):
        for agg in agg_list:
            if field_name.lower() in agg.lower():
                return agg
        return None

    # SELECT: Category
    if category_field:
        field_key = f"{category_field['table']}.{category_field['name']}"
        select_fields.append({
            "Column": {
                "Expression": {
                    "SourceRef": {
                        "Source": table_aliases[category_field["table"]]
                    }
                },
                "Property": category_field["name"]
            },
            "Name": field_key
        })
        added_fields.add(field_key)

    # SELECT: Y field
    if y_field:
        field_key = f"{y_field['table']}.{y_field['name']}"
        agg = find_aggregation(y_field["name"], tooltips_agg)

        select_fields.append({
            "Measure": {
                "Expression": {
                    "SourceRef": {
                        "Source": table_aliases[y_field["table"]]
                    }
                },
                "Property": y_field["name"]
            },
            "Name": field_key,
            "NativeReferenceName": y_field["name"]
        })
        added_fields.add(field_key)

    # SELECT: Tooltip fields
    for tooltip_name, tooltip_table in tooltips.items():
        field_key = f"{tooltip_table}.{tooltip_name}"

        if field_key in added_fields:
            continue

        agg = find_aggregation(tooltip_name, tooltips_agg)

        if agg:
            if "Min(" in agg:
                agg_name = f"Min({field_key})"
                native_name = tooltip_name

                select_fields.append({
                    "Aggregation": {
                        "Expression": {
                            "Column": {
                                "Expression": {
                                    "SourceRef": {
                                        "Source": table_aliases[tooltip_table]
                                    }
                                },
                                "Property": tooltip_name
                            }
                        },
                        "Function": 3
                    },
                    "Name": agg_name,
                    "NativeReferenceName": native_name
                })

                column_properties[agg_name] = {"displayName": native_name}

            elif "ATTR(" in agg:
                select_fields.append({
                    "Column": {
                        "Expression": {
                            "SourceRef": {
                                "Source": table_aliases[tooltip_table]
                            }
                        },
                        "Property": tooltip_name
                    },
                    "Name": field_key
                })

            else:
                select_fields.append({
                    "Measure": {
                        "Expression": {
                            "SourceRef": {
                                "Source": table_aliases[tooltip_table]
                            }
                        },
                        "Property": tooltip_name
                    },
                    "Name": field_key,
                    "NativeReferenceName": tooltip_name
                })

        else:
            select_fields.append({
                "Measure": {
                    "Expression": {
                        "SourceRef": {
                            "Source": table_aliases[tooltip_table]
                        }
                    },
                    "Property": tooltip_name
                },
                "Name": field_key,
                "NativeReferenceName": tooltip_name
            })

        added_fields.add(field_key)

    # ORDER BY
    order_by = []
    if y_field:
        order_by.append({
            "Direction": 2,
            "Expression": {
                "Measure": {
                    "Expression": {
                        "SourceRef": {
                            "Source": table_aliases[y_field["table"]]
                        }
                    },
                    "Property": y_field["name"]
                }
            }
        })

    # -----------------------------------
    # Fix the broken f-string safely here
    # -----------------------------------
    title_text = visual_data.get("title", "YTD '24 vs '23")

    # FINAL CONFIG BUILD
    config = {
        "name": ref_config.get("name", "7b799989ad444488aafd035d2bb7bec2"),
        "layouts": ref_config.get("layouts", [{
            "id": 0,
            "position": {
                "x": 196.23,
                "y": 148.20,
                "z": 1000,
                "width": 865.87,
                "height": 421.27,
                "tabOrder": 1000
            }
        }]),
        "singleVisual": {
            "visualType": "filledMap",
            "projections": projections,
            "prototypeQuery": {
                "Version": 2,
                "From": from_clause,
                "Select": select_fields,
                "OrderBy": order_by
            },
            "columnProperties": column_properties,
            "drillFilterOtherVisuals": True,
            "hasDefaultSort": True,
            "objects": ref_config.get("singleVisual", {}).get("objects", {}),
            "vcObjects": {
                "title": [{
                    "properties": {
                        "text": {
                            "expr": {
                                "Literal": {
                                    "Value": f"'{title_text}'"
                                }
                            }
                        }
                    }
                }],
                "border": [{
                    "properties": {
                        "show": {
                            "expr": {
                                "Literal": {
                                    "Value": "true"
                                }
                            }
                        }
                    }
                }]
            }
        }
    }

    return config



def generate_fieldmap_configs():
    """Main generator function"""
    print("\n" + "="*70)
    print("🗺️  POWER BI FIELD MAP GENERATOR")
    print("="*70 + "\n")
    
    # Step 1: Extract reference structure
    print("📖 Step 1: Extracting reference structure...")
    reference = extract_reference_structure(REFERENCE_FILE)
    if not reference:
        print("❌ Failed to extract reference structure. Exiting.")
        return
    
    # Step 2: Load input files
    print("\n📂 Step 2: Loading input files...")
    final_data = load_json(os.path.join(BASE_DIR, "final.json"))
    schema_data = load_json(os.path.join(BASE_DIR, "schema_output.json"))
    positions_data = load_json(os.path.join(BASE_DIR, "powerbi_chart_positions.json"))
    
    if not final_data:
        print("❌ Cannot proceed without final.json")
        return
    
    print(f"✅ Loaded {len(final_data)} visuals from final.json")
    print(f"✅ Loaded schema for {len(schema_data)} tables")
    print(f"✅ Loaded {len(positions_data)} position configs")
    
    # Step 3: Generate configs
    print("\n🛠️  Step 3: Generating configurations...\n")
    
    all_outputs = []
    
    for idx, visual in enumerate(final_data):
        source = visual.get("Source", f"Visual_{idx}")
        chart_type = visual.get("chart_type", "").lower()
        
        # Only process map charts
        if "map" not in chart_type:
            print(f"⏭️  Skipping '{source}' (not a map chart)")
            continue
        
        print(f"🛠️  Processing: {source}")
        
        # Build dynamic config
        config = build_dynamic_config(visual, schema_data, reference)
        
        # Get position data
        pos = next((p for p in positions_data if p.get("chart") == source), {})
        
        x = pos.get("x", reference.get("x", 196.23))
        y = pos.get("y", reference.get("y", 148.20))
        z = pos.get("z", reference.get("z", 1000))
        width = pos.get("width", reference.get("width", 865.87))
        height = pos.get("height", reference.get("height", 421.27))
        
        # Update layout position
        if config.get("layouts"):
            config["layouts"][0]["position"] = {
                "x": x,
                "y": y,
                "z": z,
                "width": width,
                "height": height,
                "tabOrder": z
            }
        
        # Build final output matching reference structure
        output = {
            "config": json.dumps(config),
            "filters": "[]",
            "height": height,
            "width": width,
            "x": x,
            "y": y,
            "z": z
        }
        
        all_outputs.append(output)
        print(f"   ✅ Configuration complete\n")
    
    # Step 4: Save output
    if all_outputs:
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            json.dump(all_outputs, f, indent=2, ensure_ascii=False)
        print(f"💾 Step 4: Saved {len(all_outputs)} configurations to {OUTPUT_FILE}")
    else:
        print("⚠️  No configurations generated")
    
    print("\n" + "="*70)
    print("✅ GENERATION COMPLETE")
    print("="*70 + "\n")


if __name__ == "__main__":
    generate_fieldmap_configs()