#!/usr/bin/env python3
"""
bump_generate.py - Power BI Bump Chart Configuration Generator

Usage:
    python bump_generate.py

Input files (from C:\\linedot_chart\\output):
    - final.json
    - schema_output.json
    - powerbi_chart_positions.json
    - extracted_colors.json
    - Reference Chart Configurations.txt

Output file:
    - C:\\linedot_chart\\output\\visuals_output.json
"""

import os
import json
import re
import sys
from typing import Dict, Any, Optional, Tuple

# Import Gemini API
try:
    import google.generativeai as genai
except ImportError:
    print("ERROR: google.generativeai library not found. Install with: pip install google-generativeai")
    sys.exit(1)

# ==================== CONSTANTS ====================
BASE_DIR = os.getcwd()
OUTPUT_DIR = r"C:\linedot_chart\output"

# File paths
FINAL_JSON_PATH = os.path.join(OUTPUT_DIR, "final.json")
SCHEMA_JSON_PATH = os.path.join(OUTPUT_DIR, "schema_output.json")
POSITIONS_JSON_PATH = os.path.join(OUTPUT_DIR, "powerbi_chart_positions.json")
COLORS_JSON_PATH = os.path.join(OUTPUT_DIR, "extracted_colors.json")
REFERENCE_TXT_PATH = os.path.join(OUTPUT_DIR, "Reference Chart Configurations.txt")
OUTPUT_JSON_PATH = os.path.join(OUTPUT_DIR, "visuals_output.json")

# Gemini API configuration
GEMINI_API_KEY = "AIzaSyAwOgHgl1qu1wAEqteRGgwv80cCB_caDS4"
GEMINI_MODEL_NAME = "gemini-1.5-flash"

# Configure Gemini
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel(GEMINI_MODEL_NAME)

# ==================== HELPER FUNCTIONS ====================

def load_json_file(filepath: str) -> Any:
    """Load and parse a JSON file."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"ERROR: File not found: {filepath}")
        return None
    except json.JSONDecodeError as e:
        print(f"ERROR: Invalid JSON in {filepath}: {e}")
        return None
    except Exception as e:
        print(f"ERROR: Failed to load {filepath}: {e}")
        return None


def build_schema_mapping(schema_data: Dict) -> Dict[str, str]:
    """
    Build column to table mapping from schema_output.json.
    Handles case-insensitive lookups with whitespace stripping.
    
    Returns: Dict[column_lower, table_name]
    """
    mapping = {}
    for table_name, columns in schema_data.items():
        for col_info in columns:
            col_name = col_info.get("name", "").strip().lower()
            if col_name and col_name not in mapping:
                mapping[col_name] = table_name
    
    print(f"Built schema mapping with {len(mapping)} columns from {len(schema_data)} tables")
    return mapping


def find_bump_chart_visual(final_data: list) -> Optional[Dict]:
    """
    Find the bump chart visual in final.json by title="Ranking".
    """
    for visual in final_data:
        title = visual.get("title", "").strip()
        
        if title == "Ranking":
            print(f"Found bump chart: title='{title}'")
            return visual
    
    print("WARNING: No bump chart found with title='Ranking'")
    return None


def extract_bump_prototype(reference_text: str) -> Optional[Dict]:
    """
    Extract bump chart prototype from Reference Chart Configurations.txt.
    The file contains a JSON object with a "config" key that has a stringified JSON value.
    """
    try:
        # First, parse the outer JSON structure
        reference_obj = json.loads(reference_text)
        
        # Get the config string
        config_str = reference_obj.get("config")
        
        if not config_str:
            print("ERROR: 'config' key not found in reference file")
            return None
        
        # Parse the stringified config JSON
        config_obj = json.loads(config_str)
        
        print("Extracted bump chart prototype from reference")
        return config_obj
    
    except json.JSONDecodeError as e:
        print(f"ERROR: Failed to parse JSON: {e}")
        return None
    except Exception as e:
        print(f"ERROR: Failed to extract prototype: {e}")
        return None


def resolve_field_to_queryref(field_name: str, table_hint: Optional[str], 
                               schema_mapping: Dict, is_hierarchy: bool = False) -> str:
    """
    Resolve a field name to a Power BI queryRef string.
    """
    # Clean the field name
    clean_field = field_name.strip()
    
    # Extract aggregation if present (like "AGG(Ranking Select Case Medidas Rank)")
    agg_match = re.match(r'(?:AGG|SUM|AVG|COUNT|MIN|MAX)\((.+)\)', clean_field, re.IGNORECASE)
    if agg_match:
        clean_field = agg_match.group(1).strip()
    
    # Determine table name
    if table_hint:
        table_name = table_hint
    else:
        # Lookup in schema mapping (case-insensitive)
        field_lower = clean_field.lower()
        table_name = schema_mapping.get(field_lower)
        
        if not table_name:
            print(f"WARNING: Could not resolve table for field '{clean_field}'")
            table_name = "UnknownTable"
    
    # Build queryRef based on whether it's a hierarchy
    if is_hierarchy:
        # For Date hierarchies: Table.Field.Variation.Date Hierarchy.Month
        queryref = f"{table_name}.{clean_field}.Variation.Date Hierarchy.Month"
    else:
        # Standard format: Table.Field
        queryref = f"{table_name}.{clean_field}"
    
    return queryref


def extract_field_mappings(bump_visual: Dict, schema_mapping: Dict) -> Tuple[str, str, str, str]:
    """
    Extract and resolve category, legend, and measure fields from bump visual.
    
    Returns:
        Tuple of (category_queryref, legend_queryref, measure_queryref, measure_table)
    """
    # Extract fields from visual definition
    rows = bump_visual.get("Rows", {})
    columns = bump_visual.get("Columns", {})
    legend = bump_visual.get("Legend", {})
    hierarchy = bump_visual.get("Hierarchy", [])
    
    # Category (X-axis) - from Hierarchy or Columns
    category_field = None
    category_table = None
    is_hierarchy = False
    
    if hierarchy:
        # Parse hierarchy like "Month(SELL Through_tb_bif_agg_venda.Dt Venda)"
        hier_str = hierarchy[0] if isinstance(hierarchy, list) else hierarchy
        hier_match = re.search(r'(?:Month|Year|Quarter|Day)\(([^.]+)\.([^)]+)\)', hier_str, re.IGNORECASE)
        if hier_match:
            table, field = hier_match.groups()
            category_field = field.strip()
            category_table = table.strip()
            is_hierarchy = True
    
    if not category_field and columns:
        category_field = list(columns.keys())[0]
        category_table = columns.get(category_field)
    
    # Legend (Series) - from Legend field
    legend_field = list(legend.keys())[0] if legend else None
    legend_table = legend.get(legend_field) if legend_field else None
    
    # Measure (Y-axis) - from Rows
    measure_field = list(rows.keys())[0] if rows else None
    measure_table = rows.get(measure_field) if measure_field else None
    
    # Resolve to queryRefs
    category_queryref = resolve_field_to_queryref(
        category_field or "Dt Venda", 
        category_table, 
        schema_mapping, 
        is_hierarchy=is_hierarchy
    )
    
    legend_queryref = resolve_field_to_queryref(
        legend_field or "CC Case Dimensões Ranking",
        legend_table,
        schema_mapping
    )
    
    measure_queryref = resolve_field_to_queryref(
        measure_field or "Ranking Select Case Medidas Rank",
        measure_table,
        schema_mapping
    )
    
    print(f"Resolved field mappings:")
    print(f"  Category: {category_queryref}")
    print(f"  Legend: {legend_queryref}")
    print(f"  Measure: {measure_queryref}")
    print(f"  Measure Table: {measure_table}")
    
    return category_queryref, legend_queryref, measure_queryref, measure_table


def find_chart_position(chart_title: str, positions_data: list) -> Dict[str, float]:
    """
    Find position data for chart by title (case-insensitive match).
    """
    title_lower = chart_title.strip().lower()
    
    for pos_entry in positions_data:
        entry_title = pos_entry.get("chart", "").strip().lower()
        if entry_title == title_lower:
            return {
                "x": float(pos_entry.get("x", 0)),
                "y": float(pos_entry.get("y", 0)),
                "z": float(pos_entry.get("z", 0)),
                "width": float(pos_entry.get("width", 800)),
                "height": float(pos_entry.get("height", 400))
            }
    
    print(f"WARNING: No position found for chart '{chart_title}', using defaults")
    return {"x": 11.81, "y": 210.15, "z": 0.0, "width": 1075.2, "height": 252.6}


def update_prototype_config(prototype: Dict, category_qr: str, legend_qr: str, 
                            measure_qr: str, measure_table: str, position: Dict) -> Dict:
    """
    Update the prototype with new field mappings and position.
    Preserves ALL other structure from the reference.
    """
    import copy
    config = copy.deepcopy(prototype)
    
    # Update projections ONLY
    if "singleVisual" in config and "projections" in config["singleVisual"]:
        projections = config["singleVisual"]["projections"]
        
        # Update category (X-axis)
        projections["category"] = [
            {"queryRef": category_qr, "active": True}
        ]
        
        # Update legend (Series)
        projections["legend"] = [
            {"queryRef": legend_qr}
        ]
        
        # Update tooltip - use the base measure queryRef
        projections["tooltip"] = [
            {"queryRef": measure_qr}
        ]
        
        # Update measure (Y-axis) - use the measure with _qtc suffix
        # The measure table should match what's in the reference
        measure_base = measure_qr.split('.')[-1]  # Get just the field name
        measure_with_suffix = f"{measure_table}.{measure_base}_qtc"
        
        projections["measure"] = [
            {"queryRef": measure_with_suffix}
        ]
    
    # Update prototypeQuery Select fields to match our tables
    if "singleVisual" in config and "prototypeQuery" in config["singleVisual"]:
        proto_query = config["singleVisual"]["prototypeQuery"]
        if "Select" in proto_query:
            for select_item in proto_query["Select"]:
                # Update the Name fields to match our resolved queryRefs
                if "Name" in select_item:
                    old_name = select_item["Name"]
                    
                    # Update measure references
                    if "Ranking Select Case Medidas Rank" in old_name:
                        if "_qtc" in old_name:
                            select_item["Name"] = measure_with_suffix
                        else:
                            select_item["Name"] = measure_qr
                    
                    # Update legend reference
                    elif "CC Case Dimensões Ranking" in old_name or "CC Case DimensÃµes Ranking" in old_name:
                        select_item["Name"] = legend_qr
                    
                    # Update category reference
                    elif "Dt Venda" in old_name:
                        select_item["Name"] = category_qr
    
    # Update position in layouts
    if "layouts" in config and len(config["layouts"]) > 0:
        config["layouts"][0]["position"]["x"] = position["x"]
        config["layouts"][0]["position"]["y"] = position["y"]
        config["layouts"][0]["position"]["z"] = position["z"]
        config["layouts"][0]["position"]["width"] = position["width"]
        config["layouts"][0]["position"]["height"] = position["height"]
    
    return config


def build_gemini_prompt(base_config: Dict, bump_visual: Dict, schema_summary: str) -> str:
    """
    Build a text prompt for Gemini API.
    """
    prompt = f"""You are a Power BI configuration expert. Transform and validate this bump chart configuration.

**BASE CONFIGURATION** (from reference, already has correct structure):
```json
{json.dumps(base_config, indent=2, ensure_ascii=False)}
```

**VISUAL METADATA** from final.json (for context only):
```json
{json.dumps(bump_visual, indent=2, ensure_ascii=False)}
```

**SCHEMA SUMMARY**:
{schema_summary}

**CRITICAL INSTRUCTIONS**:
1. OUTPUT ONLY THE COMPLETE JSON CONFIGURATION - NO explanations, NO markdown code fences, NO text before or after
2. PRESERVE the exact structure from the base configuration
3. KEEP all the prototypeQuery exactly as-is
4. KEEP the visualType GUID exactly as-is
5. The projections (category, legend, measure) are already correct - don't change them
6. Verify all table names in queryRef strings exist in the schema
7. Ensure the JSON is valid and complete

Return the complete configuration JSON now:"""
    
    return prompt


def call_gemini_with_retry(prompt: str, max_retries: int = 2) -> Optional[str]:
    """
    Call Gemini API with retry logic.
    """
    import time
    
    for attempt in range(max_retries + 1):
        try:
            print(f"Calling Gemini API (attempt {attempt + 1}/{max_retries + 1})...")
            response = model.generate_content(prompt)
            result_text = response.text
            
            # Clean up markdown code fences if present
            result_text = re.sub(r'```json\s*', '', result_text)
            result_text = re.sub(r'```\s*$', '', result_text)
            result_text = result_text.strip()
            
            print("Received response from Gemini")
            return result_text
        
        except Exception as e:
            print(f"Gemini API error (attempt {attempt + 1}): {e}")
            if attempt < max_retries:
                wait_time = 2 ** attempt
                print(f"  Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                print("Max retries reached")
                return None
    
    return None


def validate_config(config_obj: Dict) -> bool:
    """
    Validate the configuration structure.
    """
    try:
        # Check basic structure
        if "singleVisual" not in config_obj:
            print("Validation failed: Missing singleVisual")
            return False
        
        if "visualType" not in config_obj["singleVisual"]:
            print("Validation failed: Missing visualType")
            return False
        
        # Check projections
        projections = config_obj.get("singleVisual", {}).get("projections", {})
        required_projections = ["category", "legend", "measure"]
        
        for proj in required_projections:
            if proj not in projections:
                print(f"Validation failed: Missing {proj} projection")
                return False
            
            if not isinstance(projections[proj], list) or len(projections[proj]) == 0:
                print(f"Validation failed: {proj} projection is empty or not a list")
                return False
            
            if "queryRef" not in projections[proj][0]:
                print(f"Validation failed: {proj} projection missing queryRef")
                return False
        
        # Check layouts and position
        if "layouts" not in config_obj or len(config_obj["layouts"]) == 0:
            print("Validation failed: Missing layouts")
            return False
        
        position = config_obj["layouts"][0].get("position", {})
        required_pos = ["x", "y", "width", "height"]
        if not all(key in position for key in required_pos):
            print("Validation failed: Missing position properties")
            return False
        
        print("Configuration validation passed")
        return True
    
    except Exception as e:
        print(f"Validation error: {e}")
        return False


def create_schema_summary(schema_data: Dict) -> str:
    """
    Create a concise schema summary for the prompt.
    """
    # Focus on the two main tables we need
    relevant_tables = [
        "SELL Through_tb_bif_agg_venda",
        "Estoque - Days on Hand_tb_bif_a",
        "BR_Mesorregioes_2020_BR_Mesorre"
    ]
    
    summary_lines = ["Relevant tables:"]
    
    for table_name in relevant_tables:
        if table_name in schema_data:
            col_names = [col["name"] for col in schema_data[table_name][:10]]
            summary_lines.append(f"  {table_name}: {', '.join(col_names)}...")
    
    return "\n".join(summary_lines)


# ==================== MAIN GENERATION FUNCTION ====================

def generate_bump_chart():
    """
    Main function to generate the bump chart configuration.
    """
    print("=" * 70)
    print("Power BI Bump Chart Generator")
    print("=" * 70)
    
    # Load input files
    print("\n[1/8] Loading input files...")
    
    final_data = load_json_file(FINAL_JSON_PATH)
    schema_data = load_json_file(SCHEMA_JSON_PATH)
    positions_data = load_json_file(POSITIONS_JSON_PATH)
    colors_data = load_json_file(COLORS_JSON_PATH)
    
    if not all([final_data, schema_data, positions_data]):
        print("ERROR: Critical files missing. Aborting.")
        return 1
    
    # Load reference text
    try:
        with open(REFERENCE_TXT_PATH, 'r', encoding='utf-8') as f:
            reference_text = f.read()
        print(f"Loaded reference text")
    except Exception as e:
        print(f"ERROR: Could not load reference text: {e}")
        return 1
    
    # Find bump chart visual
    print("\n[2/8] Finding bump chart visual...")
    bump_visual = find_bump_chart_visual(final_data)
    if not bump_visual:
        print("ERROR: No bump chart found. Aborting.")
        return 1
    
    # Build schema mapping
    print("\n[3/8] Building schema mapping...")
    schema_mapping = build_schema_mapping(schema_data)
    
    # Find chart position
    print("\n[4/8] Finding chart position...")
    chart_title = bump_visual.get("title", "Ranking")
    position = find_chart_position(chart_title, positions_data)
    print(f"Position: x={position['x']}, y={position['y']}, w={position['width']}, h={position['height']}")
    
    # Extract bump chart prototype
    print("\n[5/8] Extracting bump chart prototype from reference...")
    prototype = extract_bump_prototype(reference_text)
    if not prototype:
        print("ERROR: Failed to extract prototype. Aborting.")
        return 1
    
    # Extract field mappings
    print("\n[6/8] Extracting field mappings...")
    category_qr, legend_qr, measure_qr, measure_table = extract_field_mappings(bump_visual, schema_mapping)
    
    # Update prototype with new fields and position
    print("\n[7/8] Updating configuration...")
    base_config = update_prototype_config(prototype, category_qr, legend_qr, measure_qr, measure_table, position)
    
    # Prepare Gemini prompt
    schema_summary = create_schema_summary(schema_data)
    prompt = build_gemini_prompt(base_config, bump_visual, schema_summary)
    
    # Call Gemini (optional refinement)
    print("\n[8/8] Calling Gemini for validation...")
    gemini_response = call_gemini_with_retry(prompt)
    
    if gemini_response:
        try:
            final_config = json.loads(gemini_response)
            if validate_config(final_config):
                print("Using Gemini-validated configuration")
            else:
                print("Gemini response failed validation, using base configuration")
                final_config = base_config
        except json.JSONDecodeError as e:
            print(f"Failed to parse Gemini response: {e}")
            print("Using base configuration")
            final_config = base_config
    else:
        print("Using base configuration without Gemini validation")
        final_config = base_config
    
    # Final validation
    if not validate_config(final_config):
        print("ERROR: Final configuration is invalid")
        return 1
    
    # Create output structure
    output_visual = {
        "config": json.dumps(final_config, ensure_ascii=False),
        "filters": "[]",
        "x": position["x"],
        "y": position["y"],
        "z": position["z"],
        "width": position["width"],
        "height": position["height"]
    }
    
    # Save to file
    try:
        with open(OUTPUT_JSON_PATH, 'w', encoding='utf-8') as f:
            json.dump(output_visual, f, indent=2, ensure_ascii=False)
        print(f"\nSuccessfully saved output to: {OUTPUT_JSON_PATH}")
    except Exception as e:
        print(f"ERROR: Failed to save output: {e}")
        return 1
    
    # Success message
    print("\n" + "=" * 70)
    print("GENERATION COMPLETE")
    print("=" * 70)
    print(f"Output: {OUTPUT_JSON_PATH}")
    print("\nNext steps:")
    print("1. Open visuals_output.json")
    print("2. Copy the 'config' string value")
    print("3. Paste it into your Power BI report.json as a visualContainer config")
    print("=" * 70)
    
    return 0


# ==================== ENTRY POINT ====================

if __name__ == "__main__":
    try:
        exit_code = generate_bump_chart()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n\nGeneration cancelled by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n\nFATAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)