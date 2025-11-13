import os
import json
import re
import sys
from typing import Dict, Any, Optional, Tuple, List
from collections import Counter

# Import Gemini API
try:
    import google.generativeai as genai
except ImportError:
    print("ERROR: google.generativeai library not found. Install with: pip install google-generativeai")
    sys.exit(1)

# ==================== CONSTANTS ====================
BASE_DIR = os.getcwd()
OUTPUT_DIR = r"C:\celendar\output"

# File paths
FINAL_JSON_PATH = os.path.join(OUTPUT_DIR, "final.json")
SCHEMA_JSON_PATH = os.path.join(OUTPUT_DIR, "schema_output.json")
POSITIONS_JSON_PATH = os.path.join(OUTPUT_DIR, "powerbi_chart_positions.json")
COLORS_JSON_PATH = os.path.join(OUTPUT_DIR, "extracted_colors.json")
REFERENCE_TXT_PATH = os.path.join(OUTPUT_DIR, "Reference Chart Configurations.txt")
OUTPUT_JSON_PATH = os.path.join(OUTPUT_DIR, "visual_output.json")

# Gemini API configuration
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "AIzaSyAwOgHgl1qu1wAEqteRGgwv80cCB_caDS4")
GEMINI_MODEL_NAME = "gemini-2.5-flash"

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


def find_calendar_chart_visual(final_data: list) -> Optional[Dict]:
    """Find the calendar chart visual in final.json by chart_type="calendarChart"."""
    for visual in final_data:
        chart_type = visual.get("chart_type", "").strip().lower()
        title = visual.get("title", "").strip().lower()
        source = visual.get("Source", "").strip().lower()
        
        if (chart_type == "calendarchart" or 
            "calendário" in title or 
            "calendario" in title or
            "calendar" in title or
            "calendário" in source or
            "calendar" in source):
            print(f"Found calendar chart: title='{visual.get('title')}', type='{visual.get('chart_type')}', source='{visual.get('Source')}'")
            return visual
    
    print("WARNING: No calendar chart found")
    print("Available charts in final.json:")
    for visual in final_data:
        print(f"  - Source: '{visual.get('Source', 'N/A')}', Title: '{visual.get('title', 'N/A')}', Type: '{visual.get('chart_type', 'N/A')}'")
    return None


def extract_calendar_prototype(reference_text: str) -> Optional[Dict]:
    """Extract calendar chart prototype from Reference Chart Configurations.txt."""
    try:
        reference_obj = json.loads(reference_text)
        config_str = reference_obj.get("config")
        
        if not config_str:
            print("ERROR: 'config' key not found in reference file")
            return None
        
        config_obj = json.loads(config_str)
        print("Extracted calendar chart prototype from reference")
        return config_obj
    
    except json.JSONDecodeError as e:
        print(f"ERROR: Failed to parse JSON: {e}")
        return None
    except Exception as e:
        print(f"ERROR: Failed to extract prototype: {e}")
        return None


# ==================== PERMANENT DYNAMIC FIELD ROLE ENGINE ====================

class FieldRoleEngine:
    """
    Permanent, industrial-strength field role detection engine.
    Uses multi-signal analysis: schema, cardinality, semantics, and Gemini AI.
    """
    
    def __init__(self, schema_data: Dict, calendar_visual: Dict, schema_mapping: Dict):
        self.schema_data = schema_data
        self.calendar_visual = calendar_visual
        self.schema_mapping = schema_mapping
        self.all_fields = {}
        self.field_metadata = {}
        
        # Extract all fields from visual
        self._extract_all_fields()
        
    def _extract_all_fields(self):
        """Extract all fields from the calendar visual."""
        rows = self.calendar_visual.get("Rows") or {}
        columns = self.calendar_visual.get("Columns") or {}
        legend = self.calendar_visual.get("Legend") or {}
        
        self.all_fields = {**rows, **columns, **legend}
        
        print(f"\n{'='*60}")
        print(f"FIELD ROLE ENGINE: Extracted {len(self.all_fields)} fields")
        print(f"{'='*60}")
        for field_name, table in self.all_fields.items():
            print(f"  - {field_name} (Table: {table})")
        print(f"{'='*60}\n")
    
    def _get_field_datatype(self, field_name: str, table_name: str) -> Optional[str]:
        """Get datatype from schema for a field."""
        if not table_name or table_name not in self.schema_data:
            return None
        
        for col_info in self.schema_data[table_name]:
            if col_info.get("name", "").strip().lower() == field_name.strip().lower():
                return col_info.get("type", "").lower()
        
        return None
    
    def _analyze_schema_signals(self) -> Dict[str, str]:
        """Analyze schema datatypes to infer field roles."""
        roles = {}
        
        for field_name, table_name in self.all_fields.items():
            datatype = self._get_field_datatype(field_name, table_name)
            
            if not datatype:
                continue
            
            # Date detection
            if any(dt in datatype for dt in ["date", "datetime", "timestamp"]):
                if "date" not in roles:
                    roles["date"] = field_name
            
            # Numeric measures
            elif any(dt in datatype for dt in ["int", "float", "decimal", "number", "numeric"]):
                if "measure" not in roles:
                    roles["measure"] = field_name
            
            # Text categories
            elif any(dt in datatype for dt in ["text", "string", "varchar", "char"]):
                if "category" not in roles:
                    roles["category"] = field_name
        
        return roles
    
    def _analyze_cardinality_signals(self) -> Dict[str, str]:
        """
        Analyze field cardinality patterns.
        Note: Without actual data, we use heuristics based on field names.
        """
        roles = {}
        
        for field_name in self.all_fields.keys():
            field_lower = field_name.lower()
            
            # High cardinality indicators (event_index)
            if any(pattern in field_lower for pattern in ["index", "id", "row", "number", "#"]):
                if "event_index" not in roles:
                    roles["event_index"] = field_name
            
            # Low cardinality indicators (category/legend)
            elif any(pattern in field_lower for pattern in ["type", "category", "group", "status", "class"]):
                if "category" not in roles:
                    roles["category"] = field_name
            
            # Week/label patterns
            elif any(pattern in field_lower for pattern in ["week", "label", "name"]):
                if "week_label" not in roles:
                    roles["week_label"] = field_name
        
        return roles
    
    def _ask_gemini_for_classification(self) -> Dict[str, str]:
        """Use Gemini AI to semantically classify fields."""
        try:
            # Prepare field list
            field_list = list(self.all_fields.keys())
            
            # Get hierarchy info
            hierarchy = self.calendar_visual.get("Hierarchy") or []
            
            prompt = f"""
You are a Power BI data expert. Analyze these field names and classify each into EXACTLY ONE role.

Fields: {json.dumps(field_list, indent=2)}
Hierarchy: {json.dumps(hierarchy, indent=2)}

REQUIRED ROLES (you must assign fields to ALL of these):
1. date - The primary date/time field for the calendar
2. week_label - A label or category shown in calendar cells (can be week number, status, or any categorical label)
3. category - A grouping or classification field
4. legend - A field used for color coding or legend
5. measure - A numeric value field
6. event_index - A unique identifier or row index
7. event_group - A date hierarchy field (Year/Quarter/Month) for grouping events

RULES:
- If a field name contains "date", "dt", "data", "transaction" → date
- If a field name contains "week", "label" → week_label
- If a field name contains "index", "id", "row", "#" → event_index
- If hierarchy is present, use the date field referenced in it for event_group
- The legend field is typically used for color coding
- Measure is typically a numeric value
- Category is typically a text grouping field

Return ONLY a valid JSON object with this structure:
{{
  "date": "field_name",
  "week_label": "field_name",
  "category": "field_name",
  "legend": "field_name",
  "measure": "field_name",
  "event_index": "field_name",
  "event_group": "field_name"
}}

Do not include any explanation, markdown formatting, or additional text. Only return the JSON object.
"""
            
            response = model.generate_content(prompt)
            response_text = response.text.strip()
            
            # Clean response (remove markdown code blocks if present)
            response_text = re.sub(r'^```json\s*', '', response_text)
            response_text = re.sub(r'\s*```$', '', response_text)
            response_text = response_text.strip()
            
            gemini_roles = json.loads(response_text)
            
            print(f"\n{'='*60}")
            print("GEMINI AI CLASSIFICATION RESULTS:")
            print(f"{'='*60}")
            for role, field in gemini_roles.items():
                print(f"  {role}: {field}")
            print(f"{'='*60}\n")
            
            return gemini_roles
        
        except Exception as e:
            print(f"WARNING: Gemini classification failed: {e}")
            return {}
    
    def _merge_signals(self, schema_roles: Dict, cardinality_roles: Dict, gemini_roles: Dict) -> Dict[str, str]:
        """Merge all signal sources with priority: Gemini > Schema > Cardinality."""
        final_roles = {}
        
        # Priority 1: Gemini (highest confidence)
        for role, field in gemini_roles.items():
            if field and field in self.all_fields:
                final_roles[role] = field
        
        # Priority 2: Schema signals
        for role, field in schema_roles.items():
            if role not in final_roles and field in self.all_fields:
                final_roles[role] = field
        
        # Priority 3: Cardinality signals
        for role, field in cardinality_roles.items():
            if role not in final_roles and field in self.all_fields:
                final_roles[role] = field
        
        return final_roles
    
    def _apply_fallback_logic(self, roles: Dict[str, str]) -> Dict[str, str]:
        """Apply intelligent fallback logic for missing roles."""
        
        # Get all available fields
        available_fields = list(self.all_fields.keys())
        used_fields = set(roles.values())
        
        # Fallback for date
        if "date" not in roles or not roles["date"]:
            for field in available_fields:
                if any(pattern in field.lower() for pattern in ["date", "dt", "data", "transaction", "time"]):
                    if field not in used_fields:
                        roles["date"] = field
                        used_fields.add(field)
                        break
        
        # Fallback for event_index
        if "event_index" not in roles or not roles["event_index"]:
            for field in available_fields:
                if any(pattern in field.lower() for pattern in ["index", "id", "row", "#", "num"]):
                    if field not in used_fields:
                        roles["event_index"] = field
                        used_fields.add(field)
                        break
        
        # Fallback for week_label
        if "week_label" not in roles or not roles["week_label"]:
            for field in available_fields:
                if any(pattern in field.lower() for pattern in ["week", "label", "name", "status"]):
                    if field not in used_fields:
                        roles["week_label"] = field
                        used_fields.add(field)
                        break
        
        # Fallback for category
        if "category" not in roles or not roles["category"]:
            for field in available_fields:
                if field not in used_fields:
                    roles["category"] = field
                    used_fields.add(field)
                    break
        
        # Fallback for legend (can reuse week_label or category)
        if "legend" not in roles or not roles["legend"]:
            if "week_label" in roles:
                roles["legend"] = roles["week_label"]
            elif "category" in roles:
                roles["legend"] = roles["category"]
        
        # Fallback for measure (can reuse event_index)
        if "measure" not in roles or not roles["measure"]:
            if "event_index" in roles:
                roles["measure"] = roles["event_index"]
        
        # Fallback for event_group (use date field)
        if "event_group" not in roles or not roles["event_group"]:
            if "date" in roles:
                roles["event_group"] = roles["date"]
        
        # Ultimate fallback: use first available field for any missing role
        required_roles = ["date", "week_label", "category", "legend", "measure", "event_index", "event_group"]
        for role in required_roles:
            if role not in roles or not roles[role]:
                for field in available_fields:
                    roles[role] = field
                    break
        
        return roles
    
    def detect_roles(self) -> Dict[str, Tuple[str, str]]:
        """
        Main method: Detect all field roles using multi-signal analysis.
        Returns: Dict mapping role to (queryRef, table)
        """
        print(f"\n{'='*60}")
        print("PERMANENT FIELD ROLE DETECTION ENGINE - STARTING")
        print(f"{'='*60}\n")
        
        # Step 1: Schema analysis
        print("[1/5] Analyzing schema datatypes...")
        schema_roles = self._analyze_schema_signals()
        print(f"  Schema signals detected {len(schema_roles)} roles")
        
        # Step 2: Cardinality analysis
        print("[2/5] Analyzing field cardinality patterns...")
        cardinality_roles = self._analyze_cardinality_signals()
        print(f"  Cardinality signals detected {len(cardinality_roles)} roles")
        
        # Step 3: Gemini semantic classification
        print("[3/5] Querying Gemini AI for semantic classification...")
        gemini_roles = self._ask_gemini_for_classification()
        print(f"  Gemini classified {len(gemini_roles)} roles")
        
        # Step 4: Merge all signals
        print("[4/5] Merging all signal sources...")
        merged_roles = self._merge_signals(schema_roles, cardinality_roles, gemini_roles)
        
        # Step 5: Apply fallback logic
        print("[5/5] Applying fallback logic for missing roles...")
        final_roles = self._apply_fallback_logic(merged_roles)
        
        # Extract hierarchy information
        hierarchy = self.calendar_visual.get("Hierarchy") or []
        hierarchy_type = "Quarter"  # Default
        
        for hier_str in hierarchy:
            hier_match = re.search(r'(Quarter|Month|Year|Week|Weekday)', hier_str, re.IGNORECASE)
            if hier_match:
                hierarchy_type = hier_match.group(1).capitalize()
                break
        
        # Convert to queryRef format
        result = {}
        
        for role, field_name in final_roles.items():
            if not field_name or field_name not in self.all_fields:
                continue
            
            table_name = self.all_fields[field_name]
            
            # Build queryRef
            if role == "event_group":
                queryref = f"{table_name}.{field_name}.Variation.Date Hierarchy.{hierarchy_type}"
            else:
                queryref = f"{table_name}.{field_name}"
            
            result[role] = (queryref, table_name)
        
        # Map to Power BI projection names
        field_mappings = {
            'events': result.get('event_index', result.get('category', ('Unknown.Unknown', 'Unknown'))),
            'event_group': result.get('event_group', result.get('date', ('Unknown.Unknown', 'Unknown'))),
            'cell_color': result.get('week_label', result.get('legend', ('Unknown.Unknown', 'Unknown'))),
            'start_date': result.get('date', ('Unknown.Unknown', 'Unknown')),
            'end_date': result.get('date', ('Unknown.Unknown', 'Unknown')),
            'hierarchy_type': hierarchy_type
        }
        
        # Print final summary
        print(f"\n{'='*60}")
        print(f"FINAL FIELD ROLE ASSIGNMENTS")
        print(f"{'='*60}")
        print(f"  📅 Date Field:       {final_roles.get('date', 'NOT FOUND')} → {field_mappings['start_date'][0]}")
        print(f"  🎨 Cell Color:       {final_roles.get('week_label', 'NOT FOUND')} → {field_mappings['cell_color'][0]}")
        print(f"  📊 Events:           {final_roles.get('event_index', 'NOT FOUND')} → {field_mappings['events'][0]}")
        print(f"  📆 Event Group:      {final_roles.get('event_group', 'NOT FOUND')} → {field_mappings['event_group'][0]}")
        print(f"  🔧 Hierarchy Type:   {hierarchy_type}")
        print(f"{'='*60}\n")
        
        return field_mappings


# ==================== CONFIGURATION UPDATE FUNCTIONS ====================

def find_chart_position(chart_title: str, positions_data: list) -> Dict[str, float]:
    """Find position data for chart by title (case-insensitive match)."""
    title_lower = chart_title.strip().lower()
    
    for pos_entry in positions_data:
        entry_title = pos_entry.get("chart", "").strip().lower()
        if entry_title == title_lower or "calendário" in entry_title or "calendar" in entry_title:
            return {
                "x": float(pos_entry.get("x", 0)),
                "y": float(pos_entry.get("y", 0)),
                "z": float(pos_entry.get("z", 0)),
                "width": float(pos_entry.get("width", 1100)),
                "height": float(pos_entry.get("height", 400))
            }
    
    print(f"WARNING: No position found for chart '{chart_title}', using defaults")
    return {"x": 17.05, "y": 290.69, "z": 0.0, "width": 1103.83, "height": 413.16}


def update_calendar_config(prototype: Dict, field_mappings: Dict, position: Dict) -> Dict:
    """
    Update the prototype with new field mappings and position.
    Preserves ALL other structure from the reference including colors and formatting.
    """
    import copy
    config = copy.deepcopy(prototype)
    
    # Update projections ONLY
    if "singleVisual" in config and "projections" in config["singleVisual"]:
        projections = config["singleVisual"]["projections"]
        
        projections["events"] = [{"queryRef": field_mappings['events'][0]}]
        projections["EventGroup"] = [{"queryRef": field_mappings['event_group'][0]}]
        projections["CellColor"] = [{"queryRef": field_mappings['cell_color'][0]}]
        projections["StartDate"] = [{"queryRef": field_mappings['start_date'][0]}]
        projections["EndDate"] = [{"queryRef": field_mappings['end_date'][0]}]
    
    # Update prototypeQuery
    if "singleVisual" in config and "prototypeQuery" in config["singleVisual"]:
        proto_query = config["singleVisual"]["prototypeQuery"]
        
        # Update From clause
        if "From" in proto_query:
            main_table = field_mappings['events'][1]
            proto_query["From"] = [{"Name": "s", "Entity": main_table, "Type": 0}]
        
        # Update Select clause
        if "Select" in proto_query:
            hierarchy_type = field_mappings.get('hierarchy_type', 'Quarter')
            
            proto_query["Select"] = [
                {
                    "Column": {
                        "Expression": {"SourceRef": {"Source": "s"}},
                        "Property": field_mappings['events'][0].split('.')[-1]
                    },
                    "Name": field_mappings['events'][0],
                    "NativeReferenceName": field_mappings['events'][0].split('.')[-1]
                },
                {
                    "HierarchyLevel": {
                        "Expression": {
                            "Hierarchy": {
                                "Expression": {
                                    "PropertyVariationSource": {
                                        "Expression": {"SourceRef": {"Source": "s"}},
                                        "Name": "Variation",
                                        "Property": field_mappings['start_date'][0].split('.')[-1]
                                    }
                                },
                                "Hierarchy": "Date Hierarchy"
                            }
                        },
                        "Level": hierarchy_type
                    },
                    "Name": field_mappings['event_group'][0],
                    "NativeReferenceName": f"{field_mappings['start_date'][0].split('.')[-1]} {hierarchy_type}"
                },
                {
                    "Column": {
                        "Expression": {"SourceRef": {"Source": "s"}},
                        "Property": field_mappings['cell_color'][0].split('.')[-1]
                    },
                    "Name": field_mappings['cell_color'][0],
                    "NativeReferenceName": field_mappings['cell_color'][0].split('.')[-1]
                },
                {
                    "Column": {
                        "Expression": {"SourceRef": {"Source": "s"}},
                        "Property": field_mappings['start_date'][0].split('.')[-1]
                    },
                    "Name": field_mappings['start_date'][0],
                    "NativeReferenceName": field_mappings['start_date'][0].split('.')[-1]
                }
            ]
    
    # Update position in layouts
    if "layouts" in config and len(config["layouts"]) > 0:
        config["layouts"][0]["position"]["x"] = position["x"]
        config["layouts"][0]["position"]["y"] = position["y"]
        config["layouts"][0]["position"]["z"] = position["z"]
        config["layouts"][0]["position"]["width"] = position["width"]
        config["layouts"][0]["position"]["height"] = position["height"]
    
    return config


def validate_calendar_config(config_obj: Dict) -> bool:
    """Validate the calendar configuration structure."""
    try:
        if "singleVisual" not in config_obj:
            print("Validation failed: Missing singleVisual")
            return False
        
        if "visualType" not in config_obj["singleVisual"]:
            print("Validation failed: Missing visualType")
            return False
        
        visual_type = config_obj["singleVisual"]["visualType"]
        if "calendar" not in visual_type.lower():
            print(f"WARNING: Visual type doesn't appear to be a calendar: {visual_type}")
        
        projections = config_obj.get("singleVisual", {}).get("projections", {})
        required_projections = ["events", "EventGroup", "CellColor", "StartDate", "EndDate"]
        
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
        
        if "layouts" not in config_obj or len(config_obj["layouts"]) == 0:
            print("Validation failed: Missing layouts")
            return False
        
        position = config_obj["layouts"][0].get("position", {})
        required_pos = ["x", "y", "width", "height"]
        if not all(key in position for key in required_pos):
            print("Validation failed: Missing position properties")
            return False
        
        print("Configuration validation passed ✓")
        return True
    
    except Exception as e:
        print(f"Validation error: {e}")
        return False


# ==================== MAIN GENERATION FUNCTION ====================

def generate_calendar_chart():
    """Main function to generate the calendar chart configuration."""
    print("=" * 70)
    print("Power BI Calendar Chart Generator v2.0 - PERMANENT ENGINE")
    print("=" * 70)
    
    # Load input files
    print("\n[1/7] Loading input files...")
    
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
    
    # Find calendar chart visual
    print("\n[2/7] Finding calendar chart visual...")
    calendar_visual = find_calendar_chart_visual(final_data)
    if not calendar_visual:
        print("ERROR: No calendar chart found. Aborting.")
        return 1
    
    # Build schema mapping
    print("\n[3/7] Building schema mapping...")
    schema_mapping = build_schema_mapping(schema_data)
    
    # Find chart position
    print("\n[4/7] Finding chart position...")
    chart_title = calendar_visual.get("title", "Calendário")
    position = find_chart_position(chart_title, positions_data)
    print(f"Position: x={position['x']}, y={position['y']}, w={position['width']}, h={position['height']}")
    
    # Extract calendar chart prototype
    print("\n[5/7] Extracting calendar chart prototype from reference...")
    prototype = extract_calendar_prototype(reference_text)
    if not prototype:
        print("ERROR: Failed to extract prototype. Aborting.")
        return 1
    
    # PERMANENT DYNAMIC FIELD DETECTION ENGINE
    print("\n[6/7] Running permanent dynamic field detection engine...")
    engine = FieldRoleEngine(schema_data, calendar_visual, schema_mapping)
    field_mappings = engine.detect_roles()
    
    # Update prototype with new fields and position
    print("\n[7/7] Updating configuration...")
    final_config = update_calendar_config(prototype, field_mappings, position)
    
    # Final validation
    if not validate_calendar_config(final_config):
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
        print(f"\n✓ Successfully saved output to: {OUTPUT_JSON_PATH}")
    except Exception as e:
        print(f"ERROR: Failed to save output: {e}")
        return 1
    
    # Success message
    print("\n" + "=" * 70)
    print("✓ GENERATION COMPLETE")
    print("=" * 70)
    print(f"Output: {OUTPUT_JSON_PATH}")
    print("\nNext steps:")
    print("1. Open visual_output.json")
    print("2. Copy the 'config' string value")
    print("3. Paste it into your Power BI report.json as a visualContainer config")
    print("4. Make sure the calendar custom visual is installed in your Power BI")
    print("   Visual GUID: calendarVisual74934D05B71F4C31B0F79D925EE89638")
    print("=" * 70)
    
    return 0


# ==================== ENTRY POINT ====================

if __name__ == "__main__":
    try:
        exit_code = generate_calendar_chart()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n\nGeneration cancelled by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n\nFATAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)