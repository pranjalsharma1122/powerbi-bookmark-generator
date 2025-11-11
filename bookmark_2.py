import json
import os
import re
import uuid
from typing import Dict, List, Set, Optional, Tuple

# ============================================================================
# CONFIGURATION
# ============================================================================
BASE_DIR = r"C:\Bookmark\output"
REFERENCE_FILE = r"C:\Bookmark\output\Reference Chart Configurations.txt"
VISUALS_INPUT = os.path.join(BASE_DIR, "visuals_output.json")
ACTIONS_INPUT = os.path.join(BASE_DIR, "combined_parameter_actions.json")
OUTPUT_FILE = os.path.join(BASE_DIR, "visual_output.json")

def gen_guid() -> str:
    return uuid.uuid4().hex[:20]

def load_json(filepath: str) -> dict:
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

def norm(text: str) -> str:
    """Normalize text for fuzzy matching"""
    if not text:
        return ""
    return re.sub(r'[^a-z0-9]', '', text.strip().lower())

def smart_match_chart(target_name: str, available_visuals: List[dict]) -> Optional[dict]:
    """
    Smart fuzzy matching for chart names
    Returns the best matching visual or None
    """
    if not target_name:
        return None
    
    target_norm = norm(target_name)
    best_match = None
    best_score = 0
    
    for visual in available_visuals:
        title = visual.get("title", "")
        source = visual.get("source", "")
        
        for candidate in [title, source]:
            if not candidate:
                continue
            
            candidate_norm = norm(candidate)
            score = 0
            
            # Exact match
            if target_norm == candidate_norm:
                score = 100
            # Target contained in candidate
            elif target_norm in candidate_norm:
                # Prefer shorter matches (more specific)
                score = 80 - (len(candidate_norm) - len(target_norm))
            # Candidate contained in target
            elif candidate_norm in target_norm:
                score = 60
            # Check for partial word matches
            else:
                target_words = set(target_norm.split())
                candidate_words = set(candidate_norm.split())
                common = target_words & candidate_words
                if common:
                    score = 40 + (len(common) * 10)
            
            if score > best_score:
                best_score = score
                best_match = visual
    
    # Only return matches with confidence >= 50
    if best_score >= 50:
        return best_match
    return None

def create_action_button(button_name: str, bookmark_guid: str, position: dict) -> dict:
    """Create a Power BI action button linked to a bookmark"""
    button_guid = gen_guid()
    button_config = {
        "name": button_guid,
        "layouts": [{
            "id": 0,
            "position": {
                "x": position["x"],
                "y": position["y"],
                "z": position.get("z", 10000),
                "width": position["width"],
                "height": position["height"],
                "tabOrder": position.get("tabOrder", 0)
            }
        }],
        "singleVisual": {
            "visualType": "actionButton",
            "drillFilterOtherVisuals": True,
            "objects": {
                "icon": [{
                    "properties": {
                        "shapeType": {"expr": {"Literal": {"Value": "'blank'"}}}
                    },
                    "selector": {"id": "default"}
                }],
                "text": [
                    {"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}}},
                    {
                        "properties": {
                            "text": {"expr": {"Literal": {"Value": f"'{button_name}'"}}}
                        },
                        "selector": {"id": "default"}
                    }
                ]
            },
            "vcObjects": {
                "title": [{
                    "properties": {
                        "show": {"expr": {"Literal": {"Value": "false"}}}
                    }
                }],
                "visualLink": [{
                    "properties": {
                        "show": {"expr": {"Literal": {"Value": "true"}}},
                        "type": {"expr": {"Literal": {"Value": "'Bookmark'"}}},
                        "bookmark": {"expr": {"Literal": {"Value": f"'{bookmark_guid}'"}}}
                    }
                }]
            }
        }
    }
    
    return {
        "guid": button_guid,
        "name": button_name,
        "config": json.dumps(button_config, separators=(',', ':')),
        "filters": "[]",
        "height": position["height"],
        "width": position["width"],
        "x": position["x"],
        "y": position["y"],
        "z": position.get("z", 10000)
    }

def create_powerbi_report():
    print("\n" + "="*70)
    print("STAGE 2: Creating Power BI Report with Multiple Bookmarks")
    print("="*70 + "\n")
    
    visuals_data = load_json(VISUALS_INPUT)
    actions_data = load_json(ACTIONS_INPUT)
    reference_data = load_json(REFERENCE_FILE)
    
    if not visuals_data:
        print("❌ Cannot proceed without visuals_output.json")
        return
    
    print(f"✅ Loaded {len(visuals_data)} visuals")
    print(f"✅ Loaded {len(actions_data)} actions\n")
    
    # Load reference config
    ref_config = {}
    if isinstance(reference_data, dict) and "config" in reference_data:
        try:
            ref_config = json.loads(reference_data["config"])
        except:
            pass
    
    # ========================================================================
    # 🔥 STEP 1: Parse all actions and create bookmark metadata
    # ========================================================================
    bookmark_metadata = []
    
    if not actions_data:
        print("⚠️ No actions found, creating default bookmark")
        bookmark_metadata.append({
            "action_idx": 0,
            "action_name": "Default",
            "bookmark_name": "Default View",
            "zone_charts": [v.get("title", v.get("source", "")) for v in visuals_data]
        })
    else:
        for action_idx, action in enumerate(actions_data):
            action_name = action.get("caption", f"Action {action_idx+1}")
            print(f"📌 Processing Action {action_idx+1}: '{action_name}'")
            
            # Extract bookmark names from source-field parameter
            bookmark_names = []
            for param in action.get("params", []):
                if param.get("name") == "source-field":
                    raw_values = param.get("values_in_field", [])
                    
                    # Handle boolean values
                    if raw_values and isinstance(raw_values[0], bool):
                        bookmark_names = [
                            f"{action_name} - {'Show' if v else 'Hide'}" 
                            for v in raw_values
                        ]
                    else:
                        bookmark_names = [str(v) for v in raw_values if v]
                    
                    print(f"   📋 Bookmark names: {bookmark_names}")
                    break
            
            if not bookmark_names:
                print(f"   ⚠️ No bookmarks found in action, skipping")
                continue
            
            # Extract zones for this action
            matching_zones = action.get("matching_zones", [])
            print(f"   📋 Zones in action: {len(matching_zones)}")
            
            # Ensure we have enough zones for all bookmarks
            if len(matching_zones) < len(bookmark_names):
                print(f"   ⚠️ Not enough zones ({len(matching_zones)}) for bookmarks ({len(bookmark_names)})")
                # For hide bookmarks, add empty list (no charts visible)
                matching_zones.append([])
            
            # Create metadata for each bookmark
            for bm_idx, bm_name in enumerate(bookmark_names):
                zone = matching_zones[bm_idx] if bm_idx < len(matching_zones) else []
                
                # 🔥 FIX: Extract chart names and skip text/image elements
                zone_charts = []
                for item in zone:
                    if isinstance(item, dict):
                        chart_name = item.get("name")
                        item_type = item.get("type") or ""
                        
                        # Skip text elements, images, and items without names
                        if not chart_name:
                            continue
                        if item_type.lower() in ["text", "image"]:
                            continue
                        
                        # Only add actual chart elements
                        zone_charts.append(chart_name)
                
                print(f"   🔖 '{bm_name}' → {len(zone_charts)} charts: {zone_charts}")
                
                bookmark_metadata.append({
                    "action_idx": action_idx,
                    "action_name": action_name,
                    "bookmark_name": bm_name,
                    "zone_charts": zone_charts
                })
            
            print()
    
    print(f"✅ Total bookmarks to create: {len(bookmark_metadata)}\n")
    
    # ========================================================================
    # 🔥 STEP 2: Match zone charts to actual visuals (IMPROVED)
    # ========================================================================
    print("="*70)
    print("MATCHING CHARTS TO VISUALS")
    print("="*70 + "\n")
    
    for metadata in bookmark_metadata:
        bm_name = metadata["bookmark_name"]
        zone_charts = metadata["zone_charts"]
        
        print(f"🔖 Bookmark: '{bm_name}'")
        print(f"   Target charts: {zone_charts}")
        
        matched_visuals = []
        
        for chart_name in zone_charts:
            # Try smart matching
            matched_visual = smart_match_chart(chart_name, visuals_data)
            
            if matched_visual:
                visual_guid = matched_visual["name"]
                visual_title = matched_visual.get("title", matched_visual.get("source", ""))
                matched_visuals.append(matched_visual)
                print(f"   ✅ '{chart_name}' → '{visual_title}' ({visual_guid[:8]}...)")
            else:
                print(f"   ⚠️ '{chart_name}' → No match found (skipping)")
        
        # Store matched visuals in metadata
        metadata["matched_visuals"] = matched_visuals
        
        print(f"   📊 Total matched: {len(matched_visuals)} visuals\n")
    
    # ========================================================================
    # 🔥 STEP 3: Build section with ALL available visuals
    # ========================================================================
    print("="*70)
    print("BUILDING PAGE SECTION")
    print("="*70 + "\n")
    
    section_guid = gen_guid()
    section = {
        "name": section_guid,
        "displayName": "Page 1",
        "config": "{}",
        "displayOption": 1,
        "filters": "[]",
        "height": 720.0,
        "width": 1280.0,
        "visualContainers": []
    }
    
    # 🔥 FIX: Add ALL visuals from visuals_output.json to the page
    visuals_by_guid = {}
    
    for visual in visuals_data:
        visual_guid = visual["name"]
        
        if visual_guid not in visuals_by_guid:
            visuals_by_guid[visual_guid] = visual
            
            # Add to section
            section["visualContainers"].append({
                "config": visual.get("config", "{}"),
                "filters": visual.get("filters", "[]"),
                "height": visual.get("height", 400),
                "width": visual.get("width", 640),
                "x": visual.get("x", 0),
                "y": visual.get("y", 0),
                "z": visual.get("z", 0)
            })
    
    all_visuals_on_page = list(visuals_by_guid.values())
    print(f"✅ Added ALL {len(all_visuals_on_page)} visuals to page\n")
    
    # ========================================================================
    # 🔥 STEP 4: Create bookmarks with proper visibility states
    # ========================================================================
    print("="*70)
    print("CREATING BOOKMARKS")
    print("="*70 + "\n")
    
    bookmarks = []
    bookmark_guid_map = {}
    
    for metadata in bookmark_metadata:
        bm_name = metadata["bookmark_name"]
        matched_visuals = metadata["matched_visuals"]
        
        bm_guid = gen_guid()
        bookmark_guid_map[bm_name] = bm_guid
        
        # Build visual container states
        visual_containers = {}
        
        # Get GUIDs of visuals that should be visible
        visible_guids = {v["name"] for v in matched_visuals}
        
        for visual in all_visuals_on_page:
            visual_guid = visual["name"]
            
            try:
                visual_cfg = json.loads(visual.get("config", "{}"))
                visual_type = visual_cfg.get("singleVisual", {}).get("visualType", "lineChart")
                
                state = {
                    "singleVisual": {
                        "visualType": visual_type,
                        "objects": {}
                    }
                }
                
                # Hide visuals not in this bookmark's zone
                if visual_guid not in visible_guids:
                    state["singleVisual"]["display"] = {"mode": "hidden"}
                
                visual_containers[visual_guid] = state
            except Exception as e:
                print(f"   ⚠️ Error processing visual {visual_guid[:8]}: {e}")
                continue
        
        bookmark = {
            "displayName": bm_name,
            "name": bm_guid,
            "explorationState": {
                "version": "1.3",
                "activeSection": section_guid,
                "sections": {
                    section_guid: {
                        "visualContainers": visual_containers
                    }
                },
                "objects": {}
            },
            "options": {"targetVisualNames": []}
        }
        
        bookmarks.append(bookmark)
        
        print(f"🔖 '{bm_name}'")
        print(f"   Shows: {len(visible_guids)} visuals")
        print(f"   Hides: {len(all_visuals_on_page) - len(visible_guids)} visuals")
        print(f"   GUID: {bm_guid[:12]}...\n")
    
    # ========================================================================
    # 🔥 STEP 5: Create action buttons for all bookmarks
    # ========================================================================
    print("="*70)
    print("CREATING ACTION BUTTONS")
    print("="*70 + "\n")
    
    button_x = 104
    button_spacing = 160
    button_y = 41.17
    
    for i, metadata in enumerate(bookmark_metadata):
        bm_name = metadata["bookmark_name"]
        bm_guid = bookmark_guid_map[bm_name]
        
        button_pos = {
            "x": button_x + (i * button_spacing),
            "y": button_y,
            "width": 145.46,
            "height": 61.75,
            "z": 10000 + i
        }
        
        button = create_action_button(bm_name, bm_guid, button_pos)
        
        section["visualContainers"].append({
            "config": button["config"],
            "filters": "[]",
            "height": button["height"],
            "width": button["width"],
            "x": button["x"],
            "y": button["y"],
            "z": button["z"]
        })
        
        print(f"🔘 Button '{bm_name}' → x={button_pos['x']}, links to {bm_guid[:12]}...")
    
    print()
    
    # ========================================================================
    # 🔥 STEP 6: Assemble final report structure
    # ========================================================================
    print("="*70)
    print("ASSEMBLING FINAL REPORT")
    print("="*70 + "\n")
    
    main_config = {
        "version": ref_config.get("version", "5.67"),
        "themeCollection": ref_config.get("themeCollection", {
            "baseTheme": {
                "name": "CY25SU10",
                "version": {"visual": "2.1.0", "report": "3.0.0", "page": "2.3.0"},
                "type": 2
            }
        }),
        "activeSectionIndex": 0,
        "bookmarks": bookmarks,
        "defaultDrillFilterOtherVisuals": True,
        "linguisticSchemaSyncVersion": 0,
        "settings": ref_config.get("settings", {
            "useNewFilterPaneExperience": True,
            "allowChangeFilterTypes": True,
            "useStylableVisualContainerHeader": True,
            "queryLimitOption": 6,
            "exportDataMode": 1,
            "useDefaultAggregateDisplayName": True,
            "useEnhancedTooltips": True
        }),
        "objects": ref_config.get("objects", {})
    }
    
    output = {
        "config": json.dumps(main_config, separators=(',', ':')),
        "layoutOptimization": 0,
        "resourcePackages": reference_data.get("resourcePackages", [{
            "resourcePackage": {
                "disabled": False,
                "items": [{
                    "name": "CY25SU10",
                    "path": "BaseThemes/CY25SU10.json",
                    "type": 202
                }],
                "name": "SharedResources",
                "type": 2
            }
        }]),
        "sections": [section]
    }
    
    # Save output
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    
    print(f"✅ SUCCESS! Report saved to: {OUTPUT_FILE}")
    print(f"\n📊 Summary:")
    print(f"   - Bookmarks: {len(bookmarks)}")
    print(f"   - Action Buttons: {len(bookmark_metadata)}")
    print(f"   - Total Visuals on Page: {len(all_visuals_on_page)}")
    print(f"   - Total Visual Containers: {len(section['visualContainers'])}")
    print(f"   - Section GUID: {section_guid[:12]}...")
    
    print("\n" + "="*70)
    print("📝 NEXT STEPS:")
    print("="*70)
    print("1. Copy the content of visual_output.json")
    print("2. Open your Power BI .pbip folder")
    print("3. Navigate to: YourReport.Report/.platform")
    print("4. Open 'report.json' and replace its entire content")
    print("5. Save and open Power BI Desktop")
    print("6. Load the .pbip file")
    print("7. Go to View tab → Bookmarks pane")
    print("8. Test each bookmark by clicking the buttons!")
    print("="*70 + "\n")

if __name__ == "__main__":
    create_powerbi_report()