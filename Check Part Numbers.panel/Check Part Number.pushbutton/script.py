# -*- coding: UTF-8 -*-
"""
Checks for inconsistent Part Numbers in elements (e.g., glass panels)
that share common defining parameter values.
Ignores blank Part Numbers when checking for discrepancies.
Uses pyrevit's DB alias for Revit API access.
"""

from pyrevit import revit, DB, forms, script 
from System.Collections.Generic import List  

# --- Script Configuration ---

# VITAL: Parameters that define a unique "type" of part.
# Ensure these names EXACTLY match the parameters in your Revit elements.
KEY_PARAMETER_NAMES = ["Side 1", "Side 2", "Comments", "Material", "Glazing Step", "Frit Type"] 

# Parameter whose consistency needs to be checked within each group.
VALUE_PARAMETER_NAME = "Part Number"

# List of BuiltInCategories to search for these elements.
RELEVANT_CATEGORIES_BIC = [
    DB.BuiltInCategory.OST_Windows,
    # To add more, use comma and DB.BuiltInCategory.OST_AnotherCategory, e.g.:
    # DB.BuiltInCategory.OST_CurtainPanels,
    # DB.BuiltInCategory.OST_GenericModel,
]

# --- End of Script Configuration ---

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

def get_param_value_as_string(element, param_name):
    """
    Safely retrieves a parameter's value as a string using DB prefixed types.
    """
    p = element.LookupParameter(param_name)
    if p:
        if p.StorageType == DB.StorageType.ElementId: 
            el_id = p.AsElementId()
            if el_id and el_id != DB.ElementId.InvalidElementId:
                mat_elem = doc.GetElement(el_id)
                return mat_elem.Name if mat_elem else "Invalid Material ID"
            return "No Material Assigned"
        elif p.HasValue:
            val_str = p.AsValueString()
            if val_str is not None and val_str != "": 
                return val_str
            
            storage_type = p.StorageType 
            if storage_type == DB.StorageType.String:
                s_val = p.AsString()
                return s_val if s_val is not None else ""
            elif storage_type == DB.StorageType.Double:
                return str(p.AsDouble())
            elif storage_type == DB.StorageType.Integer:
                return str(p.AsInteger())
            else:
                try:
                    return str(p.AsObject())
                except Exception:
                    return "ERROR_READING_PARAM_VALUE"
        else:
            return "" # Parameter exists but has no value, treat as empty string
    return "PARAM_NOT_FOUND" # Parameter does not exist on the element

def run_check():
    """Main function to perform the check and report discrepancies."""
    elements_data = []

    if not RELEVANT_CATEGORIES_BIC:
        forms.alert("The 'RELEVANT_CATEGORIES_BIC' list is empty. Please specify categories to check.", exitscript=True)
    
    bic_list_net = List[DB.BuiltInCategory]() 
    for cat_bic_enum_val in RELEVANT_CATEGORIES_BIC:
        bic_list_net.Add(cat_bic_enum_val)
    
    if bic_list_net.Count == 0: 
        forms.alert("No valid categories specified for filtering.", exitscript=True)
        
    category_filter = DB.ElementMulticategoryFilter(bic_list_net)
    collected_elements = DB.FilteredElementCollector(doc).WherePasses(category_filter).WhereElementIsNotElementType().ToElements()

    if not collected_elements:
        cat_names_str = [str(c).replace("OST_", "") for c in RELEVANT_CATEGORIES_BIC]
        forms.alert("No elements found in the specified categories: {}.\nPlease check the 'RELEVANT_CATEGORIES_BIC' list or model content.".format(", ".join(cat_names_str)), exitscript=True)

    output.print_md("Processing {} elements from categories: {}...".format(len(collected_elements), ", ".join([str(c).replace("OST_", "") for c in RELEVANT_CATEGORIES_BIC])))

    for el in collected_elements:
        current_key_values = []
        valid_element_for_grouping = True
        for p_name in KEY_PARAMETER_NAMES: # Uses the corrected KEY_PARAMETER_NAMES
            val = get_param_value_as_string(el, p_name)
            if val == "PARAM_NOT_FOUND":
                valid_element_for_grouping = False
                break
            current_key_values.append(val)
        
        if not valid_element_for_grouping:
            # This element is missing one of the key parameters, so it can't be grouped.
            # Depending on strictness, you might want to log this or inform the user.
            # For now, it's skipped from the consistency check.
            continue

        part_number_val = get_param_value_as_string(el, VALUE_PARAMETER_NAME)
        if part_number_val == "PARAM_NOT_FOUND": 
            part_number_val = "PART_NUMBER_PARAM_MISSING" 
        
        elements_data.append({
            'key': tuple(current_key_values),
            'part_number': part_number_val,
            'id': el.Id,
            'element_obj': el 
        })

    if not elements_data:
        # This message now means no elements had ALL the KEY_PARAMETER_NAMES
        forms.alert("No elements found possessing all the required key parameters: {}.\nPlease verify parameter names and that elements in the selected categories have these parameters.".format(", ".join(KEY_PARAMETER_NAMES)), exitscript=True)

    grouped_by_key = {}
    for data_item in elements_data:
        key = data_item['key']
        if key not in grouped_by_key:
            grouped_by_key[key] = []
        grouped_by_key[key].append(data_item)

    discrepancies_found_count = 0
    report_lines = ["# Part Number Consistency Check Report"]
    report_lines.append("---")
    report_lines.append("**Grouping Criteria (Key Parameters):** `{}`".format("`, `".join(KEY_PARAMETER_NAMES)))
    report_lines.append("**Checking for consistent (non-blank):** `{}`".format(VALUE_PARAMETER_NAME))
    report_lines.append("**Categories Scanned:** `{}`".format("`, `".join([str(c).replace("OST_", "") for c in RELEVANT_CATEGORIES_BIC])))
    report_lines.append("---")

    all_discrepant_element_ids_for_selection = [] 

    for key, items_in_group in grouped_by_key.items():
        if len(items_in_group) <= 1: 
            continue

        non_blank_part_numbers_set = set()
        for item in items_in_group:
            pn = item['part_number']
            if pn and pn != "PART_NUMBER_PARAM_MISSING": 
                non_blank_part_numbers_set.add(pn)
        
        if len(non_blank_part_numbers_set) > 1:
            discrepancies_found_count += 1
            report_lines.append("## DISCREPANCY FOUND for Type:")
            
            key_str_parts = []
            for i, p_name in enumerate(KEY_PARAMETER_NAMES):
                key_str_parts.append("* **{}**: '{}'".format(p_name, key[i]))
            report_lines.append("\n".join(key_str_parts))
            
            report_lines.append("\n  **Differing non-blank Part Numbers found:**")
            
            pn_to_ids_and_count_for_reporting = {}
            for item in items_in_group:
                pn = item['part_number']
                if pn in non_blank_part_numbers_set: 
                    if pn not in pn_to_ids_and_count_for_reporting:
                        pn_to_ids_and_count_for_reporting[pn] = {'ids': [], 'count': 0}
                    pn_to_ids_and_count_for_reporting[pn]['ids'].append(item['id'])
                    pn_to_ids_and_count_for_reporting[pn]['count'] += 1
                    all_discrepant_element_ids_for_selection.append(item['id']) 

            for pn_report, data_report in sorted(pn_to_ids_and_count_for_reporting.items()):
                ids_str_list = [str(eid.IntegerValue) for eid in data_report['ids']]
                if len(ids_str_list) > 5:
                    ids_display_str = ", ".join(ids_str_list[:3]) + "... and {} more".format(len(ids_str_list)-3)
                else:
                    ids_display_str = ", ".join(ids_str_list)
                
                report_lines.append("    * Part Number: `'{}'` (Count: {}) (Element IDs: {})".format(
                    pn_report, 
                    data_report['count'], 
                    ids_display_str
                ))
            report_lines.append("---")

    if discrepancies_found_count == 0:
        report_lines.append("\n**No discrepancies (among non-blank Part Numbers) found matching the criteria.**")
    else:
        report_lines.insert(4, "**Total Discrepancy Groups (with differing non-blank Part Numbers): {}**".format(discrepancies_found_count))

    final_report = "\n".join(report_lines)
    output.print_md("--- PART NUMBER CONSISTENCY REPORT ---") 
    output.print_md(final_report) 
    output.print_md("--- END OF REPORT ---") 

    if all_discrepant_element_ids_for_selection: 
        unique_discrepant_ids = list(set(all_discrepant_element_ids_for_selection)) 
        selection_ids_net_list = List[DB.ElementId]()
        for eid_py in unique_discrepant_ids: 
            selection_ids_net_list.Add(eid_py)

        if forms.alert("Discrepancies found. Do you want to select all {} involved elements (with differing non-blank Part Numbers) in the model?".format(selection_ids_net_list.Count), 
                       yes=True, no=True): 
            if selection_ids_net_list.Count > 0:
                uidoc.Selection.SetElementIds(selection_ids_net_list)
                output.print_md("\n**Selected {} elements involved in discrepancies.**".format(selection_ids_net_list.Count))
            else: 
                output.print_md("\n**No elements to select for discrepancies.**")

if __name__ == '__main__':
    run_check()