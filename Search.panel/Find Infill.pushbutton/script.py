from pyrevit import revit, DB, forms
from System.Collections.Generic import List  # Import .NET List

# Access the current document and UI document
doc = revit.doc
uidoc = revit.uidoc

try:
    # Prompt user for search input
    search_input = forms.ask_for_string(
        prompt="Enter the 'Part Number' values to search for Windows (separated by commas):",
        title="Search Windows",
        default=""
    )
    
    if search_input:
        # Split the input by commas, strip whitespace, and convert to lowercase for case-insensitive matching
        search_terms = [term.strip().lower() for term in search_input.split(',') if term.strip()]
        
        if not search_terms:
            forms.alert("No valid 'Part Number' values provided.", title="Invalid Input", warn_icon=True)
        else:
            # Collect all Windows in the model
            collector = DB.FilteredElementCollector(doc)\
                        .OfCategory(DB.BuiltInCategory.OST_Windows)\
                        .WhereElementIsNotElementType()
            
            # Filter elements by their "Part Number" parameter for exact matches
            matching_elements = []
            for el in collector:
                # Use LookupParameter to access custom or shared parameters by name
                part_number_param = el.LookupParameter("Part Number")
                if part_number_param:
                    part_number_value = part_number_param.AsString()
                    if part_number_value:
                        part_number_lower = part_number_value.lower()
                        # Check for exact matches
                        if part_number_lower in search_terms:
                            matching_elements.append(el)
            
            # Select and highlight matching elements
            if matching_elements:
                # Convert Python list to .NET List[ElementId]
                element_ids = List[DB.ElementId]([el.Id for el in matching_elements])
                uidoc.Selection.SetElementIds(element_ids)
                forms.alert(
                    "Found and selected {} Windows with exact 'Part Number' values: {}.".format(
                        len(matching_elements), ", ".join(search_terms)
                    ),
                    title="Search Complete",
                    warn_icon=False
                )
            else:
                forms.alert("No Windows found with exact 'Part Number' values: {}.".format(
                    ", ".join(search_terms)), title="No Results", warn_icon=True)
    else:
        forms.alert("Search cancelled or no input provided.", title="Search Cancelled", warn_icon=True)

except Exception as e:
    forms.alert("An error occurred: {}".format(str(e)), title="Error", warn_icon=True)
