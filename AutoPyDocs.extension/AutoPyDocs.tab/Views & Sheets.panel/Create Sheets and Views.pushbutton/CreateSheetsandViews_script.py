# -*- coding: utf-8 -*-   
__title__ = "Create Floor Plan views and place them on created sheets"
__author__ = "abhiramnair"

# Imports
import clr
import os
import time
import xlrd

clr.AddReference("RevitAPI")
from pyrevit import revit, forms, script
import Autodesk.Revit.DB as DB
from datetime import datetime
from Extract.RunData import get_run_data
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List


# Record the start time
start_time = time.time()
manual_time = 300
tool_name = __title__ 
app = __revit__.Application  # Returns the Revit Application Object
ui_doc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document  # Get the Active Document
model_name = doc.Title
output = script.get_output()
rvt_year = "Revit " + str(app.VersionNumber)
user_name = app.Username
header_data = " | ".join([
    datetime.now().strftime("%Y-%m-%d %H:%M:%S"), rvt_year, tool_name, model_name, user_name
])



# Step 1: User selects discipline
disciplines = ["AR-Architecture", "SC-Structural Concrete" , "SS-Structural Steel", "AG-Signage", "AI-Interiors", "AC-Acoustics", "FLS-Fire & Life Safety"]
selected_discipline = forms.SelectFromList.show(disciplines, title="Select Discipline", multiselect=False)
if not selected_discipline:
    TaskDialog.Show("Error", "No discipline selected.")
    script.exit()



# Get all title block symbols (family types)
title_block_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_TitleBlocks).ToElements()

# Collect family names to display in the selection list
title_block_family_names = list(set(symbol.Family.Name for symbol in title_block_symbols))

# Check if any title block family names were retrieved
if not title_block_family_names:
    TaskDialog.Show("Error", "No title block families available in the project.")
    script.exit()

# Display a selection form for title block families
selected_family_name = forms.SelectFromList.show(title_block_family_names, title="Select Title Block Family for Sheets", multiselect=False)

# Check if a title block family was selected
if not selected_family_name:
    TaskDialog.Show("Error", "No title block family selected.")
    script.exit()

# Filter the symbols to get only those belonging to the selected family
selected_family_symbols = [symbol for symbol in title_block_symbols if symbol.Family.Name == selected_family_name]

# Check if any types are available within the selected family
if not selected_family_symbols:
    TaskDialog.Show("Error", "No title block types available in the selected family.")
    script.exit()

# Display the available types in a selection list
title_block_type_names = [symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() for symbol in selected_family_symbols]
selected_type_name = forms.SelectFromList.show(title_block_type_names, title="Select Title Block Type for Sheets", multiselect=False)

# Retrieve the selected type for placing on sheets
selected_title_block_type = next((symbol for symbol in selected_family_symbols if symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == selected_type_name), None)

# Verify title block type selection
if not selected_title_block_type:
    TaskDialog.Show("Error", "Selected title block type not found.")
    script.exit()


# Step 2: User selects Excel file
excel_file = forms.pick_file(file_ext='xlsx', title="Select Excel File")
if not excel_file:
    TaskDialog.Show("Error", "No Excel file selected.")
    script.exit()

# Load workbook and select the correct sheet
excel_workbook = xlrd.open_workbook(excel_file)
excel_worksheet = excel_workbook.sheet_by_index(0)  # Adjust sheet index as needed


def duplicate_template(base_template, scale, new_template_name):
    # Retrieve the current view template names to ensure uniqueness
    existing_names = [vt.Name.lower() for vt in FilteredElementCollector(doc)
                      .OfClass(View)
                      .WhereElementIsNotElementType()
                      .ToElements() if vt.IsTemplate]

    if new_template_name.lower() in existing_names:
        print("Template name '{}' already exists.".format(new_template_name))
        return

    # Copy the template
    element_ids = List[ElementId]([base_template.Id])
    options = CopyPasteOptions()

    copied_ids = ElementTransformUtils.CopyElements(
        doc, element_ids, doc, Transform.Identity, options
    )

    v_new = doc.GetElement(copied_ids[0])
    
    # Ensure the new view is a template
    if not v_new.IsTemplate:
        v_new.IsTemplate = True

    # Set the new template name
    try:
        v_new.Name = new_template_name
    except Exception as e:
        print("Failed to set the name '{}': {}".format(new_template_name, e))
        return

    # Set the view scale and detail level
    detail_level_param = v_new.LookupParameter("Detail Level")
    view_scale_param = v_new.LookupParameter("View Scale")

    if detail_level_param and view_scale_param:
        non_controlled_ids = List[ElementId]([detail_level_param.Id, view_scale_param.Id])
        v_new.SetNonControlledTemplateParameterIds(non_controlled_ids)

        # Set view scale
        view_scale_param.Set(scale)

        # Set detail level to Medium for scales >= 300
        if scale >= 300:
            detail_level_param.Set(int(ViewDetailLevel.Medium))

        # Remove non-controlled IDs
        non_controlled_ids.Remove(detail_level_param.Id)
        non_controlled_ids.Remove(view_scale_param.Id)
        v_new.SetNonControlledTemplateParameterIds(non_controlled_ids)
    else:
        print("Failed to locate view parameters.")



# Function to convert float to string without decimals if it's an integer
def clean_float_value(value):
    try:
        # Check if the value is a float and remove decimals if it’s effectively an integer
        float_value = float(value)
        if float_value.is_integer():
            return str(int(float_value))  # Convert to int to remove decimals
        else:
            return str(value)  # Keep as string if it has decimals
    except ValueError:
        return str(value)  # If it's not a float, return it as string
        
combined_data = []
for row in range(1, excel_worksheet.nrows):  # Start from 1 to skip the header

    col_a = str(excel_worksheet.cell_value(row,  0)).strip()  #INSTANCE PARAMETER
    col_b = str(excel_worksheet.cell_value(row,  1)).strip()  #INSTANCE PARAMETER
    col_c = str(excel_worksheet.cell_value(row,  2)).strip()  #INSTANCE PARAMETER
    col_d = str(excel_worksheet.cell_value(row,  3)).strip()  #INSTANCE PARAMETER
    col_e = str(excel_worksheet.cell_value(row,  4)).strip()  #INSTANCE PARAMETER
    col_f = str(excel_worksheet.cell_value(row,  5)).strip()  #INSTANCE PARAMETER
    col_g = str(excel_worksheet.cell_value(row,  6)).strip()  #INSTANCE PARAMETER
    col_h = str(excel_worksheet.cell_value(row,  7)).strip()  #INSTANCE PARAMETER
    col_i = str(excel_worksheet.cell_value(row,  8)).strip()  #INSTANCE PARAMETER


    col_j = str(excel_worksheet.cell_value(row,  9)).strip()  #SHEET NUMBER
    col_k = str(excel_worksheet.cell_value(row, 10)).strip()  #SHEET NAME
    col_l = str(excel_worksheet.cell_value(row, 11)).strip()  #VIEW TYPE
    col_m = str(excel_worksheet.cell_value(row, 12)).strip()  #SCALE
    col_n = str(excel_worksheet.cell_value(row, 13)).strip()  #LEVEL
    col_o = str(excel_worksheet.cell_value(row, 14)).strip()  #ELEVATION/SECTION VIEW NAME
    col_p = str(excel_worksheet.cell_value(row, 15)).strip()  #SCOPE BOX


    col_q = str(excel_worksheet.cell_value(row, 16)).strip()  #DRAWN BY
    col_r = str(excel_worksheet.cell_value(row, 17)).strip()  #DESIGNED BY
    col_s = str(excel_worksheet.cell_value(row, 18)).strip()  #CHECKED BY
    col_t = str(excel_worksheet.cell_value(row, 19)).strip()  #APPROVED BY


    # Apply the function to all columns that may contain numeric values
    col_a = clean_float_value(col_a)
    col_b = clean_float_value(col_b)
    col_c = clean_float_value(col_c)
    col_d = clean_float_value(col_d)

    col_f = clean_float_value(col_f)
    col_g = clean_float_value(col_g)

    col_i = clean_float_value(col_i)
    col_j = clean_float_value(col_j)
    col_k = clean_float_value(col_k)
    col_l = clean_float_value(col_l)
    col_m = clean_float_value(col_m)
    col_n = clean_float_value(col_n)
    col_o = clean_float_value(col_o)
    col_p = clean_float_value(col_p)
    col_q = clean_float_value(col_q)
    col_r = clean_float_value(col_r)
    col_s = clean_float_value(col_s)
    col_t = clean_float_value(col_t)


    # Update column I values to "FLOOR PLAN" if they match specific keywords
    if col_l.lower() in ["key plan", "overall plan", "floor plan"]:
        col_l = "FLOOR PLAN"
    if col_l.lower() in ["framing plan", "structual framing plan", "framing"]:
        col_l = "STRUCTURAL FRAMING PLAN"
    if col_l.lower() in ["foundation plan", "structual foundation plan", "foundation"]:
        col_l = "STRUCTURAL FOUNDATION PLAN"
    col_l = col_l.title()
    # Replace ":" with "/" in column H values
    col_m = col_m.replace(":", "/")


    split_discipline = selected_discipline.split("-")[0]
    # Append updated row data
    combined_data.append([col_a, col_b, col_c, col_d, col_e, col_f, col_g, col_h,col_i, col_j, col_k, col_l,col_m,col_n,col_o, col_p, col_q, col_r,col_s, col_t, split_discipline])


# Step 4: Retrieve and filter view templates
view_templates = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).ToElements()
floor_plan_templates = [vt for vt in view_templates if vt.IsTemplate and vt.ViewType == ViewType.FloorPlan]
reflected_ceiling_plan_templates = [vt for vt in view_templates if vt.IsTemplate and vt.ViewType == ViewType.CeilingPlan]
elevation_template = [vt for vt in view_templates if vt.IsTemplate and vt.ViewType == ViewType.Elevation]


# Remove templates with " Wall" in their names
elevation_templates = [vt for vt in elevation_template if "wall" not in vt.Name.strip().lower()]



all_existing_templates = floor_plan_templates + elevation_templates + reflected_ceiling_plan_templates


#Check if the required templates exist based on Column J and K values

templates_to_create = []
for row in combined_data:
    col_l, col_m = row[11], row[12]
    # Check for scales 1:150 or 1:300 in col_j and construct template name accordingly
    if col_m in ["1/100", "1/150", "1/250", "1/300", "1/500"]:
        template_name = "{}_{}_{}".format(split_discipline, col_l, col_m)



        if "elevation" in template_name.lower():
            template_exists = next((vt for vt in elevation_templates if vt.Name.lower() == template_name.lower()),None)
            if template_exists == None:
                templates_to_create.append(template_name)
            else:
                continue

        elif "section" in template_name.lower():
            template_exists = next((vt for vt in elevation_templates if vt.Name.lower() == template_name.lower()),None)
            if template_exists == None:
                templates_to_create.append(template_name)
            else:
                continue    
        
        elif "framing" in template_name.lower():
            template_exists = next((vt for vt in reflected_ceiling_plan_templates if vt.Name.lower() == template_name.lower()),None)
            if template_exists == None:
                templates_to_create.append(template_name)
            else:
                continue    

        else:
            template_exists = next((vt for vt in floor_plan_templates if vt.Name.lower() == template_name.lower()),None)

            if template_exists == None:
                templates_to_create.append(template_name)

            else:
                continue

t = Transaction(doc, "Create templates")
t.Start()

templates_to_create = set(templates_to_create)
existing_template_names = {vt.Name.lower() for vt in all_existing_templates}  # Cache all existing template names in lowercase

for template_name in templates_to_create:


    if '1/100' in template_name:
        if template_name.lower() in existing_template_names:
            continue
        scale = 100
        # Create a new template name with 1/100
        new_template_name = template_name.replace("1/100", "1/100")  # Change the scale to 1/100

        # Check if the base template exists in the document
        base_plan_template = next((vt for vt in floor_plan_templates if vt.Name.lower() == new_template_name.lower()),None)
        # Special case for structural plans: Create "ST/SS_Structural Foundation Plan_1/100" if it doesn't exist
        if not base_plan_template and "Structural Foundation" in new_template_name:
            structural_foundation_template = next((vt for vt in floor_plan_templates if vt.Name.lower() == "Structural Foundation Plan".lower()), None)
            if structural_foundation_template:
                element_ids = List[ElementId]([structural_foundation_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_structural_foundation_template = doc.GetElement(copied_ids[0])
                base_plan_template = new_structural_foundation_template


        # Check if the base template exists in the document
        base_elevation_template = next((vt for vt in elevation_templates if vt.Name.lower() == new_template_name.lower()),None)
        base_arch_elevation_template = None
        # Special case for elevations: Create "Elevation_1/100" if it doesn't exist
        if not base_elevation_template and "{}_Elevation_1/100".format(split_discipline) in new_template_name:
            base_arch_elevation_template = next((vt for vt in elevation_templates if vt.Name.lower() == "architectural elevation"), None)
            if base_arch_elevation_template:
                element_ids = List[ElementId]([base_arch_elevation_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_elevation_template = doc.GetElement(copied_ids[0])
                base_elevation_template = new_elevation_template


        # Check if the base template exists in the document
        base_section_template = next((vt for vt in elevation_templates if vt.Name.lower() == new_template_name.lower()),None)
        base_arch_section_template = None
        # Special case for sections: Create "AR_Section_1/100" if it doesn't exist
        if not base_section_template and "{}_Section_1/100".format(split_discipline) in new_template_name:
            base_arch_section_template = next((vt for vt in elevation_templates if vt.Name.lower() == "architectural section"), None)
            if base_arch_section_template:
                element_ids = List[ElementId]([base_arch_section_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_section_template = doc.GetElement(copied_ids[0])
                base_section_template = new_section_template

        # Check if the base template exists in the document
        base_rcp_template = next((vt for vt in reflected_ceiling_plan_templates if vt.Name.lower() == new_template_name.lower()),None)
        # Special case for sections: Create "AR_Section_1/100" if it doesn't exist
        if not base_rcp_template and "Structural Framing" in new_template_name:
            base_rcp_template = next((vt for vt in reflected_ceiling_plan_templates if vt.Name.lower() == "structural framing plan"), None)
            if base_rcp_template:
                element_ids = List[ElementId]([base_rcp_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_rcp_template = doc.GetElement(copied_ids[0])
                base_rcp_template = new_rcp_template


        # Process each template type
        if base_plan_template:
            duplicate_template(base_plan_template, scale, template_name)
        if base_elevation_template:
            duplicate_template(base_elevation_template, scale, template_name)
        if base_section_template:
            duplicate_template(base_section_template, scale, template_name)
        if base_rcp_template:
            duplicate_template(base_rcp_template, scale, template_name)



    if '1/150' in template_name:
        if template_name.lower() in existing_template_names:
            continue
        scale = 150
        # Create a new template name with 1/100
        new_template_name = template_name.replace("1/150", "1/100")  # Change the scale to 1/100


        # Check if the base template exists in the document
        base_plan_template = next((vt for vt in floor_plan_templates if vt.Name.lower() == new_template_name.lower()),None)
        # Special case for structural plans: Create "ST/SS_Structural Foundation Plan_1/100" if it doesn't exist
        if not base_plan_template and "Structural Foundation" in new_template_name:
            structural_foundation_template = next((vt for vt in floor_plan_templates if vt.Name.lower() == "Structural Foundation Plan".lower()), None)
            if structural_foundation_template:
                element_ids = List[ElementId]([structural_foundation_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_structural_foundation_template = doc.GetElement(copied_ids[0])
                base_plan_template = new_structural_foundation_template


        # Check if the base template exists in the document
        base_elevation_template = next((vt for vt in elevation_templates if vt.Name.lower() == new_template_name.lower()),None)
        base_arch_elevation_template = None
        # Special case for elevations: Create "Elevation_1/100" if it doesn't exist
        if not base_elevation_template and "{}_Elevation_1/100".format(split_discipline) in new_template_name:
            base_arch_elevation_template = next((vt for vt in elevation_templates if vt.Name.lower() == "architectural elevation"), None)
            if base_arch_elevation_template:
                element_ids = List[ElementId]([base_arch_elevation_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_elevation_template = doc.GetElement(copied_ids[0])
                base_elevation_template = new_elevation_template


        # Check if the base template exists in the document
        base_section_template = next((vt for vt in elevation_templates if vt.Name.lower() == new_template_name.lower()),None)
        base_arch_section_template = None
        # Special case for sections: Create "AR_Section_1/100" if it doesn't exist
        if not base_section_template and "{}_Section_1/100".format(split_discipline) in new_template_name:
            base_arch_section_template = next((vt for vt in elevation_templates if vt.Name.lower() == "architectural section"), None)
            if base_arch_section_template:
                element_ids = List[ElementId]([base_arch_section_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_section_template = doc.GetElement(copied_ids[0])
                base_section_template = new_section_template


        # Check if the base template exists in the document
        base_rcp_template = next((vt for vt in reflected_ceiling_plan_templates if vt.Name.lower() == new_template_name.lower()),None)
        # Special case for sections: Create "AR_Section_1/100" if it doesn't exist
        if not base_rcp_template and "Structural Framing" in new_template_name:
            base_rcp_template = next((vt for vt in reflected_ceiling_plan_templates if vt.Name.lower() == "structural framing plan"), None)
            if base_rcp_template:
                element_ids = List[ElementId]([base_rcp_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_rcp_template = doc.GetElement(copied_ids[0])
                base_rcp_template = new_rcp_template


        # Process each template type
        if base_plan_template:
            duplicate_template(base_plan_template, scale, template_name)
        if base_elevation_template:
            duplicate_template(base_elevation_template, scale, template_name)
        if base_section_template:
            duplicate_template(base_section_template, scale, template_name)
        if base_rcp_template:
            duplicate_template(base_rcp_template, scale, template_name)


    if '1/250' in template_name:
        if template_name.lower() in existing_template_names:
            continue
        scale = 250
        # Create a new template name with 1/100
        new_template_name = template_name.replace("1/250", "1/100")  # Change the scale to 1/100



        # Check if the base template exists in the document
        base_plan_template = next(
            (vt for vt in floor_plan_templates if vt.Name.lower() == new_template_name.lower()),
            None
        )
        # Special case for structural plans: Create "ST/SS_Structural Foundation Plan_1/100" if it doesn't exist
        if not base_plan_template and "Structural Foundation" in new_template_name:
            structural_foundation_template = next((vt for vt in floor_plan_templates if vt.Name.lower() == "Structural Foundation Plan".lower()), None)
            if structural_foundation_template:
                element_ids = List[ElementId]([structural_foundation_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_structural_foundation_template = doc.GetElement(copied_ids[0])
                base_plan_template = new_structural_foundation_template


        # Check if the base template exists in the document
        base_elevation_template = next((vt for vt in elevation_templates if vt.Name.lower() == new_template_name.lower()),None)
        base_arch_elevation_template = None
        # Special case for elevations: Create "Elevation_1/100" if it doesn't exist
        if not base_elevation_template and "{}_Elevation_1/100".format(split_discipline) in new_template_name:
            base_arch_elevation_template = next((vt for vt in elevation_templates if vt.Name.lower() == "architectural elevation"), None)
            if base_arch_elevation_template:
                element_ids = List[ElementId]([base_arch_elevation_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_elevation_template = doc.GetElement(copied_ids[0])
                base_elevation_template = new_elevation_template


        # Check if the base template exists in the document
        base_section_template = next((vt for vt in elevation_templates if vt.Name.lower() == new_template_name.lower()),None)
        base_arch_section_template = None
        # Special case for sections: Create "AR_Section_1/100" if it doesn't exist
        if not base_section_template and "{}_Section_1/100".format(split_discipline) in new_template_name:
            base_arch_section_template = next((vt for vt in elevation_templates if vt.Name.lower() == "architectural section"), None)
            if base_arch_section_template:
                element_ids = List[ElementId]([base_arch_section_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_section_template = doc.GetElement(copied_ids[0])
                base_section_template = new_section_template


        # Check if the base template exists in the document
        base_rcp_template = next((vt for vt in reflected_ceiling_plan_templates if vt.Name.lower() == new_template_name.lower()),None)
        # Special case for sections: Create "AR_Section_1/100" if it doesn't exist
        if not base_rcp_template and "Structural Framing" in new_template_name:
            base_rcp_template = next((vt for vt in reflected_ceiling_plan_templates if vt.Name.lower() == "structural framing plan"), None)
            if base_rcp_template:
                element_ids = List[ElementId]([base_rcp_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_rcp_template = doc.GetElement(copied_ids[0])
                base_rcp_template = new_rcp_template

        # Process each template type
        if base_plan_template:
            duplicate_template(base_plan_template, scale, template_name)
        if base_elevation_template:
            duplicate_template(base_elevation_template, scale, template_name)
        if base_section_template:
            duplicate_template(base_section_template, scale, template_name)
        if base_rcp_template:
            duplicate_template(base_rcp_template, scale, template_name)


    # Check if the template name has a scale pattern we want to modify
    if '1/300' in template_name:
        if template_name.lower() in existing_template_names:
            continue
        scale = 300
        # Create a new template name with 1/100
        new_template_name = template_name.replace("1/300", "1/100")  # Change the scale to 1/100


        # Check if the base template exists in the document
        base_plan_template = next(
            (vt for vt in floor_plan_templates if vt.Name.lower() == new_template_name.lower()),
            None
        )
        # Special case for structural plans: Create "ST/SS_Structural Foundation Plan_1/100" if it doesn't exist
        if not base_plan_template and "Structural Foundation" in new_template_name:
            structural_foundation_template = next((vt for vt in floor_plan_templates if vt.Name.lower() == "Structural Foundation Plan".lower()), None)
            if structural_foundation_template:
                element_ids = List[ElementId]([structural_foundation_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_structural_foundation_template = doc.GetElement(copied_ids[0])
                base_plan_template = new_structural_foundation_template


        # Check if the base template exists in the document
        base_elevation_template = next((vt for vt in elevation_templates if vt.Name.lower() == new_template_name.lower()),None)
        base_arch_elevation_template = None
        # Special case for elevations: Create "Elevation_1/100" if it doesn't exist
        if not base_elevation_template and "{}_Elevation_1/100".format(split_discipline) in new_template_name:
            base_arch_elevation_template = next((vt for vt in elevation_templates if vt.Name.lower() == "architectural elevation"), None)
            if base_arch_elevation_template:
                element_ids = List[ElementId]([base_arch_elevation_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_elevation_template = doc.GetElement(copied_ids[0])
                base_elevation_template = new_elevation_template


        # Check if the base template exists in the document
        base_section_template = next((vt for vt in elevation_templates if vt.Name.lower() == new_template_name.lower()),None)
        base_arch_section_template = None
        # Special case for sections: Create "AR_Section_1/100" if it doesn't exist
        if not base_section_template and "{}_Section_1/100".format(split_discipline) in new_template_name:
            base_arch_section_template = next((vt for vt in elevation_templates if vt.Name.lower() == "architectural section"), None)
            if base_arch_section_template:
                element_ids = List[ElementId]([base_arch_section_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_section_template = doc.GetElement(copied_ids[0])
                base_section_template = new_section_template

        # Check if the base template exists in the document
        base_rcp_template = next((vt for vt in reflected_ceiling_plan_templates if vt.Name.lower() == new_template_name.lower()),None)
        # Special case for sections: Create "AR_Section_1/100" if it doesn't exist
        if not base_rcp_template and "Structural Framing" in new_template_name:
            base_rcp_template = next((vt for vt in reflected_ceiling_plan_templates if vt.Name.lower() == "structural framing plan"), None)
            if base_rcp_template:
                element_ids = List[ElementId]([base_rcp_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_rcp_template = doc.GetElement(copied_ids[0])
                base_rcp_template = new_rcp_template

        # Process each template type
        if base_plan_template:
            duplicate_template(base_plan_template, scale, template_name)
        if base_elevation_template:
            duplicate_template(base_elevation_template, scale, template_name)
        if base_section_template:
            duplicate_template(base_section_template, scale, template_name)
        if base_rcp_template:
            duplicate_template(base_rcp_template, scale, template_name)



    # Check if the template name has a scale pattern we want to modify
    if '1/500' in template_name:
        if template_name.lower() in existing_template_names:
            continue

        scale = 500
        # Create a new template name with 1/100
        new_template_name = template_name.replace("1/500", "1/100")  # Change the scale to 1/100


        # Check if the base template exists in the document
        base_plan_template = next(
            (vt for vt in floor_plan_templates if vt.Name.lower() == new_template_name.lower()),
            None
        )
        # Special case for structural plans: Create "ST/SS_Structural Foundation Plan_1/100" if it doesn't exist
        if not base_plan_template and "Structural Foundation" in new_template_name:
            structural_foundation_template = next((vt for vt in floor_plan_templates if vt.Name.lower() == "Structural Foundation Plan".lower()), None)
            if structural_foundation_template:
                element_ids = List[ElementId]([structural_foundation_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_structural_foundation_template = doc.GetElement(copied_ids[0])
                base_plan_template = new_structural_foundation_template


        # Check if the base template exists in the document
        base_elevation_template = next((vt for vt in elevation_templates if vt.Name.lower() == new_template_name.lower()),None)
        base_arch_section_template = None
        # Special case for elevations: Create "Elevation_1/100" if it doesn't exist
        if not base_elevation_template and "{}_Elevation_1/100".format(split_discipline) in new_template_name:
            base_arch_elevation_template = next((vt for vt in elevation_templates if vt.Name.lower() == "architectural elevation"), None)
            if base_arch_elevation_template:
                element_ids = List[ElementId]([base_arch_elevation_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_elevation_template = doc.GetElement(copied_ids[0])
                base_elevation_template = new_elevation_template


        # Check if the base template exists in the document
        base_section_template = next((vt for vt in elevation_templates if vt.Name.lower() == new_template_name.lower()),None)
        base_arch_section_template = None
        # Special case for sections: Create "AR_Section_1/100" if it doesn't exist
        if not base_section_template and "{}_Section_1/100".format(split_discipline) in new_template_name:
            base_arch_section_template = next((vt for vt in elevation_templates if vt.Name.lower() == "architectural section"), None)
            if base_arch_section_template:
                element_ids = List[ElementId]([base_arch_section_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_section_template = doc.GetElement(copied_ids[0])
                base_section_template = new_section_template

                     

        # Check if the base template exists in the document
        base_rcp_template = next((vt for vt in reflected_ceiling_plan_templates if vt.Name.lower() == new_template_name.lower()),None)
        # Special case for sections: Create "AR_Section_1/100" if it doesn't exist
        if not base_rcp_template and "Structural Framing" in new_template_name:
            base_rcp_template = next((vt for vt in reflected_ceiling_plan_templates if vt.Name.lower() == "structural framing plan"), None)
            if base_rcp_template:
                element_ids = List[ElementId]([base_rcp_template.Id])
                options = CopyPasteOptions()
                copied_ids = ElementTransformUtils.CopyElements(doc, element_ids, doc, Transform.Identity, options)
                new_rcp_template = doc.GetElement(copied_ids[0])
                base_rcp_template = new_rcp_template

        # Process each template type
        if base_plan_template:
            duplicate_template(base_plan_template, scale, template_name)
        if base_elevation_template:
            duplicate_template(base_elevation_template, scale, template_name)
        if base_section_template:
            duplicate_template(base_section_template, scale, template_name)
        if base_rcp_template:
            duplicate_template(base_rcp_template, scale, template_name)

t.Commit()




# Step 4: Retrieve and filter view templates
updated_view_templates = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).ToElements()

# Filter templates for Floor Plan, Elevation, and Section
updated_floor_plan_templates = [vt for vt in updated_view_templates if vt.IsTemplate and vt.ViewType == ViewType.FloorPlan]
updated_elevation_template = [vt for vt in updated_view_templates if vt.IsTemplate and vt.ViewType == ViewType.Elevation]

# Remove templates with " Wall" in their names
updated_elevation_templates = [vt for vt in updated_elevation_template if "wall" not in vt.Name.strip().lower()]


updated_rcp_templates = [vt for vt in updated_view_templates if vt.IsTemplate and vt.ViewType == ViewType.CeilingPlan]

all_templates = updated_floor_plan_templates + updated_elevation_templates  + updated_rcp_templates



# Step 5: Match each row's data to a view template and levels
matching_rows = []  # To keep track of rows that matched templates and levels
selected_templates = []  # To store the selected template for each row
template_suggestions = []  # To keep track of suggestions for unmatched templates
level_mismatch = []

for row in combined_data:
        

    col_a, col_b, col_c, col_d, col_e, col_f, col_g, col_h, col_i, col_j, col_k, col_l,col_m, col_n,col_o, col_p, col_q, col_r, col_s, col_t, split_discipline = row

    # Find the corresponding level in Revit that matches the column K data
    level_name = col_n
    matching_level = []
    level_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

    for level in level_collector:

        if level.Name.lower() == level_name.lower():
            matching_level.append(level)
        else:
            level_mismatch.append((level_name, "No Matching Level In Document"))

    # Use a case-insensitive check to ensure matches
    matching_template = next(
        (vt for vt in all_templates if 
         split_discipline.lower() in vt.Name.lower() and
         col_l.lower() in vt.Name.lower() and
         col_m.lower() in vt.Name.lower()),
        None
    )

    if level_mismatch:
        matching_view = next((v for v in FilteredElementCollector(doc).OfClass(View).ToElements()if v.Name == col_o), None)
        if matching_view:
            # Append row data along with the view and no level
            matching_rows.append((row, matching_template, None, matching_view))
            
    # Store row only if both template and level are matched
    if matching_template is not None and matching_level:
        matching_rows.append((row, matching_template, matching_level[0], None))  # Append row data, matching template, and level
        selected_templates.append(matching_template)  # Store the selected template
    else:
        selected_templates.append(None)  # Append None if no matching template is found
        # Add a suggestion for unmatched templates based on column data
        suggestion = "{}_{}_{}".format(split_discipline, col_l, col_m)  # Example naming convention for suggestions
        template_suggestions.append(suggestion)



# Filter combined data to keep only those rows that had a matching template
filtered_combined_data = [row for row, template in zip(combined_data, selected_templates) if template is not None]

templates_to_create = []
# Print unique suggestions for unmatched templates
if template_suggestions:
    unique_suggestions = list(set(template_suggestions))
    templates_to_create.append(unique_suggestions)

# Step 6: Create views on matching levels within a transaction

# ALL VIEW TYPES

view_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()

view_types_rcp = [vt for vt in view_types if vt.ViewFamily == ViewFamily.CeilingPlan]
if not view_types_rcp:
    TaskDialog.Show("Error", "No RCP View Types available.")
    script.exit()

rcp_type = view_types_rcp[0]


view_types_plans = [vt for vt in view_types if vt.ViewFamily == ViewFamily.FloorPlan]
if not view_types_plans:
    TaskDialog.Show("Error", "No Floor Plan View Types available.")
    script.exit()

floor_plan_type = view_types_plans[0]


tg = TransactionGroup(doc, "Create Floor Plans")
tg.Start()

counter = 1
new_names = []
plan_placed_on_sheet = []
skipped_views = []
existing_views = FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType()
    # Collect all views
all_the_views = FilteredElementCollector(doc) \
    .OfClass(View) \
    .WhereElementIsNotElementType() \
    .ToElements()

# Filter only elevation views
plan_views = [view for view in all_the_views if view.ViewType == ViewType.FloorPlan]
elevation_views = [view for view in all_the_views if view.ViewType == ViewType.Elevation]
section_views = [view for view in all_the_views if view.ViewType == ViewType.Section]
scope_box_views = elevation_views

existing_view_names = {v.Name for v in FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType()}

existing_sheets = {s.Name for s in FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType()}


# Filter out views that are already placed on sheets
placed_view_ids = {
    vp.ViewId for vp in FilteredElementCollector(doc).OfClass(Viewport)
}
placed_views = [
    view for view in elevation_views if view.Id in placed_view_ids
]
unplaced_views = [
    view for view in elevation_views if view.Id not in placed_view_ids
]

all_scopeboxes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).ToElements()

# Place each created view on its corresponding sheet
for row_data, matching_template, matching_level, matching_view in matching_rows:

    matched_views = []
    col_a, col_b, col_c, col_d, col_e, col_f, col_g, col_h,col_i,  col_j, col_k, col_l, col_m, col_n, col_o, col_p, col_q, col_r, col_s, col_t, split_discipline = row_data



    if col_l.lower() in ["elevation"]:

        t = Transaction(doc, "place elevations")
        t.Start()

        # Filter out views that are already placed on sheets
        placed_view_ids = {
            vp.ViewId for vp in FilteredElementCollector(doc).OfClass(Viewport)
        }
        placed_views = [
            view for view in elevation_views if view.Id in placed_view_ids
        ]
        unplaced_views = [
            view for view in elevation_views if view.Id not in placed_view_ids
        ]

        elevation_scope_box_views = []
        matched_elevation_scope_box = []
        view_names = [name.strip() for name in col_o.split(",")]
        scopeboxname_excel = [name.strip() for name in col_p.split(",")]

        for scopebox in all_scopeboxes:
                for sb_name in scopeboxname_excel:
                    if sb_name == scopebox.Name:
                        matched_elevation_scope_box.append(scopebox)
                    
        for view_name in view_names:
            for view in unplaced_views:
                if view.Name.lower()== view_name.lower():
                    matched_views.append(view)

        for view_name in view_names:
            for view in elevation_views:
                if view.Name.lower()== view_name.lower():
                    elevation_scope_box_views.append(view)

        for new_view in matched_views:
            if new_view:
                if matching_template:
                    primary_view_id = new_view.GetPrimaryViewId()
                    if primary_view_id != ElementId.InvalidElementId:
                        primary_view = doc.GetElement(primary_view_id)
                        primary_view.ViewTemplateId = matching_template.Id
                    else:
                        new_view.ViewTemplateId = matching_template.Id
            else:
                new_view = None

        if col_p: 
            if len(elevation_scope_box_views) == len(matched_elevation_scope_box):  # Ensure all lists have matching lengths
                for i, view in enumerate(elevation_scope_box_views):
                    try:
                        sb = matched_elevation_scope_box[i]  # Access `scopy` safely
                        view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(sb.Id)
                    except IndexError:
                        print("Index out of range for list of scope box:", i)
            elif len(matched_elevation_scope_box) == 1:

                try:
                    for view in elevation_scope_box_views:
                        for sb in matched_elevation_scope_box:
                            view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(sb.Id)
                except IndexError:
                        print("Index out of range for list of scope box:", i)
            else:
                print("List Length Mismatch for Scope Box and Views")

        else:
            # If col_p is empty, remove the assigned scope box
            for view in elevation_scope_box_views:
                try:
                    scope_box_param = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
                    if scope_box_param:
                        scope_box_param.Set(ElementId.InvalidElementId)  # Use this to clear the scope box
                except Exception as e:
                    print("Error removing scope box from view {}: {}".format(view.Name, e))


        # Place the view on the corresponding sheet
        sheet_name = col_k
        sheet_number = col_j


        # Check if there’s an existing sheet with the same number but a different discipline
        conflicting_sheet = next((s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
                                if s.SheetNumber == sheet_number and 
                                s.LookupParameter("Sub-Discipline").AsString() != selected_discipline), None)

        if conflicting_sheet:
            t.RollBack()
            TaskDialog.Show("Error", "A sheet with the same sheet number exists under a different discipline. Please use a unique sheet number or discipline.")
            script.exit()

        else:

            existing_sheet = next((s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
                        if s.Name == sheet_name and s.SheetNumber == sheet_number and s.LookupParameter("Sub-Discipline").AsString() == selected_discipline), None)

            if existing_sheet:
                sheet = existing_sheet

            else:
                # Create a new sheet if it doesn't exist
                sheet = ViewSheet.Create(doc, selected_title_block_type.Id)
                sheet.Name = sheet_name
                sheet.SheetNumber = sheet_number
                existing_sheets.add(sheet_name)
                if new_view:
                    new_names.append((new_view.Name, sheet.Name))

            if sheet:
                # Initialize parameter values
                parameter_map = {
                    0: col_a, 1: col_b, 2: col_c, 3: col_d, 4: col_e,
                    5: col_f, 6: col_g, 7: col_h, 8: col_i}

                # Set sheet parameters if available
                folder_parameter = sheet.LookupParameter("Folder")
                if folder_parameter and folder_parameter.StorageType == StorageType.String:
                    folder_parameter.Set("Elevation")
                subdiscipline_parameter = sheet.LookupParameter("Sub-Discipline")
                if subdiscipline_parameter and folder_parameter.StorageType == StorageType.String:
                    subdiscipline_parameter.Set(selected_discipline)

                for idx, param_value in parameter_map.items():
                    parameter_name = excel_worksheet.cell_value(0, idx).strip()
                    if parameter_name:
                        sheet_param = sheet.LookupParameter(parameter_name)
                        if sheet_param and sheet_param.StorageType == StorageType.String:
                            sheet_param.Set(param_value)

                # Set additional sheet parameters
                for param_name, param_value in [("Drawn By", col_q), ("Designed By", col_r), ("Checked By", col_s), ("Approved By", col_t)]:
                    param = sheet.LookupParameter(param_name)
                    if param and param.StorageType == StorageType.String:
                        param.Set(param_value)  

                existing_sheets.add(sheet.Name)

        t.Commit()

        # Position the view on the sheet
        title_block_bb = selected_title_block_type.get_BoundingBox(None)
        sheet_center = XYZ((title_block_bb.Max.X + title_block_bb.Min.X) / 2- 0.18209,(title_block_bb.Max.Y + title_block_bb.Min.Y) / 2, 0)
        offset = 0.2
        
        if sheet:
            for index, new_view in enumerate(matched_views):
                view_position = XYZ(sheet_center.X, sheet_center.Y - (index*offset), 0)
                try:

                    t=Transaction(doc, "place views")
                    t.Start()
                    Viewport.Create(doc, sheet.Id, new_view.Id, view_position)

                    t.Commit()

                except Exception as e:
                    print("Failed to place view {} on sheet {}: {}".format(new_view.Name, sheet_name, str(e)))

                    # Record the end time and runtime
                    end_time = time.time()
                    runtime = end_time - start_time

                    # Log the error details
                    error_occured = ("Failed to place view {} on sheet {}: {}".format(new_view.Name, sheet_name, str(e)))
                    run_result = "Error"
                    element_count = 10

                    # Function to log run data in case of error
                    get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)
                    
    
        else:
            skipped_views.append((sheet.SheetNumber, sheet.Name) )

    elif col_l.lower() in ["section"]:
        
        t = Transaction(doc, "place sections")
        t.Start()

        new_view = None
        # Filter out views that are already placed on sheets
        placed_view_ids = {
            vp.ViewId for vp in FilteredElementCollector(doc).OfClass(Viewport)
        }
        placed_views = [
            view for view in section_views if view.Id in placed_view_ids
        ]
        unplaced_views = [
            view for view in section_views if view.Id not in placed_view_ids
        ]

        section_scope_box_views = []
        matched_section_scope_box = []
        view_names = [name.strip() for name in col_o.split(",")]
        scopeboxname_excel = [name.strip() for name in col_p.split(",")]

        for scopebox in all_scopeboxes:
                for sb_name in scopeboxname_excel:
                    if sb_name == scopebox.Name:
                        matched_section_scope_box.append(scopebox)
                    
        for view_name in view_names:
            for view in unplaced_views:
                if view.Name.lower()== view_name.lower():
                    matched_views.append(view)

        for view_name in view_names:
            for view in section_views:
                if view.Name.lower()== view_name.lower():
                    section_scope_box_views.append(view)


        for new_view in matched_views:
            if new_view:
                if matching_template:
                    primary_view_id = new_view.GetPrimaryViewId()
                    if primary_view_id != ElementId.InvalidElementId:
                        primary_view = doc.GetElement(primary_view_id)
                        primary_view.ViewTemplateId = matching_template.Id
                    else:
                        new_view.ViewTemplateId = matching_template.Id
            else:
                new_view = None


        if col_p: 

            if len(section_scope_box_views) == len(matched_section_scope_box):  # Ensure all lists have matching lengths
                for i, view in enumerate(section_scope_box_views):
                    try:
                        sb = matched_section_scope_box[i]  # Access `scopy` safely
                        view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(sb.Id)
                    except IndexError:
                        print("Index out of range for list of scope box:", i)
            elif len(matched_section_scope_box) == 1:
                try:
                    for view in section_scope_box_views:
                        for sb in matched_section_scope_box:
                            view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(sb.Id)
                except IndexError:
                        print("Index out of range for list of scope box:", i)
            else:
                print("List Length Mismatch for Scope Box and Views")

        else:
            # If col_p is empty, remove the assigned scope box
            for view in section_scope_box_views:
                try:
                    scope_box_param = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
                    if scope_box_param:
                        scope_box_param.Set(ElementId.InvalidElementId)  # Use this to clear the scope box
                except Exception as e:
                    print("Error removing scope box from view {}: {}".format(view.Name, e))


        # Place the view on the corresponding sheet
        sheet_name = col_k
        sheet_number = col_j


        # Check if there’s an existing sheet with the same number but a different discipline
        conflicting_sheet = next((s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
                                if s.SheetNumber == sheet_number and 
                                s.LookupParameter("Sub-Discipline").AsString() != selected_discipline), None)

        if conflicting_sheet:
            t.RollBack()
            TaskDialog.Show("Error", "A sheet with the same sheet number exists under a different discipline. Please use a unique sheet number or discipline.")
            script.exit()

        else:

            existing_sheet = next((s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
                        if s.Name == sheet_name and s.SheetNumber == sheet_number and s.LookupParameter("Sub-Discipline").AsString() == selected_discipline), None)

            if existing_sheet:
                sheet = existing_sheet

            else:
                # Create a new sheet if it doesn't exist
                sheet = ViewSheet.Create(doc, selected_title_block_type.Id)
                sheet.Name = sheet_name
                sheet.SheetNumber = sheet_number
                existing_sheets.add(sheet_name)
                if new_view:
                    new_names.append((new_view.Name, sheet.Name))

            if sheet:
                # Initialize parameter values
                parameter_map = {
                    0: col_a, 1: col_b, 2: col_c, 3: col_d, 4: col_e,
                    5: col_f, 6: col_g, 7: col_h, 8: col_i}

                # Set sheet parameters if available
                folder_parameter = sheet.LookupParameter("Folder")
                if folder_parameter and folder_parameter.StorageType == StorageType.String:
                    folder_parameter.Set("Section")
                subdiscipline_parameter = sheet.LookupParameter("Sub-Discipline")
                if subdiscipline_parameter and folder_parameter.StorageType == StorageType.String:
                    subdiscipline_parameter.Set(selected_discipline)

                for idx, param_value in parameter_map.items():
                    parameter_name = excel_worksheet.cell_value(0, idx).strip()
                    if parameter_name:
                        sheet_param = sheet.LookupParameter(parameter_name)
                        if sheet_param and sheet_param.StorageType == StorageType.String:
                            sheet_param.Set(param_value)

                # Set additional sheet parameters
                for param_name, param_value in [("Drawn By", col_q), ("Designed By", col_r), ("Checked By", col_s), ("Approved By", col_t)]:
                    param = sheet.LookupParameter(param_name)
                    if param and param.StorageType == StorageType.String:
                        param.Set(param_value)  

                existing_sheets.add(sheet.Name)

        t.Commit()

        # Position the view on the sheet
        title_block_bb = selected_title_block_type.get_BoundingBox(None)
        sheet_center = XYZ((title_block_bb.Max.X + title_block_bb.Min.X) / 2- 0.18209,(title_block_bb.Max.Y + title_block_bb.Min.Y) / 2, 0)
        offset = 0.2
        
        if sheet:
            for index, new_view in enumerate(matched_views):
                view_position = XYZ(sheet_center.X, sheet_center.Y - (index*offset), 0)
                try:

                    t=Transaction(doc, "place views")
                    t.Start()
                    Viewport.Create(doc, sheet.Id, new_view.Id, view_position)

                    t.Commit()

                except Exception as e:
                    print("Failed to place view {} on sheet {}: {}".format(new_view.Name, sheet_name, str(e)))

                    # Record the end time and runtime
                    end_time = time.time()
                    runtime = end_time - start_time

                    # Log the error details
                    error_occured = ("Failed to place view {} on sheet {}: {}".format(new_view.Name, sheet_name, str(e)))
                    run_result = "Error"
                    element_count = 10

                    # Function to log run data in case of error
                    get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)
                    
    
        else:
            skipped_views.append((sheet.SheetNumber, sheet.Name) )

    elif col_l.lower() in ["structural framing plan"]:

        if matching_level:
            t = Transaction(doc, "create framing plan")
            t.Start()
            # Retrieve the ViewFamilyType for Reflected Ceiling Plans
            rcp_type = None
            for vft in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
                if vft.ViewFamily == DB.ViewFamily.CeilingPlan:
                    rcp_type = vft
                    break

            if not rcp_type:
                raise Exception("No Ceiling Plan ViewFamilyType found in the project.")
            
            level = matching_level[0]
            view_name = "{}_{}".format(col_l,level.Name)
            # Check if a view with the same name exists and meets the criteria

            new_view = None
            unique_view_name = view_name


            if new_view is None:
                counter = 1
                while unique_view_name in existing_view_names:
                    unique_view_name = "{}_{}".format(view_name, counter)
                    counter += 1

                # Create the floor plan view and set name
                new_view = ViewPlan.Create(doc, rcp_type.Id, level.Id)
                new_view.Name = unique_view_name
                existing_view_names.add(new_view.Name)

            # Assign the view template
            if matching_template:
                new_view.ViewTemplateId = matching_template.Id

            # if col_p:
            #     scope_boxes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).ToElements()
            #     scope_box = None
            #     for sb in scope_boxes:
            #         sb_name = sb.Name
            #         if sb_name == (col_p):
            #             scope_box = sb
            #             break

            #         if scope_box is not None:      
            #             new_view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(scope_box.Id)
            # else:
            #     scope_box_param = new_view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
            #     if scope_box_param:
            #         scope_box_param.Set(None)  # Remove scope box


            # Place the view on the corresponding sheet
            sheet_name = col_k
            sheet_number = col_j


            # Check if there’s an existing sheet with the same number but a different discipline
            conflicting_sheet = next((s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
                                    if s.SheetNumber == sheet_number and 
                                    s.LookupParameter("Sub-Discipline").AsString() != selected_discipline), None)

            if conflicting_sheet:
                t.RollBack()
                TaskDialog.Show("Error", "A sheet with the same sheet number exists under a different discipline. Please use a unique sheet number or discipline.")
                script.exit()

            else:

                existing_sheet = next((s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
                            if s.Name == sheet_name and s.SheetNumber == sheet_number and s.LookupParameter("Sub-Discipline").AsString() == selected_discipline), None)

                if existing_sheet:
                    sheet = existing_sheet

                else:
                    # Create a new sheet if it doesn't exist
                    sheet = ViewSheet.Create(doc, selected_title_block_type.Id)
                    sheet.Name = sheet_name
                    sheet.SheetNumber = sheet_number
                    existing_sheets.add(sheet_name)
                    if new_view:
                        new_names.append((new_view.Name, sheet.Name))


                # Set sheet parameters if available
                folder_parameter = sheet.LookupParameter("Folder")
                if folder_parameter and folder_parameter.StorageType == StorageType.String:
                    folder_parameter.Set("Floor Plan")
                subdiscipline_parameter = sheet.LookupParameter("Sub-Discipline")
                if subdiscipline_parameter and folder_parameter.StorageType == StorageType.String:
                    subdiscipline_parameter.Set(selected_discipline)

                a_value = excel_worksheet.cell_value(0, 0)  
                if a_value:
                    a_parameter = sheet.LookupParameter(a_value)
                    if a_parameter and a_parameter.StorageType == StorageType.String:
                        a_parameter.Set(col_a)
                    else:
                        print("No parameter found for {}".format(a_value))


                b_value = excel_worksheet.cell_value(0, 1).strip()
                if b_value:
                    b_parameter = sheet.LookupParameter(b_value)
                    if b_parameter and b_parameter.StorageType == StorageType.String:
                        b_parameter.Set(col_b)
                    else:
                        print("No parameter found for {}".format(b_value))

                c_value = excel_worksheet.cell_value(0, 2).strip()
                if c_value:
                    c_parameter = sheet.LookupParameter(c_value)
                    if c_parameter and c_parameter.StorageType == StorageType.String:
                        c_parameter.Set(col_c)
                    else:
                        print("No parameter found for {}".format(c_value))



                d_value = excel_worksheet.cell_value(0, 3).strip()
                if d_value:
                    d_parameter = sheet.LookupParameter(d_value)
                    if d_parameter and d_parameter.StorageType == StorageType.String:
                        d_parameter.Set(col_d)
                    else:
                        print("No parameter found for {}".format(d_value))


                e_value = excel_worksheet.cell_value(0, 4).strip()
                if e_value:
                    e_parameter = sheet.LookupParameter(e_value)
                    if e_parameter and e_parameter.StorageType == StorageType.String:
                        e_parameter.Set(col_e)
                    else:
                        print("No parameter found for {}".format(e_value))


                f_value = excel_worksheet.cell_value(0, 5).strip()
                if f_value:
                    f_parameter = sheet.LookupParameter(f_value)
                    if f_parameter and f_parameter.StorageType == StorageType.String:
                        f_parameter.Set(col_f)
                    else:
                        print("No parameter found for {}".format(f_value))


                g_value = excel_worksheet.cell_value(0, 6).strip()
                if g_value:
    
                    g_parameter = sheet.LookupParameter(g_value)

                    if g_parameter and g_parameter.StorageType == StorageType.String:
                        g_parameter.Set(col_g)
                    else:
                        print("No parameter found for {}".format(g_value))


                h_value = excel_worksheet.cell_value(0, 7).strip()
                if h_value:
                    h_parameter = sheet.LookupParameter(h_value)
                    if h_parameter and h_parameter.StorageType == StorageType.String:
                        h_parameter.Set(col_h)
                    else:
                        print("No parameter found for {}".format(h_value))



                i_value = excel_worksheet.cell_value(0, 8).strip()
                if i_value:
                    i_parameter = sheet.LookupParameter(i_value)
                    if i_parameter and i_parameter.StorageType == StorageType.String:
                        i_parameter.Set(col_i)
                    else:
                        print("No parameter found for {}".format(i_value))

                drawn_parameter = sheet.LookupParameter("Drawn By")
                if drawn_parameter and drawn_parameter.StorageType == StorageType.String:
                    drawn_parameter.Set(col_q)
                designed_parameter = sheet.LookupParameter("Designed By")
                if designed_parameter and designed_parameter.StorageType == StorageType.String:
                    designed_parameter.Set(col_r)
                checked_parameter = sheet.LookupParameter("Checked By")
                if checked_parameter and checked_parameter.StorageType == StorageType.String:
                    checked_parameter.Set(col_s)
                approved_parameter = sheet.LookupParameter("Approved By")
                if approved_parameter and approved_parameter.StorageType == StorageType.String:
                    approved_parameter.Set(col_t)

                existing_sheets.add(sheet.Name)

            t.Commit()

            # Position the view on the sheet
            title_block_bb = selected_title_block_type.get_BoundingBox(None)
            sheet_center = XYZ((title_block_bb.Max.X + title_block_bb.Min.X) / 2- 0.18209,(title_block_bb.Max.Y + title_block_bb.Min.Y) / 2, 0)
            
            if sheet:

                # Check if the sheet already has views placed on it
                viewport_collector = FilteredElementCollector(doc, sheet.Id).OfClass(Viewport)
                if viewport_collector.GetElementCount() == 0:
                    try:
                        t=Transaction(doc, "place views")
                        t.Start()
                        Viewport.Create(doc, sheet.Id, new_view.Id, sheet_center)

                        tb_center = (title_block_bb.Min + title_block_bb.Max) / 2

                        # Get the viewport on the sheet
                        viewports = list(FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_Viewports))
                        if not viewports:
                            print("No viewport found on the sheet.")

                        t.Commit()

                    except Exception as e:
                        print("Failed to place view {} on sheet {}: {}".format(new_view.Name, sheet_name, str(e)))

                        # Record the end time and runtime
                        end_time = time.time()
                        runtime = end_time - start_time

                        # Log the error details
                        error_occured = ("Failed to place view {} on sheet {}: {}".format(new_view.Name, sheet_name, str(e)))
                        run_result = "Error"
                        element_count = 10

                        # Function to log run data in case of error
                        get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)
                        
                        
                else:
                        skipped_views.append((sheet.SheetNumber, sheet.Name) )

    # Create a view name using the level information
    elif matching_level:
            t = Transaction(doc, "create plans")
            t.Start()

            level = matching_level

            view_name = "{}_{}".format(col_l,level.Name)

            # Check if a view with the same name exists and meets the criteria

            new_view = None
            unique_view_name = view_name


            if new_view is None:
                counter = 1
                while unique_view_name in existing_view_names:
                    unique_view_name = "{}_{}".format(view_name, counter)
                    counter += 1

                # Create the floor plan view and set name
                new_view = ViewPlan.Create(doc, floor_plan_type.Id, level.Id)
                new_view.Name = unique_view_name
                existing_view_names.add(new_view.Name)

            # Assign the view template
            if matching_template:
                new_view.ViewTemplateId = matching_template.Id

            if col_p:
                scope_boxes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).ToElements()
                scope_box = None
                for sb in scope_boxes:
                    sb_name = sb.Name
                    if sb_name == (col_p):
                        scope_box = sb
                        break

                if scope_box is not None:      
                    new_view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(scope_box.Id)
            else:
                scope_box_param = new_view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
                if scope_box_param:
                    scope_box_param.Set(None)  # Remove scope box

            # Place the view on the corresponding sheet
            sheet_name = col_k
            sheet_number = col_j


            # Check if there’s an existing sheet with the same number but a different discipline
            conflicting_sheet = next((s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
                                    if s.SheetNumber == sheet_number and 
                                    s.LookupParameter("Sub-Discipline").AsString() != selected_discipline), None)

            if conflicting_sheet:
                t.RollBack()
                TaskDialog.Show("Error", "A sheet with the same sheet number exists under a different discipline. Please use a unique sheet number or discipline.")
                script.exit()

            else:

                existing_sheet = next((s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
                            if s.Name == sheet_name and s.SheetNumber == sheet_number and s.LookupParameter("Sub-Discipline").AsString() == selected_discipline), None)

                if existing_sheet:
                    sheet = existing_sheet

                else:
                    # Create a new sheet if it doesn't exist
                    sheet = ViewSheet.Create(doc, selected_title_block_type.Id)
                    sheet.Name = sheet_name
                    sheet.SheetNumber = sheet_number
                    existing_sheets.add(sheet_name)
                    if new_view:
                        new_names.append((new_view.Name, sheet.Name))

                if sheet:
                    # Initialize parameter values
                    parameter_map = {
                        0: col_a, 1: col_b, 2: col_c, 3: col_d, 4: col_e,
                        5: col_f, 6: col_g, 7: col_h, 8: col_i}

                    # Set sheet parameters if available
                    folder_parameter = sheet.LookupParameter("Folder")
                    if folder_parameter and folder_parameter.StorageType == StorageType.String:
                        folder_parameter.Set("Floor Plan")
                    subdiscipline_parameter = sheet.LookupParameter("Sub-Discipline")
                    if subdiscipline_parameter and folder_parameter.StorageType == StorageType.String:
                        subdiscipline_parameter.Set(selected_discipline)

                    for idx, param_value in parameter_map.items():
                        parameter_name = excel_worksheet.cell_value(0, idx).strip()
                        if parameter_name:
                            sheet_param = sheet.LookupParameter(parameter_name)
                            if sheet_param and sheet_param.StorageType == StorageType.String:
                                sheet_param.Set(param_value)

                    # Set additional sheet parameters
                    for param_name, param_value in [("Drawn By", col_q), ("Designed By", col_r), ("Checked By", col_s), ("Approved By", col_t)]:
                        param = sheet.LookupParameter(param_name)
                        if param and param.StorageType == StorageType.String:
                            param.Set(param_value)  

                    existing_sheets.add(sheet.Name)

            t.Commit()

            # Position the view on the sheet
            title_block_bb = selected_title_block_type.get_BoundingBox(None)
            sheet_center = XYZ((title_block_bb.Max.X + title_block_bb.Min.X) / 2- 0.18209,(title_block_bb.Max.Y + title_block_bb.Min.Y) / 2, 0)
            
            if sheet:

                # Check if the sheet already has views placed on it
                viewport_collector = FilteredElementCollector(doc, sheet.Id).OfClass(Viewport)
                if viewport_collector.GetElementCount() == 0:
                    try:
                        t=Transaction(doc, "place views")
                        t.Start()
                        Viewport.Create(doc, sheet.Id, new_view.Id, sheet_center)

                        tb_center = (title_block_bb.Min + title_block_bb.Max) / 2

                        # Get the viewport on the sheet
                        viewports = list(FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_Viewports))
                        if not viewports:
                            print("No viewport found on the sheet.")

                        t.Commit()

                    except Exception as e:
                        print("Failed to place view {} on sheet {}: {}".format(new_view.Name, sheet_name, str(e)))

                        # Record the end time and runtime
                        end_time = time.time()
                        runtime = end_time - start_time

                        # Log the error details
                        error_occured = ("Failed to place view {} on sheet {}: {}".format(new_view.Name, sheet_name, str(e)))
                        run_result = "Error"
                        element_count = 10

                        # Function to log run data in case of error
                        get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)
                        
                        
                else:

                    t=Transaction(doc, "place views")
                    t.Start()
                    doc.Delete(new_view.Id)
                    t.Commit()

                    skipped_views.append((sheet.SheetNumber, sheet.Name) )
    
    else:
        if matching_view:
            t = Transaction(doc, "place plans")
            t.Start()

            # Filter out views that are already placed on sheets
            placed_view_ids = {
                vp.ViewId for vp in FilteredElementCollector(doc).OfClass(Viewport)
            }
            placed_views = [
                view for view in plan_views if view.Id in placed_view_ids
            ]
            unplaced_views = [
                view for view in plan_views if view.Id not in placed_view_ids
            ]

            plan_scope_box_views = []
            matched_plan_scope_box = []
            view_names = [name.strip() for name in col_o.split(",")]
            scopeboxname_excel = [name.strip() for name in col_p.split(",")]

            for scopebox in all_scopeboxes:
                    for sb_name in scopeboxname_excel:
                        if sb_name == scopebox.Name:
                            matched_plan_scope_box.append(scopebox)
  
            for view_name in view_names:
                for view in unplaced_views:
                    if view.Name.lower()== view_name.lower():
                        matched_views.append(view)

            for view_name in view_names:
                for view in plan_views:
                    if view.Name.lower()== view_name.lower():
                        plan_scope_box_views.append(view)

            for new_view in matched_views:
                if new_view:
                    if matching_template:
                        primary_view_id = new_view.GetPrimaryViewId()
                        if primary_view_id != ElementId.InvalidElementId:
                            primary_view = doc.GetElement(primary_view_id)
                            primary_view.ViewTemplateId = matching_template.Id
                        else:
                            new_view.ViewTemplateId = matching_template.Id
                else:
                    new_view = None


            if col_p: 
                if len(plan_scope_box_views) == len(matched_plan_scope_box):  # Ensure all lists have matching lengths
                    for i, view in enumerate(plan_scope_box_views):
                        try:
                            sb = matched_plan_scope_box[i]  # Access `scopy` safely
                            view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(sb.Id)
                        except IndexError:
                            print("Index out of range for list of scope box:", i)
                elif len(matched_plan_scope_box) == 1:

                    try:
                        for view in plan_scope_box_views:
                            for sb in matched_plan_scope_box:
                                view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(sb.Id)
                    except IndexError:
                            print("Index out of range for list of scope box:", i)
                else:
                    print("List Length Mismatch for Scope Box and Views")

            else:
                # If col_p is empty, remove the assigned scope box
                for view in plan_scope_box_views:
                    try:
                        scope_box_param = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
                        if scope_box_param:
                            scope_box_param.Set(ElementId.InvalidElementId)  # Use this to clear the scope box
                    except Exception as e:
                        print("Error removing scope box from view {}: {}".format(view.Name, e))


            # Place the view on the corresponding sheet
            sheet_name = col_k
            sheet_number = col_j


            # Check if there’s an existing sheet with the same number but a different discipline
            conflicting_sheet = next((s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
                                    if s.SheetNumber == sheet_number and 
                                    s.LookupParameter("Sub-Discipline").AsString() != selected_discipline), None)

            if conflicting_sheet:
                t.RollBack()
                TaskDialog.Show("Error", "A sheet with the same sheet number exists under a different discipline. Please use a unique sheet number or discipline.")
                script.exit()

            else:

                existing_sheet = next((s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
                            if s.Name == sheet_name and s.SheetNumber == sheet_number and s.LookupParameter("Sub-Discipline").AsString() == selected_discipline), None)

                if existing_sheet:
                    sheet = existing_sheet


                else:
                    # Create a new sheet if it doesn't exist
                    sheet = ViewSheet.Create(doc, selected_title_block_type.Id)
                    sheet.Name = sheet_name
                    sheet.SheetNumber = sheet_number
                    existing_sheets.add(sheet_name)
                    if new_view:
                        new_names.append((new_view.Name, sheet.Name))
                        plan_placed_on_sheet.append((new_view.Name, sheet.Name))

                if sheet:
                    # Initialize parameter values
                    parameter_map = {
                        0: col_a, 1: col_b, 2: col_c, 3: col_d, 4: col_e,
                        5: col_f, 6: col_g, 7: col_h, 8: col_i}

                    # Set sheet parameters if available
                    folder_parameter = sheet.LookupParameter("Folder")
                    if folder_parameter and folder_parameter.StorageType == StorageType.String:
                        folder_parameter.Set("Floor Plan")
                    subdiscipline_parameter = sheet.LookupParameter("Sub-Discipline")
                    if subdiscipline_parameter and folder_parameter.StorageType == StorageType.String:
                        subdiscipline_parameter.Set(selected_discipline)

                    for idx, param_value in parameter_map.items():
                        parameter_name = excel_worksheet.cell_value(0, idx).strip()
                        if parameter_name:
                            sheet_param = sheet.LookupParameter(parameter_name)
                            if sheet_param and sheet_param.StorageType == StorageType.String:
                                sheet_param.Set(param_value)

                    # Set additional sheet parameters
                    for param_name, param_value in [("Drawn By", col_q), ("Designed By", col_r), ("Checked By", col_s), ("Approved By", col_t)]:
                        param = sheet.LookupParameter(param_name)
                        if param and param.StorageType == StorageType.String:
                            param.Set(param_value)  

                    existing_sheets.add(sheet.Name)

            t.Commit()

            # Position the view on the sheet
            title_block_bb = selected_title_block_type.get_BoundingBox(None)
            sheet_center = XYZ((title_block_bb.Max.X + title_block_bb.Min.X) / 2- 0.18209,(title_block_bb.Max.Y + title_block_bb.Min.Y) / 2, 0)
            offset = 0.2
            
            if sheet:
                for index, new_view in enumerate(matched_views):
                    view_position = XYZ(sheet_center.X, sheet_center.Y - (index*offset), 0)
                    try:

                        t=Transaction(doc, "place views")
                        t.Start()
                        Viewport.Create(doc, sheet.Id, new_view.Id, view_position)

                        t.Commit()

                    except Exception as e:
                        print("Failed to place view {} on sheet {}: {}".format(new_view.Name, sheet_name, str(e)))

                        # Record the end time and runtime
                        end_time = time.time()
                        runtime = end_time - start_time

                        # Log the error details
                        error_occured = ("Failed to place view {} on sheet {}: {}".format(new_view.Name, sheet_name, str(e)))
                        run_result = "Error"
                        element_count = 10

                        # Function to log run data in case of error
                        get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)
                        
        
            else:
                skipped_views.append((sheet.SheetNumber, sheet.Name) )


tg.Assimilate()

# Record the end time
end_time = time.time()
runtime = end_time - start_time

run_result = "Tool ran successfully"
element_count = counter if counter else 0
error_occured = "Nil"
get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)

if templates_to_create or level_mismatch or new_names or skipped_views:
    output.print_md(header_data)


if skipped_views:
    output.print_md("## ⚠️ Skipping View Placement")  # Markdown Heading 2 
    output.print_md("---")  # Markdown Line Break
    output.print_md("❌ Sheet Already Occupied, Skipping View Placement")
    output.print_table(table_data=skipped_views, columns=["SHEET NUMBER", "SHEET NAME"])  # Print a Table


if new_names:
    #Print the header for skipped views
    output.print_md("## VIEWS AND SHEETS CREATED 😊")  # Markdown Heading 2 
    output.print_md("---")  # Markdown Line Break
    output.print_md(" New View Created and placed on Newly created Sheet.")  # Print a Line
    # Create a table to display the skipped views
    output.print_table(table_data=new_names, columns=["VIEW NAME", "SHEET NAME"])  # Print a Table

if templates_to_create:
    # Flatten templates_to_create if it contains nested lists
    if any(isinstance(item, list) for item in templates_to_create):
        templates_to_create = [subitem for item in templates_to_create for subitem in (item if isinstance(item, list) else [item])]

    # Collect all views in the model
    views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType()

    # Get the names of all existing template views
    existing_template_names = [view.Name.strip().lower() for view in views if view.IsTemplate]

    # Remove templates already present in the model from templates_to_create
    templates_to_actually_create = [
        template for template in templates_to_create
        if template.strip().lower() not in existing_template_names
    ]

    if templates_to_actually_create:
        # Print the header for skipped views
        output.print_md("## ⚠️ Missing Templates ☹️")  # Markdown Heading 2 
        output.print_md("---")  # Markdown Line Break
        output.print_md("❌ Views skipped due to missing templates.")  # Print a Line
        # Create a table to display the skipped views
        output.print_table(table_data=templates_to_create, columns=["MISSING TEMPLATE", "COMMENT"])  # Print a Table


if level_mismatch is not None and level in level_mismatch:
    # Print the header for skipped views
    output.print_md("## ⚠️ Level Mismatch ☹️")  # Markdown Heading 2 
    output.print_md("---")  # Markdown Line Break
    output.print_md("❌ Views skipped due to Level Mismatch.")  # Print a Line
    # Create a table to display the skipped views
    output.print_table(table_data=level_mismatch, columns=["EXCEL LEVEL NAME", "COMMENT"])  # Print a Table

    print("\n\n")
    output.print_md("---")  # Markdown Line Break


# Ensure to print or log anything else if necessary
runtime = time.time() - start_time
print("Script completed successfully in {:.2f} seconds.".format(runtime))