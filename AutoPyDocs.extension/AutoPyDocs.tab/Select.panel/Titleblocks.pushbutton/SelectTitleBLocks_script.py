# -*- coding: utf-8 -*-  

__title__ = "Select Title Blocks in Sheets"
__author__ = "Your Name"

# Imports
import math
import os
import time
from datetime import datetime
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import WorksharingUtils
from System.Collections.Generic import List
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons

# Get the active document
doc = __revit__.ActiveUIDocument.Document

# Step 1: Get the list of all sheets in the project with both name and number
sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()

# Step 2: Extract the names and numbers of the sheets for user selection
sheet_data = ["{} - {}".format(sheet.SheetNumber, sheet.Name) for sheet in sheets]

# Step 3: Ask the user to select one or more sheets (multiselect enabled)
selected_sheet_data = forms.SelectFromList.show(sheet_data, title="Select Sheets", button_name="Select", multiselect=True)

if not selected_sheet_data:
    TaskDialog.Show("Error", "No sheets selected.")
    script.exit()

# Step 4: Filter selected sheets based on user selection
selected_sheets = [sheet for sheet in sheets if "{} - {}".format(sheet.SheetNumber, sheet.Name) in selected_sheet_data]

# Step 5: Collect only one title block from each selected sheet
title_block_ids = List[ElementId]()
for sheet in selected_sheets:
    # Filter the title blocks by category and check if they're placed in the selected sheet
    title_blocks_in_sheet = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()

    # Select the first title block found in the sheet
    if title_blocks_in_sheet:
        first_title_block = title_blocks_in_sheet[0]  # Pick the first title block
        title_block_ids.Add(first_title_block.Id)

# Step 6: Add title blocks to selection
if title_block_ids.Count > 0:
    uidoc = __revit__.ActiveUIDocument
    
    # Select the title blocks in the Revit UI
    uidoc.Selection.SetElementIds(title_block_ids)
    
    TaskDialog.Show("Success", "{} title blocks selected from the chosen sheets.".format(title_block_ids.Count))
else:
    TaskDialog.Show("Error", "No title blocks found in the selected sheets.")