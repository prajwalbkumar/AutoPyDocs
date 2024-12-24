# -*- coding: utf-8 -*-
'''Align & Dim Levels'''
__title__ = " Align & Dim Levels"
__author__ = "romaramnani"

import clr
import time

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB  import *
from pyrevit            import revit, forms, script
from Autodesk.Revit.DB  import FilteredElementCollector, BuiltInCategory, View, ViewType, Viewport, XYZ
from collections        import defaultdict
from datetime           import datetime

from Extract.RunData    import get_run_data
from doc_functions      import get_view_on_sheets
from view_functions     import align_levels, get_levels_in_view
from g_curve_functions  import refArray, refLine
from dim_functions      import datum_points

import Autodesk.Revit.DB    as DB

start_time = time.time()
manual_time = 500 

doc = revit.doc
output = script.get_output()

tool_name = __title__ 
model_name = doc.Title
app = __revit__.Application  # Returns the Revit Application Object
rvt_year = "Revit " + str(app.VersionNumber)
output = script.get_output()
user_name = app.Username
header_data = " | ".join([
    datetime.now().strftime("%Y-%m-%d %H:%M:%S"), rvt_year, tool_name, model_name, user_name
])

def get_views_for_levels(doc):
    # Initialize dictionaries to categorize views
    view_dict = {
        'Elevations': [],
        'Sections': [],
        'Detail Views': []
    }

    # Create custom button options
    options = ['No borrowers, good to go!', 'Wait, need to check!']

    # Show the dialog box with custom buttons
    sync_choice = forms.alert(
        "Ensure the model is synced and all elements are released before proceeding.",
        options=options,
        exitscript=True
    )

    if sync_choice == 'Wait, need to check!':
        script.exit()  # Exit the script

    views_collector = FilteredElementCollector(doc).OfClass(View)
    viewports_collector = FilteredElementCollector(doc).OfClass(Viewport).ToElements()

    view_to_sheet_map = {vp.ViewId: vp.SheetId for vp in viewports_collector}

    # Categorize views based on grid lines and check if they are placed on a sheet
    for view in views_collector:
        if view.Id in view_to_sheet_map:
            if view.ViewType == ViewType.Elevation and not view.IsTemplate:
                grid_lines = FilteredElementCollector(doc, view.Id).OfClass(Grid).ToElements()
                if len(grid_lines) <= 1:
                    view_dict['Detail Views'].append(view)
                else:
                    view_dict['Elevations'].append(view)
            elif view.ViewType == ViewType.Section and not view.IsTemplate:
                grid_lines = FilteredElementCollector(doc, view.Id).OfClass(Grid).ToElements()
                if len(grid_lines) <= 1:
                    view_dict['Detail Views'].append(view)
                else:
                    view_dict['Sections'].append(view)

    selected_categories = forms.SelectFromList.show(
        sorted(view_dict.keys()),           
        title='Select View Categories',     
        multiselect=True                    
    )

    if not selected_categories:
        forms.alert("No view category was selected. Exiting script.", exitscript=True)

    selected_views = []
    for category in selected_categories:
        selected_views.extend(view_dict[category])

    if not selected_views:
        forms.alert("No views were found for the selected categories. Exiting script.", exitscript=True)

    # Create a dictionary for view names and their corresponding objects
    view_name_dict = {view.Name: view for view in selected_views}

    selected_view_names = forms.SelectFromList.show(
        sorted(view_name_dict.keys()),          
        title='Select Views',                   
        multiselect=True                        
    )

    if selected_view_names:
        final_selected_views = [view_name_dict[name] for name in selected_view_names]
    else:
        forms.alert("No views were selected.", exitscript=True)
    
    return final_selected_views

def offset_line_elev(revit_line, offset_distance, direction):
    start_point = revit_line.GetEndPoint(0)
    end_point = revit_line.GetEndPoint(1)
    extend_vector = direction * offset_distance

    new_start_point = XYZ(start_point.X + extend_vector.X, start_point.Y + extend_vector.Y, start_point.Z + extend_vector.Z)
    new_end_point = XYZ(end_point.X + extend_vector.X, end_point.Y + extend_vector.Y, end_point.Z + extend_vector.Z)

    # Create an offset line using the normal vector and the offset distance
    offset_line = DB.Line.CreateBound(new_start_point, new_end_point)
    return offset_line

selected_views = get_views_for_levels(doc)

align_choice = forms.alert(
    "Are your levels aligned?",
    options= ['No, please align levels and dimension them!', 'Yes, only dimension levels!'],
    exitscript=True)
if not align_choice:
    forms.alert("No selction found. Exiting script.", exitscript=True)
if align_choice == 'No, please align levels and dimension them!':
    align_levels(doc, selected_views, header_data)

total_view_count = 0
view_list_length =0
total_grid_count = 0
grid_list_length = 0

try:
    # Create transaction to create dimensions
    t = DB.Transaction(doc, "Dimension grids")
    t.Start()

    for view in selected_views:
        s_factor = view.Scale
        view_direction = view.ViewDirection.Normalize()
        offset_distance1 = -abs(4 * s_factor/100)
        offset_distance2 = -abs(5.5 * s_factor/100)
        
        levels_list = get_levels_in_view(doc, view)
        for level in levels_list:
            level.SetDatumExtentType(DB.DatumEnds.End0, view, DB.DatumExtentType.ViewSpecific)
            level.SetDatumExtentType(DB.DatumEnds.End1, view, DB.DatumExtentType.ViewSpecific)
        #print(levels_list)

        start_points, end_points, level_curves = datum_points(levels_list , view)
        #print(level_curves)

        if view_direction.IsAlmostEqualTo(XYZ(0.0, -1.0, 0.0)): 
            line = refLine(end_points)
            direction = level_curves[0].Direction
        elif view_direction.IsAlmostEqualTo(XYZ(0.0, 1.0, 0.0)):   
            line = refLine(start_points)  
            l_direction = level_curves[0].Direction
            direction = l_direction.Negate()
        elif view_direction.IsAlmostEqualTo(XYZ(1.0, 0.0, 0.0)): 
            line = refLine(end_points)
            direction = level_curves[0].Direction
        elif view_direction.IsAlmostEqualTo(XYZ(-1.0, 0.0, 0.0)):   
            line = refLine(start_points)  
            l_direction = level_curves[0].Direction
            direction = l_direction.Negate()

        #direction = level_curves[0].Direction
        if len(levels_list) >= 2:
            line_offset1 = offset_line_elev(line, offset_distance1, direction)
            n = len(levels_list ) - 1
            end_levels = [levels_list [0], levels_list [n]]
            ref_overall = refArray(end_levels)
            d1 = doc.Create.NewDimension(view, line_offset1, ref_overall)

        if len(levels_list)>2:
            ref = refArray(levels_list)   
            line_offset2 = offset_line_elev(line, offset_distance2, direction) 
            d2 = doc.Create.NewDimension(view, line_offset2, ref)

        view_list_length+= 1                   
    total_grid_count += grid_list_length
    total_view_count += view_list_length

    t.Commit()

    end_time = time.time()
    runtime = end_time - start_time
 
    run_result = "Tool ran successfully"
    if total_grid_count:
        element_count = total_grid_count
    else:
        element_count = 0

    error_occured ="Nil"

    get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)
    
except Exception as e:
    
    t.RollBack()
    #print("Error occurred: {}".format(str(e)))
    
    end_time = time.time()
    runtime = end_time - start_time

    error_occured = ("Error occurred: %s", str(e))    
    run_result = "Error"
    element_count = 0
    
    get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)

finally:
    t.Dispose()

output.print_md(header_data)

# print("Script runtime: {:.2f} seconds".format(runtime))





