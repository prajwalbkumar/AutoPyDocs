# -*- coding: utf-8 -*-     
'''Align Levels'''
__title__ = "Align Levels"
__author__ = "Abhiram Nair"


# Imports
import math
import os
import time
from datetime import datetime
from Extract.RunData import get_run_data
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import WorksharingUtils
from System.Collections.Generic import List
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons

# Record the start time
start_time = time.time()
manual_time = 45

script_dir = os.path.dirname(__file__)
ui_doc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document  # Get the Active Document
app = __revit__.Application  # Returns the Revit Application Object
rvt_year = int(app.VersionNumber)
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

# If the user selects 'Wait, need to check!', exit the script
if sync_choice == 'Wait, need to check!':
    script.exit()  # Exit the script

# Collect all views in the document
views_collector = FilteredElementCollector(doc).OfClass(View)

# Collect all viewports in the document (viewports are views placed on sheets)
viewports_collector = FilteredElementCollector(doc).OfClass(Viewport).ToElements()

# Create a dictionary mapping View IDs to Sheet IDs
view_to_sheet_map = {vp.ViewId: vp.SheetId for vp in viewports_collector}

# Categorize views based on grid lines and check if they are placed on a sheet
for view in views_collector:
    if view.Id in view_to_sheet_map:
        if view.ViewType == ViewType.Elevation and not view.IsTemplate:
            # Check for grid lines in elevation views
            grid_lines = FilteredElementCollector(doc, view.Id).OfClass(Grid).ToElements()
            if len(grid_lines) == 1:
                view_dict['Detail Views'].append(view)
            else:
                view_dict['Elevations'].append(view)
        elif view.ViewType == ViewType.Section and not view.IsTemplate:
            # Check for grid lines in section views
            grid_lines = FilteredElementCollector(doc, view.Id).OfClass(Grid).ToElements()
            if len(grid_lines) == 1:
                view_dict['Detail Views'].append(view)
            else:
                view_dict['Sections'].append(view)

# Show a dialog box for selecting desired view categories
selected_categories = forms.SelectFromList.show(
    sorted(view_dict.keys()),           # Sorted list of category names
    title='Select View Categories',     # Title of the dialog box
    multiselect=True                    # Allow multiple selections
)

if not selected_categories:
    forms.alert("No view category was selected. Exiting script.", exitscript=True)

# Collect selected views based on chosen categories
selected_views = []
for category in selected_categories:
    selected_views.extend(view_dict[category])

# Ensure at least one view was selected
if not selected_views:
    forms.alert("No views were found for the selected categories. Exiting script.", exitscript=True)

# Create a dictionary for view names and their corresponding objects
view_name_dict = {view.Name: view for view in selected_views}

# Show the selection window for the user to choose views
selected_view_names = forms.SelectFromList.show(
    sorted(view_name_dict.keys()),          # Sort and display the list of view names
    title='Select Views',                   # Title of the selection window
    multiselect=True                        # Allow multiple selections
)

if selected_view_names:
    # Get the selected views from the dictionary
    final_selected_views = [view_name_dict[name] for name in selected_view_names]
else:
    forms.alert("No views were selected.", exitscript=True)


# Set to track unique unowned element IDs
unique_unowned_element_ids = set()
# List to collect unowned elements data across all views
consolidated_unowned_element_data = []

for view in final_selected_views: 
    level_collector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
    collected_elements = level_collector
    owned_elements = []
    unowned_elements = []
    elements_to_checkout = List[ElementId]()

    for level in collected_elements:
        elements_to_checkout.Add(level.Id)

    WorksharingUtils.CheckoutElements(doc, elements_to_checkout)

    for level in collected_elements:    
        worksharingStatus = WorksharingUtils.GetCheckoutStatus(doc, level.Id)
        if not worksharingStatus == CheckoutStatus.OwnedByOtherUser:
            owned_elements.append(doc.GetElement(level.Id))
        else:
            unowned_elements.append(doc.GetElement(level.Id))

    # Collect unowned elements data only if they are not already in the unique set
    for element in unowned_elements:
        if element.Id not in unique_unowned_element_ids:
            unique_unowned_element_ids.add(element.Id)  # Add ID to the set
            try:
                consolidated_unowned_element_data.append([
                    output.linkify(element.Id), 
                    element.Category.Name.upper(), 
                    "REQUEST OWNERSHIP", 
                    WorksharingUtils.GetWorksharingTooltipInfo(doc, element.Id).Owner
                ])
            except Exception as e:
                # Handle exception if necessary
                continue

# After processing all views, print the consolidated report
if consolidated_unowned_element_data:
    output.print_md("## ⚠️ Elements Skipped ☹️")  # Markdown Heading 2
    output.print_md("---")  # Markdown Line Break
    output.print_md("❌ Make sure you have Ownership of the Elements - Request access. Refer to the **Table Report** below for reference")  # Print a Line
    output.print_table(table_data=consolidated_unowned_element_data, columns=["ELEMENT ID", "CATEGORY", "TO-DO", "CURRENT OWNER"])  # Print a Table



    # Record the end time and runtime
    end_time = time.time()
    runtime = time.time() - start_time

    # Log the error details
    error_occured = ("Error occurred: Element Unowned")
    run_result = "Error"
    element_count = 10

    # Function to log run data in case of error
    get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)

    print("Script runtime: {:.2f} seconds".format(runtime))


else:
    output.print_md("All levels are owned. No elements skipped.")  # Message when no unowned elements are found

    # Start transaction to ensure all Revit modifications happen inside it
    t = Transaction(doc, "Align Levels")

    skipped_views = []  # List to store names of skipped views

    try:

        t.Start()

        total_level_count = 0

        for view in final_selected_views:
            view_direction = view.ViewDirection.Normalize()
            right_direction = view.RightDirection

            # Determine if the view is a section or elevation
            if view.ViewType in [ViewType.Elevation, ViewType.Section]:
                # Initialize lists to store points
                crop_box_pts_x = []
                crop_box_pts_y = []
                crop_box_pts_z = []

                # Retrieve the crop region
                crop_manager = view.GetCropRegionShapeManager()
                crop_region = crop_manager.GetCropShape()

                # Explode curve loop into individual lines
                for curve_loop in crop_region:
                    for curve in curve_loop:
                        # Get XYZ coordinates of all end points
                        crop_box_pts_x.append(curve.GetEndPoint(0).X)
                        crop_box_pts_y.append(curve.GetEndPoint(0).Y)
                        crop_box_pts_z.append(curve.GetEndPoint(0).Z)

                # Sort crop box coordinates
                crop_box_x = sorted(crop_box_pts_x)
                crop_box_y = sorted(crop_box_pts_y)
                crop_box_z = sorted(crop_box_pts_z)

                min_pt = XYZ(crop_box_x[0], crop_box_y[0], crop_box_z[0])
                max_pt = XYZ(crop_box_x[-1], crop_box_y[-1], crop_box_z[-1])

                # Determine the depth direction based on the view direction

                # Case 1: (0.000000000, -1.000000000, 0.000000000) - Y- (down)
                if view_direction.IsAlmostEqualTo(XYZ(0.0, -1.0, 0.0)):
                    line_start = XYZ(min_pt.X, max_pt.Y, min_pt.Z)
                    line_end = XYZ(max_pt.X, min_pt.Y, min_pt.Z)

                # Case 2: (0.000000000, 1.000000000, 0.000000000) - Y+ (up)
                elif view_direction.IsAlmostEqualTo(XYZ(0.0, 1.0, 0.0)):
                    line_start = XYZ(min_pt.X, max_pt.Y, min_pt.Z)
                    line_end = XYZ(max_pt.X, min_pt.Y, min_pt.Z)

                # Case 3: (1.000000000, 0.000000000, 0.000000000) - X+ (right)
                elif view_direction.IsAlmostEqualTo(XYZ(1.0, 0.0, 0.0)):
                    line_start = XYZ(min_pt.X, min_pt.Y, min_pt.Z)
                    line_end = XYZ(max_pt.X, max_pt.Y, min_pt.Z)

                # Case 4: (-1.000000000, 0.000000000, 0.000000000) - X- (left)
                elif view_direction.IsAlmostEqualTo(XYZ(-1.0, 0.0, 0.0)):
                    line_start = XYZ(min_pt.X, min_pt.Y, min_pt.Z)
                    line_end = XYZ(max_pt.X, max_pt.Y, min_pt.Z)

                else:
                    skipped_views.append((view.Name, output.linkify(view.Id)))
                    continue

                # Retrieve levels in the current view
                level_collector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
                
                # Inside your level processing loop
                for level in level_collector:
                    curves = level.GetCurvesInView(DatumExtentType.ViewSpecific, view)
                    for level_curve in curves:
                        # Get the Z value from the curve's end points
                        line_point = level_curve.GetEndPoint(0)
                        z_value = line_point.Z
                        s_factor = view.Scale
                        extension_distance_1 = 12 * s_factor / 100
                        extension_distance_2 = -8 * s_factor / 100

                        level.HideBubbleInView(DatumEnds.End0, view)
                        level.HideBubbleInView(DatumEnds.End1, view)

                        if view_direction.IsAlmostEqualTo(XYZ(0.0, -1.0, 0.0)):
                            start_point_adjusted = XYZ(line_start.X + right_direction.X * extension_distance_2, line_start.Y + right_direction.Y * extension_distance_2, z_value)
                            end_point_adjusted = XYZ(line_end.X + right_direction.X * extension_distance_1, line_end.Y + right_direction.Y * extension_distance_1, z_value)
                            new_level_line = Line.CreateBound(start_point_adjusted, end_point_adjusted)
                            level.SetCurveInView(DatumExtentType.ViewSpecific, view, new_level_line)
                            level.ShowBubbleInView(DatumEnds.End1, view)

                        elif view_direction.IsAlmostEqualTo(XYZ(0.0, 1.0, 0.0)):
                            start_point_adjusted = XYZ(line_start.X + right_direction.X * extension_distance_1, line_start.Y + right_direction.Y * extension_distance_1, z_value)
                            end_point_adjusted = XYZ(line_end.X + right_direction.X * extension_distance_2, line_end.Y + right_direction.Y * extension_distance_2, z_value)
                            new_level_line = Line.CreateBound(start_point_adjusted, end_point_adjusted)
                            level.SetCurveInView(DatumExtentType.ViewSpecific, view, new_level_line)
                            level.ShowBubbleInView(DatumEnds.End0, view)

                        elif view_direction.IsAlmostEqualTo(XYZ(1.0, 0.0, 0.0)):
                            start_point_adjusted = XYZ(line_start.X + right_direction.X * extension_distance_2, line_start.Y + right_direction.Y * extension_distance_2, z_value)
                            end_point_adjusted = XYZ(line_end.X + right_direction.X * extension_distance_1, line_end.Y + right_direction.Y * extension_distance_1, z_value)
                            new_level_line = Line.CreateBound(start_point_adjusted, end_point_adjusted)
                            level.SetCurveInView(DatumExtentType.ViewSpecific, view, new_level_line)
                            level.ShowBubbleInView(DatumEnds.End1, view)

                        elif view_direction.IsAlmostEqualTo(XYZ(-1.0, 0.0, 0.0)):
                            start_point_adjusted = XYZ(line_start.X + right_direction.X * extension_distance_1, line_start.Y + right_direction.Y * extension_distance_1, z_value)
                            end_point_adjusted = XYZ(line_end.X + right_direction.X * extension_distance_2, line_end.Y + right_direction.Y * extension_distance_2, z_value)
                            new_level_line = Line.CreateBound(start_point_adjusted, end_point_adjusted)
                            level.SetCurveInView(DatumExtentType.ViewSpecific, view, new_level_line)
                            level.ShowBubbleInView(DatumEnds.End0, view)

                    total_level_count += 1

        t.Commit()


        # Record the end time
        end_time = time.time()
        runtime = end_time - start_time

        run_result = "Tool ran successfully"
        element_count = total_level_count if total_level_count else 0
        error_occured = "Nil"
        get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)



    except Exception as e:
    # Print error message
        print("Error occurred: {}".format(str(e)))

        # Record the end time and runtime
        end_time = time.time()
        runtime = end_time - start_time

        # Log the error details
        error_occured = "Error occurred: {}".format(str(e))
        run_result = "Error"
        element_count = 10

        # Function to log run data in case of error
        get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)

        t.RollBack()

    finally:
        t.Dispose()

    # Print skipped views message at the end
    if skipped_views:
        output.print_md(header_data)
        # Print the header for skipped views
        output.print_md("## ⚠️ Views Skipped ☹️")  # Markdown Heading 2 
        output.print_md("---")  # Markdown Line Break
        output.print_md("❌ Views skipped due to being inclined or not cardinally aligned. Refer to the **Table Report** below for reference.")  # Print a Line
        
        # Create a table to display the skipped views
        output.print_table(table_data=skipped_views, columns=["VIEW NAME", "ELEMENT ID"])  # Print a Table
        
        print("\n\n")
        output.print_md("---")  # Markdown Line Break

    runtime = time.time() - start_time
    print("Script runtime: {:.2f} seconds".format(runtime))