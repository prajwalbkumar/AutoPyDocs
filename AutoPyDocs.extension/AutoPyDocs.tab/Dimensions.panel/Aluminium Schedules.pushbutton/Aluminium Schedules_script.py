# -*- coding: utf-8 -*-
'''Create Views'''
__title__ = "Create Views"
__author__ = "prajwalbkumar"

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

start_time = time.time()
minoroffset = 1
majoroffset = 2.2
manual_time = 4

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

try:
    view_types = {
        "Floor Plans": ViewType.FloorPlan,
        "Elevations": ViewType.Elevation,
        "All Views on Sheet": ViewType.DrawingSheet
    }

    # Show a dialog box for selecting desired view types
    selected_view_type_names = forms.SelectFromList.show(sorted(view_types.keys()), title='Select View Types', multiselect=True)

    if not selected_view_type_names:
        forms.alert("No view type was selected. Exiting script.", exitscript=True)

    if selected_view_type_names == [sorted(view_types.keys())[0]]:
        # List of all sheets in the document
        all_sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).ToElements()
        if not all_sheets:
            forms.alert("No sheet found in the document", exitscript=True)

        all_sheet_name = {sheet.get_Parameter(BuiltInParameter.SHEET_NAME).AsString() : sheet for sheet in all_sheets}
        selected_sheet_name = forms.SelectFromList.show(sorted(all_sheet_name.keys()), title='Select Views', multiselect=True)    

        if not selected_sheet_name:
            forms.alert("No sheet selected", exitscript=True)

        selected_sheets = [all_sheet_name[name] for name in selected_sheet_name]
        selected_views = [doc.GetElement(view) for sheet in selected_sheets for view in sheet.GetAllPlacedViews()]

    else:
        selected_view_types = [view_types[name] for name in selected_view_type_names if name in view_types] # Convert the selected view type names to their corresponding ViewType enums

        # Collect all views in the document
        views_collector = FilteredElementCollector(doc).OfClass(View)

        # Filter views by the selected types
        filtered_views = [view for view in views_collector if view.ViewType in selected_view_types and not view.IsTemplate]

        if not filtered_views:
            forms.alert("No views of the selected types were found in the document.", exitscript=True) # If no views found, show a message and exit

        # Collect all Viewport elements in the document
        viewports_collector = FilteredElementCollector(doc).OfClass(Viewport)
        views_on_sheets_ids = {viewport.ViewId for viewport in viewports_collector} # Get the IDs of all views that are placed on sheets
        filtered_views_on_sheets = [
            view for view in filtered_views
            if view.Id in views_on_sheets_ids  # Check that the view is placed on a sheet
        ]
        if not filtered_views_on_sheets:
            forms.alert("No views of the selected types were found on sheets in the document.", 
                        exitscript=True) # If no views found, show a message and exit
        view_dict = {view.Name: view for view in filtered_views_on_sheets} # Create a dictionary for view names and their corresponding objects

        # Show the selection window for the user to choose views
        selected_view_names = forms.SelectFromList.show(
            sorted(view_dict.keys()),          # Sort and display the list of view names
            title='Select Views',              # Title of the selection window
            multiselect=True                   # Allow multiple selections
        )

        if selected_view_names:
            # Get the selected views from the dictionary
            selected_views = [view_dict[name] for name in selected_view_names]
            
            # Output or perform actions with the selected views
            #for view in selected_views:
            #    print("Selected View: {}".format(view.Name))
        else:
            forms.alert("No views were selected.")

    t = Transaction(doc, "Dimension Curain Grid")
    t.Start()
    geom_view = None

    # Collect all linked instances
    linked_instance = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    if linked_instance:
        documentation_file = forms.alert("Is this a Documentation File or a Live File", warn_icon=False, options=["Documentation File", "Live File"])

        if not documentation_file:
            forms.alert("No file option selected. Exiting script.", exitscript=True)

        if documentation_file == "Documentation File":
            link_name = []
            for link in linked_instance:
                link_name.append(link.Name)

            ar_instance_name = forms.SelectFromList.show(link_name, title = "Select AR Linked File", width=600, height=600, button_name="Select File", multiselect=False)

            if not ar_instance_name:
                script.exit()

            for link in linked_instance:
                if ar_instance_name == link.Name:
                    ar_instance = link
                    break

            ar_doc = ar_instance.GetLinkDocument()
            if not ar_doc:
                forms.alert("No instance found of the selected AR File.\n"
                            "Use Manage Links to Load the Link in the File!", title = "Link Missing", warn_icon = True)
                script.exit()

            walls_in_view = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
            views_in_link = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()

            for link_view in views_in_link:
                if link_view.Name == "3D-Navisworks-Export":
                    geom_view = link_view
                    break
        
        else:
            linked_instance = None
            ar_doc = doc
    else:
        linked_instance = None
        ar_doc = doc

    equalize_option = forms.alert("Would you like to Equalize equalize dimensions", title="Equalize Dimensions", options= ["Yes", "No"])
    if not equalize_option:
        forms.alert("No Equalize option selected. Exiting script.", exitscript=True)

    dimensioned_items = 0

    for view in selected_views:
        if not linked_instance:
            walls_in_view = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
            geom_view = view

        if not walls_in_view:
            error_in_view = True

        walls_in_view = [wall for wall in walls_in_view if wall.CurtainGrid]

        # For Horizontal Dimension
        vertical_wall_dict = {}
        for wall in walls_in_view:
            wall_location_value = abs((wall.Location.Curve.GetEndPoint(1) - wall.Location.Curve.GetEndPoint(0)).X + (wall.Location.Curve.GetEndPoint(1) - wall.Location.Curve.GetEndPoint(0)).Y)
            grid_location_value = 0
            for line_id in wall.CurtainGrid.GetVGridLineIds():
                line = ar_doc.GetElement(line_id)
                line_curve_point = line.FullCurve.GetEndPoint(0)
                grid_location_value += line_curve_point.X + line_curve_point.Y

            vertical_wall_dict.update({wall: wall_location_value + grid_location_value})
                
        unique_locations = set(vertical_wall_dict.values())
        stacked_walls = {}
        for location in unique_locations:
            stacked_walls[location] = [key for key in vertical_wall_dict.keys() if vertical_wall_dict[key] == location]
        
        # For Vertical Dimension
        aligned_wall_dict = {}
        for wall in walls_in_view:
            wall_location_value = abs(wall.Location.Curve.GetEndPoint(1).Z) + wall.LookupParameter("Base Offset").AsDouble()
            grid_location_value = 0
            for line_id in wall.CurtainGrid.GetUGridLineIds():
                line = ar_doc.GetElement(line_id)
                line_curve_point = line.FullCurve.GetEndPoint(0)
                grid_location_value += line_curve_point.Z
            aligned_wall_dict.update({wall: wall_location_value + grid_location_value})

        unique_locations = set(aligned_wall_dict.values())
        aligned_walls = {}
        for location in unique_locations:
            aligned_walls[location] = [key for key in aligned_wall_dict.keys() if aligned_wall_dict[key] == location]

            
        if view.ViewType == ViewType.Elevation:


            for wall in walls_in_view:
                vertical_grid = []
                horizontal_grid = []
                vertical_point_list = []
                # try:
                wall_line = wall.Location.Curve
                wall_start = wall_line.GetEndPoint(0)
                wall_end = wall_line.GetEndPoint(1)

                angle_to_view = ((wall_start - wall_end).Normalize()).AngleTo(view.ViewDirection)
                rounded_result = round(angle_to_view, 4)
                if rounded_result == 1.5708: # A constant value do not change
                    
                    # Extracting Physical Edges
                    options = Options()
                    options.View = view
                    options.IncludeNonVisibleObjects = True
                    options.ComputeReferences = True

                    vertical_array = ReferenceArray()
                    vertical_overall_array = ReferenceArray()
                    internal_vertical_array = ReferenceArray()
                    first_vertical_array = ReferenceArray()
                    last_vertical_array = ReferenceArray()
                    horizontal_array = ReferenceArray()
                    horizontal_overall_array = ReferenceArray()

                    # Extract the top stacked curtain wall visbile in view
                    wall_location = abs(wall.Location.Curve.GetEndPoint(1).Z) + wall.LookupParameter("Base Offset").AsDouble()
                    grid_location_value = 0
                    for line_id in wall.CurtainGrid.GetUGridLineIds():
                        line = ar_doc.GetElement(line_id)
                        line_curve_point = line.FullCurve.GetEndPoint(0)
                        grid_location_value += line_curve_point.Z
                    unique_identifier = wall_location + grid_location_value

                    if unique_identifier in aligned_walls:
                        query_aligned_walls = aligned_walls[unique_identifier]

                        if len(query_aligned_walls) > 1:
                            left_wall_point = 0
                            left_wall = None

                            for wall in query_aligned_walls:
                                if wall not in walls_in_view:
                                    continue
                                # Get the Left Most Wall
                                current_wall_point = wall.Location.Curve.GetEndPoint(1).X + wall.Location.Curve.GetEndPoint(1).Y
                                if  current_wall_point > left_wall_point:
                                    left_wall_point = current_wall_point
                                    left_wall = wall

                        else:
                            left_wall = wall
                        
                        # Once the topmost wall of a stack is retrieved, remove the stack from the dict all together
                        aligned_walls.pop(unique_identifier)

                    else:
                        continue

                    for geometry in left_wall.get_Geometry(options):
                        if (geometry.ToString() == "Autodesk.Revit.DB.Solid"):
                            faces = geometry.Faces
                            for face in faces:
                                faceNormal = face.FaceNormal
                                if not int(faceNormal.Z) == 0: #and int(faceNormal.Y) == 0
                                    if linked_instance:
                                        horizontal_array.Append(face.Reference.CreateLinkReference(ar_instance))
                                        horizontal_overall_array.Append(face.Reference.CreateLinkReference(ar_instance))
                                    else:
                                        horizontal_array.Append(face.Reference)
                                        horizontal_overall_array.Append(face.Reference)
                                    continue

                    if horizontal_array:
                        # Vertical Dimensions
                        wall_line = left_wall.Location.Curve
                        wall_start = wall_line.GetEndPoint(0)
                        wall_end = wall_line.GetEndPoint(1)

                        if wall_start.X + wall_start.Y < wall_end.X + wall_end.Y:
                            vertical_line = Line.CreateUnbound(wall_line.Evaluate((wall_line.Length + minoroffset), False), XYZ(0,0,1))
                            major_vertical_line = Line.CreateUnbound(wall_line.Evaluate((wall_line.Length + majoroffset), False), XYZ(0,0,1))
                        else:
                            vertical_line = Line.CreateUnbound(wall_line.Evaluate((wall_line.Length + minoroffset), False), XYZ(0,0,1))
                            major_vertical_line = Line.CreateUnbound(wall_line.Evaluate((wall_line.Length + majoroffset), False), XYZ(0,0,1))

                        dimensioned_items += horizontal_overall_array.Size
                        doc.Create.NewDimension(view, major_vertical_line, horizontal_overall_array)

                        if left_wall.CurtainGrid.NumULines > 0:
                            horizontal_grid_ids = left_wall.CurtainGrid.GetUGridLineIds()
                            for id in horizontal_grid_ids:
                                grid = ar_doc.GetElement(id)
                                horizontal_grid.append(grid)
                            for grid in horizontal_grid:
                                checkpoint = []
                                grid_lines = grid.get_Geometry(options)
                                for line in grid_lines:
                                    if not line.GetEndPoint(0).Z in checkpoint:
                                        if linked_instance:
                                            horizontal_array.Append(line.Reference.CreateLinkReference(ar_instance))
                                        else:
                                            horizontal_array.Append(line.Reference)
                                        checkpoint.append(line.GetEndPoint(0).Z)
                            dimensioned_items += horizontal_array.Size
                            doc.Create.NewDimension(view, vertical_line, horizontal_array)
                    


                    # Horizontal Dimension Starts
                    # Extract the top stacked curtain wall visbile in view
                    wall_location = abs((wall.Location.Curve.GetEndPoint(1) - wall.Location.Curve.GetEndPoint(0)).X + (wall.Location.Curve.GetEndPoint(1) - wall.Location.Curve.GetEndPoint(0)).Y)
                    grid_location_value = 0
                    for line_id in wall.CurtainGrid.GetVGridLineIds():
                        line = ar_doc.GetElement(line_id)
                        line_curve_point = line.FullCurve.GetEndPoint(0)
                        grid_location_value += line_curve_point.X + line_curve_point.Y
                    unique_identifier = wall_location + grid_location_value

                    if unique_identifier in stacked_walls:
                        query_stacked_wall = stacked_walls[unique_identifier]

                        if len(query_stacked_wall) > 1:
                            top_wall_height = 0
                            top_wall = None

                            for wall in query_stacked_wall:
                                if wall not in walls_in_view:
                                    continue
                                current_wall_height = wall.Location.Curve.GetEndPoint(1).Z + wall.LookupParameter("Unconnected Height").AsDouble() + wall.LookupParameter("Base Offset").AsDouble()
                                if  current_wall_height > top_wall_height:
                                    top_wall_height = current_wall_height
                                    top_wall = wall

                        else:
                            top_wall = wall
                        
                        # Once the topmost wall of a stack is retrieved, remove the stack from the dict all together
                        stacked_walls.pop(unique_identifier)

                    else:
                        continue

                    for geometry in top_wall.get_Geometry(options):
                        if (geometry.ToString() == "Autodesk.Revit.DB.Solid"):
                            faces = geometry.Faces
                            for face in faces:
                                faceNormal = face.FaceNormal
                                if not int(faceNormal.Z) == 0:
                                    continue
                                face_normal_condition = int(face.FaceNormal.DotProduct(view.ViewDirection))
                                if not abs(face_normal_condition) == 1:
                                    if linked_instance:
                                        vertical_array.Append(face.Reference.CreateLinkReference(ar_instance))
                                        vertical_overall_array.Append(face.Reference.CreateLinkReference(ar_instance))
                                    else:
                                        vertical_array.Append(face.Reference)
                                        vertical_overall_array.Append(face.Reference)

                    if vertical_array:
                        # Horizontal Dimensions
                        wall_line = top_wall.Location.Curve
                        wall_start = wall_line.GetEndPoint(0)
                        wall_end = wall_line.GetEndPoint(1)

                        unconnected_height = top_wall.LookupParameter("Unconnected Height").AsDouble()
                        wall_base_offset = top_wall.LookupParameter("Base Offset").AsDouble()
                        major_horizontal_line = Line.CreateBound(XYZ(wall_start.X, wall_start.Y, wall_start.Z + wall_base_offset + unconnected_height + majoroffset), XYZ(wall_end.X, wall_end.Y, wall_end.Z + wall_base_offset + unconnected_height + majoroffset))
                        horizontal_line = Line.CreateBound(XYZ(wall_start.X, wall_start.Y, wall_start.Z + wall_base_offset + unconnected_height + minoroffset), XYZ(wall_end.X, wall_end.Y, wall_end.Z + wall_base_offset + unconnected_height + minoroffset))
                        dimensioned_items += vertical_overall_array.Size
                        doc.Create.NewDimension(view, major_horizontal_line, vertical_overall_array)

                        if top_wall.CurtainGrid.NumVLines > 0:
                            vertical_grid_ids = top_wall.CurtainGrid.GetVGridLineIds()
                            for id in vertical_grid_ids:
                                grid = ar_doc.GetElement(id)
                                vertical_grid.append(grid)

                            for i, grid in enumerate(vertical_grid):
                                checkpoint = []
                                grid_lines = grid.get_Geometry(options)
                                for line in grid_lines:
                                    end_point = line.GetEndPoint(0)
                                    if not round(end_point.X, 4) + round(end_point.Y, 4) in checkpoint:                                                
                                        if linked_instance:
                                            internal_vertical_array.Append(line.Reference.CreateLinkReference(ar_instance))
                                            vertical_array.Append(line.Reference.CreateLinkReference(ar_instance))
                                        else:
                                            internal_vertical_array.Append(line.Reference)
                                            vertical_array.Append(line.Reference)
                                        checkpoint.append(round(end_point.X, 4) + round(end_point.Y, 4))

                            if top_wall.CurtainGrid.NumVLines > 2:

                                first_vertical_array.Append(internal_vertical_array[0])
                                first_vertical_array.Append(vertical_overall_array[0])

                                last_vertical_array.Append(internal_vertical_array[internal_vertical_array.Size - 1])
                                last_vertical_array.Append(vertical_overall_array[vertical_overall_array.Size - 1])
                                dimensioned_items += first_vertical_array.Size
                                doc.Create.NewDimension(view, horizontal_line, first_vertical_array)
                                dimensioned_items += last_vertical_array.Size
                                doc.Create.NewDimension(view, horizontal_line, last_vertical_array)
                                dimensioned_items += internal_vertical_array.Size
                                internal_dimension = doc.Create.NewDimension(view, horizontal_line, internal_vertical_array)

                                if equalize_option == "Yes":
                                    segment_value = set()
                                    segments = internal_dimension.Segments
                                    for segment in segments:
                                        segment_value.add(segment.ValueString)
                                    if len(segment_value) == 1: 
                                        internal_dimension.LookupParameter("Equality Display").Set(1)

                            else:
                                dimensioned_items += vertical_array.Size
                                doc.Create.NewDimension(view, horizontal_line, vertical_array)
                #     dimensioned_items += 1          

                # except:
                #     continue

        if view.ViewType == ViewType.FloorPlan:
            plan_minoroffset = - minoroffset
            plan_majoroffset = - majoroffset
            for wall in walls_in_view:
                vertical_grid = []
                horizontal_grid = []

                try:
                    # if wall.CurtainGrid:
                    if ar_doc.GetElement(wall.LevelId).Name != view.GenLevel.Name:
                        continue
                    wall_line = wall.Location.Curve
                    # Extracting Physical Edges
                    options = Options()
                    options.View = view
                    options.IncludeNonVisibleObjects = True
                    options.ComputeReferences = True

                    vertical_array = ReferenceArray()
                    vertical_overall_array = ReferenceArray()
                    internal_vertical_array = ReferenceArray()
                    first_vertical_array = ReferenceArray()
                    last_vertical_array = ReferenceArray()
                    horizontal_array = ReferenceArray()
                    horizontal_overall_array = ReferenceArray()
                    
                    count = 0
                    for geometry in wall.get_Geometry(options):
                        if count == 2:
                            break
                        if (geometry.ToString() == "Autodesk.Revit.DB.Solid"):
                            faces = geometry.Faces
                            for face in faces:
                                if linked_instance:
                                    vertical_array.Append(face.Reference.CreateLinkReference(ar_instance))
                                    vertical_overall_array.Append(face.Reference.CreateLinkReference(ar_instance))
                                else:
                                    vertical_array.Append(face.Reference)
                                    vertical_overall_array.Append(face.Reference)
                                    
                            count += 1

                    if vertical_array:
                        # Horizontal Dimensions
                        major_horizontal_line = wall_line.CreateOffset(majoroffset, XYZ.BasisZ)
                        horizontal_line = wall_line.CreateOffset(minoroffset, XYZ.BasisZ)

                        dimensioned_items += vertical_overall_array.Size
                        doc.Create.NewDimension(view, major_horizontal_line, vertical_overall_array)

                        if wall.CurtainGrid.NumVLines > 0:
                            vertical_grid_ids = wall.CurtainGrid.GetVGridLineIds()
                            for id in vertical_grid_ids:
                                grid = ar_doc.GetElement(id)
                                vertical_grid.append(grid)

                            for grid in vertical_grid:
                                # checkpoint = []
                                grid_lines = grid.get_Geometry(options)
                                for line in grid_lines:
                                    end_point = line.GetEndPoint(0)
                                    # if not round(end_point.X, 4) + round(end_point.Y, 4) in checkpoint:
                                    if linked_instance:
                                        internal_vertical_array.Append(line.Reference.CreateLinkReference(ar_instance))
                                        vertical_array.Append(line.Reference.CreateLinkReference(ar_instance))
                                    else:
                                        internal_vertical_array.Append(line.Reference)
                                        vertical_array.Append(line.Reference)

                                    break
                                        # checkpoint.append(round(end_point.X, 4) + round(end_point.Y, 4))

                            if wall.CurtainGrid.NumVLines > 2:

                                first_vertical_array.Append(internal_vertical_array[0])
                                first_vertical_array.Append(vertical_overall_array[0])

                                last_vertical_array.Append(internal_vertical_array[internal_vertical_array.Size - 1])
                                last_vertical_array.Append(vertical_overall_array[vertical_overall_array.Size - 1])

                                doc.Create.NewDimension(view, horizontal_line, first_vertical_array)
                                dimensioned_items += first_vertical_array.Size
                                doc.Create.NewDimension(view, horizontal_line, last_vertical_array)
                                dimensioned_items += last_vertical_array.Size
                                internal_dimension = doc.Create.NewDimension(view, horizontal_line, internal_vertical_array)
                                dimensioned_items += internal_vertical_array.Size

                                if equalize_option == "Yes":
                                    segment_value = set()
                                    segments = internal_dimension.Segments
                                    for segment in segments:
                                        segment_value.add(segment.ValueString)
                                    if len(segment_value) == 1: 
                                        internal_dimension.LookupParameter("Equality Display").Set(1)

                            else:
                                dimensioned_items += vertical_array.Size
                                doc.Create.NewDimension(view, horizontal_line, vertical_array)
 
                except:
                    continue
                            
    t.Commit()

    # Record the end time
    end_time = time.time()
    runtime = end_time - start_time

    run_result = "Tool ran successfully"
    element_count = dimensioned_items 
    error_occured = "Nil"

    # Function to log run data
    get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)

except Exception as e:
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