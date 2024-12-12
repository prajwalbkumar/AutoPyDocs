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
manual_time = 300

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


dimensions = []

skipped_views = []
view_types = {
    'Floor Plans': ViewType.FloorPlan,
    'Reflected Ceiling Plans': ViewType.CeilingPlan,
    'Area Plans': ViewType.AreaPlan, 
    'Elevations': ViewType.Elevation
}

# Show a dialog box for selecting desired view types
selected_view_type_names = forms.SelectFromList.show(
    sorted(view_types.keys()),            # Sorted list of view type names
    title='Select View Types',                # Title of the dialog box
    multiselect=True                          # Allow multiple selections
)
if not selected_view_type_names:
    forms.alert("No view type was selected. Exiting script.", 
                exitscript=True)
selected_view_types = [view_types[name] 
                    for name in selected_view_type_names 
                    if name in view_types] # Convert the selected view type names to their corresponding ViewType enums

# Collect all views in the document
views_collector = FilteredElementCollector(doc).OfClass(View)

# Filter views by the selected types
filtered_views = [view for view in views_collector 
                if view.ViewType in selected_view_types 
                and not view.IsTemplate]
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




#Collect curtain wall types from the document

wall_types = FilteredElementCollector(doc).OfClass(WallType).ToElements()

# Filter only curtain wall types
curtain_wall_types = [wt for wt in wall_types if wt.Kind == WallKind.Curtain]

# Get the names of the curtain wall types

curtain_wall_type_names = [wt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() for wt in curtain_wall_types]

#Show selection dialog for curtain wall types
selected_curtain_wall_type_names = forms.SelectFromList.show(sorted(curtain_wall_type_names),title="Select Curtain Wall Types to Dimension",multiselect=True)
if not selected_curtain_wall_type_names:
    forms.alert("No curtain wall types were selected. Exiting script.", exitscript=True)

# Filter the selected curtain wall types
selected_curtain_wall_types = [
    wt for wt in curtain_wall_types
    if wt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() in selected_curtain_wall_type_names
]


# Prompt the user at the start of the script
equalize_choice = forms.alert(
    "Do you want to equalize the dimensions?",
    title="Equalize Dimensions",
    yes=True,
    no=True
)

# Exit if the user cancels the prompt
if equalize_choice is None:
    print("User canceled the script.")
    script.exit()

# Proceed based on user choice
equalize_dimensions = True if equalize_choice else False



for view in selected_views:

    view_bounding_box_transform = view.get_BoundingBox(None).Transform.Inverse
    view_outline = Outline(view_bounding_box_transform.OfPoint(view.get_BoundingBox(None).Min), view_bounding_box_transform.OfPoint(view.get_BoundingBox(None).Max))

    # Collect all linked instances
    linked_instance = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    '''if linked_instance:
        link_name = []
        for link in linked_instance:
            link_name.append(link.Name)

        ar_instance_name = forms.SelectFromList.show(link_name, title = "Select URS File", width=600, height=600, button_name="Select File", multiselect=False)

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

        # link_transform = ar_instance.GetTotalTransform().Inverse

        # if view.ViewType == ViewType.Elevation:
        #     view_direction = view.ViewDirection  # Normal vector to the view plane
        #     view_origin = view.Origin  # Elevation view’s origin
        #     view_box = view.get_BoundingBox(None)

        #     # # Define offset and depth for filtering within the view's bounding box range
        #     # offset = 0  # Offset distance from view origin, if any
        #     # depth = view_box.Max.Z - view_box.Min.Z  # Example depth within clipping planes

        #     
        #     # # Define transformed bounding box in the linked document’s coordinate system
        #     # inverse_transform = link_transform.Inverse
        #     # min_point = inverse_transform.OfPoint(view_origin + (view_direction * offset))
        #     # max_point = inverse_transform.OfPoint(view_origin + (view_direction * (offset + depth)))

        #     transformed_bounding_box = Outline(view_box.Min, view_box.Max)

        #     # Collect elements in the linked document within the bounding box
            # walls_in_view = FilteredElementCollector(ar_doc).WherePasses(BoundingBoxIntersectsFilter(view_outline)).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()'''


    walls_in_view = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()

    # Filter walls to match the selected curtain wall types based on name
    filtered_walls_in_view = [
        wall for wall in walls_in_view
        if wall.Name in selected_curtain_wall_type_names  # Compare the wall's type name to the selected ones
    ]

    walls_processed = 0
    if not filtered_walls_in_view:
        skipped_views.append((view.Name, output.linkify(view.Id)))
        continue

    try:
            
        t = Transaction(doc, "Dimension Curain Grid")
        t.Start()

        if view.ViewType == ViewType.Elevation:
            for wall in filtered_walls_in_view:
                vertical_grid = []
                horizontal_grid = []
                vertical_point_list = []
                horizontal_point_list = []
                try:
                    if wall.CurtainGrid:
                        wall_line = wall.Location.Curve
                        wall_start = wall_line.GetEndPoint(0)
                        wall_end = wall_line.GetEndPoint(1)

                        # if not view_outline.Contains(link_transform.OfPoint(wall_start - wall_end), 0):
                        #     print("YAY")
                        #     continue

                        angle_to_view = ((wall_start - wall_end).Normalize()).AngleTo(view.ViewDirection)
                        rounded_result = round(angle_to_view, 4)
                        if rounded_result == 1.5708:
                            # Extracting Physical Edges
                            options = Options()
                            options.View = view
                            options.IncludeNonVisibleObjects = True
                            options.ComputeReferences = True

                            vertical_array = ReferenceArray()
                            vertical_overall_array = ReferenceArray()
                            horizontal_array = ReferenceArray()
                            horizontal_overall_array = ReferenceArray()

                            for geometry in wall.get_Geometry(options):
                                if (geometry.ToString() == "Autodesk.Revit.DB.Solid"):
                                    faces = geometry.Faces
                                    for face in faces:
                                        faceNormal = face.FaceNormal
                                        if not int(faceNormal.Z) == 0: #and int(faceNormal.Y) == 0
                                            '''if linked_instance:
                                                horizontal_array.Append(face.Reference.CreateLinkReference(ar_instance))
                                                horizontal_overall_array.Append(face.Reference.CreateLinkReference(ar_instance))'''

                                            horizontal_array.Append(face.Reference)
                                            horizontal_overall_array.Append(face.Reference)
                                            continue

                                        face_normal_condition = int(face.FaceNormal.DotProduct(view.ViewDirection))
                                        if not abs(face_normal_condition) == 1:
                                            '''if linked_instance:
                                                vertical_array.Append(face.Reference.CreateLinkReference(ar_instance))
                                                vertical_overall_array.Append(face.Reference.CreateLinkReference(ar_instance))'''
                                            
                                            vertical_array.Append(face.Reference)
                                            vertical_overall_array.Append(face.Reference)

                            if horizontal_array:
                                # Vertical Dimensions
                                if wall_start.X + wall_start.Y < wall_end.X + wall_end.Y:
                                    vertical_line = Line.CreateUnbound(wall_line.Evaluate((wall_line.Length + minoroffset), False), XYZ(0,0,1))
                                    major_vertical_line = Line.CreateUnbound(wall_line.Evaluate((wall_line.Length + majoroffset), False), XYZ(0,0,1))
                                else:
                                    vertical_line = Line.CreateUnbound(wall_line.Evaluate((wall_line.Length + minoroffset), False), XYZ(0,0,1))
                                    major_vertical_line = Line.CreateUnbound(wall_line.Evaluate((wall_line.Length + majoroffset), False), XYZ(0,0,1))

                                dim = doc.Create.NewDimension(view, major_vertical_line, horizontal_overall_array)

                                if wall.CurtainGrid.NumULines > 0:
                                    horizontal_grid_ids = wall.CurtainGrid.GetUGridLineIds()
                                    for id in horizontal_grid_ids:
                                        grid = doc.GetElement(id)
                                        horizontal_grid.append(grid)

                                    for grid in horizontal_grid:
                                        checkpoint = []
                                        grid_lines = grid.get_Geometry(options)
                                        for line in grid_lines:
                                            if not line.GetEndPoint(0).Z in checkpoint:
                                                '''if linked_instance:
                                                    horizontal_array.Append(line.Reference.CreateLinkReference(ar_instance))'''
                                                
                                                horizontal_array.Append(line.Reference)
                                                checkpoint.append(line.GetEndPoint(0).Z)

                                    dim = doc.Create.NewDimension(view, vertical_line, horizontal_array)

                            if vertical_array:
                                # Horizontal Dimensions
                                unconnected_height = wall.LookupParameter("Unconnected Height").AsDouble()
                                major_horizontal_line = Line.CreateBound(XYZ(wall_start.X, wall_start.Y, wall_start.Z + unconnected_height + majoroffset), XYZ(wall_end.X, wall_end.Y, wall_end.Z + unconnected_height + majoroffset))
                                horizontal_line = Line.CreateBound(XYZ(wall_start.X, wall_start.Y, wall_start.Z + unconnected_height + minoroffset), XYZ(wall_end.X, wall_end.Y, wall_end.Z + unconnected_height + minoroffset))
                                doc.Create.NewDimension(view, major_horizontal_line, vertical_overall_array)
                                

                                if wall.CurtainGrid.NumVLines > 0:
                                    vertical_grid_ids = wall.CurtainGrid.GetVGridLineIds()
                                    for id in vertical_grid_ids:
                                        grid = doc.GetElement(id)
                                        vertical_grid.append(grid)

                                    for grid in vertical_grid:
                                        checkpoint = []
                                        grid_lines = grid.get_Geometry(options)
                                        for line in grid_lines:
                                            end_point = line.GetEndPoint(0)
                                            if not (end_point.X + end_point.Y) in checkpoint:
                                                '''if linked_instance:
                                                    vertical_array.Append(line.Reference.CreateLinkReference(ar_instance))'''
                                                
                                                vertical_array.Append(line.Reference)
                                                checkpoint.append(round(end_point.X, 4) + round(end_point.Y, 4))

                                    dim = doc.Create.NewDimension(view, horizontal_line, vertical_array)
                                    dimensions.append(dim)
                                    



                            walls_processed += 1
                except:
                    continue

        if view.ViewType == ViewType.FloorPlan:
            minoroffset = - minoroffset
            majoroffset = - majoroffset
            for wall in filtered_walls_in_view:
                vertical_grid = []
                horizontal_grid = []
                vertical_point_list = []
                horizontal_point_list = []
                try:
                    if wall.CurtainGrid:
                        if doc.GetElement(wall.LevelId).Name != view.GenLevel.Name:
                            continue
                        wall_line = wall.Location.Curve
                        # Extracting Physical Edges
                        options = Options()
                        options.View = view
                        options.IncludeNonVisibleObjects = True
                        options.ComputeReferences = True

                        vertical_array = ReferenceArray()
                        vertical_overall_array = ReferenceArray()
                        horizontal_array = ReferenceArray()
                        horizontal_overall_array = ReferenceArray()
                        count = 0
                        for geometry in wall.get_Geometry(options):
                            if count == 2:
                                break
                            if (geometry.ToString() == "Autodesk.Revit.DB.Solid"):
                                faces = geometry.Faces
                                for face in faces:
                                    '''if linked_instance:
                                        vertical_array.Append(face.Reference.CreateLinkReference(ar_instance))
                                        vertical_overall_array.Append(face.Reference.CreateLinkReference(ar_instance))'''
                                    
                                    vertical_array.Append(face.Reference)
                                    vertical_overall_array.Append(face.Reference)
                                        
                                count += 1

                        if vertical_array:
                            # Horizontal Dimensions
                            major_horizontal_line = wall_line.CreateOffset(majoroffset, XYZ.BasisZ)
                            horizontal_line = wall_line.CreateOffset(minoroffset, XYZ.BasisZ)

                            # plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0,0,0))
                            # sketch_plane = SketchPlane.Create(doc, plane)
                            # model_line = doc.Create.NewModelCurve(horizontal_line, sketch_plane)

                            dim = doc.Create.NewDimension(view, major_horizontal_line, vertical_overall_array)


                            if wall.CurtainGrid.NumVLines > 0:
                                vertical_grid_ids = wall.CurtainGrid.GetVGridLineIds()
                                for id in vertical_grid_ids:
                                    grid = doc.GetElement(id)
                                    vertical_grid.append(grid)

                                for grid in vertical_grid:
                                    checkpoint = []
                                    grid_lines = grid.get_Geometry(options)
                                    for line in grid_lines:
                                        end_point = line.GetEndPoint(0)
                                        if not (end_point.X + end_point.Y) in checkpoint:
                                            '''if linked_instance:
                                                vertical_array.Append(line.Reference.CreateLinkReference(ar_instance))'''

                                            vertical_array.Append(line.Reference)
                                            checkpoint.append(round(end_point.X, 4) + round(end_point.Y, 4))


                                dim = doc.Create.NewDimension(view, horizontal_line, vertical_array)
                                if equalize_choice:
                                    eq_param = dim.LookupParameter("Equality Display")
                                    if eq_param:
                                        eq_param.Set(1)  # Enable equality


                                # if unequal_panel:
                                #     for i, gridd in enumerate(vertical_grid):
                                #         if i == 0 or i == len(vertical_grid) - 1: 
                                #             checkpoint = []
                                #             grid_lines = gridd.get_Geometry(options)
                                #             for line in grid_lines:
                                #                 end_point = line.GetEndPoint(0)
                                #                 if not (end_point.X + end_point.Y) in checkpoint:
                                #                     vertical_array3.Append(line.Reference)
                                #                     print(vertical_array3)
                                #                     checkpoint.append(round(end_point.X, 4) + round(end_point.Y, 4))

                                #                     dim = doc.Create.NewDimension(view, horizontal_line, vertical_array3)
                                #                     dim.AreSegmentsEqual = not dim.AreSegmentsEqual

                            walls_processed += 1
                except:
                    continue
                
        if equalize_choice:
            for dim in dimensions:
                eq_param = dim.LookupParameter("Equality Display")
                if eq_param:
                    eq_param.Set(1)  # Enable equality

          
        t.Commit()

        


        # Record the end time
        end_time = time.time()
        runtime = end_time - start_time

        run_result = "Tool ran successfully"
        element_count = walls_processed 
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

# Print skipped views message at the end
if skipped_views:
    output.print_md(header_data)
    # Print the header for skipped views
    output.print_md("## ⚠️ Views Skipped ☹️")  # Markdown Heading 2 
    output.print_md("---")  # Markdown Line Break
    output.print_md("❌ Selected Curtain wall types not found in selected Views. Refer to the **Table Report** below for reference.")  # Print a Line
    
    # Create a table to display the skipped views
    output.print_table(table_data=skipped_views, columns=["VIEW NAME", "ELEMENT ID"])  # Print a Table
    
    print("\n\n")
output.print_md("---")  # Markdown Line Break