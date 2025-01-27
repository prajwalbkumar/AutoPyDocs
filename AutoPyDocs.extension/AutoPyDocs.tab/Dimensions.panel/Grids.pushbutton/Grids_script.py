# -*- coding: utf-8 -*-
'''Align & Dim Grids'''
__title__ = " Align & Dim Grids"
__author__ = "romaramnani"

import clr
import time

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB  import *
from pyrevit            import revit, forms, script
from Autodesk.Revit.DB  import FilteredElementCollector, BuiltInCategory, View, ViewType, Viewport, XYZ
from collections        import defaultdict
from datetime           import datetime

from view_functions     import align_grids, get_grids_in_view
from Extract.RunData    import get_run_data
from doc_functions      import get_view_on_sheets
from g_curve_functions  import refArray, refLine

import Autodesk.Revit.DB    as DB

start_time = time.time()
manual_time = 500 

doc = revit.doc
output = script.get_output()
app = __revit__.Application 
rvt_year = "Revit" + str(app.SubVersionNumber)
model_name = doc.Title
tool_name = __title__ 
user_name = app.Username
header_data = " | ".join([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), rvt_year, tool_name, model_name, user_name])

view_types = {
        'Floor Plans': ViewType.FloorPlan,
        'Reflected Ceiling Plans': ViewType.CeilingPlan,
        'Area Plans': ViewType.AreaPlan, 
        'Sections' : ViewType.Section, 
        'Elevations': ViewType.Elevation
    }
selected_views = get_view_on_sheets(doc, view_types)

def offset_line(revit_line, offset_distance):
    start_point = revit_line.GetEndPoint(0)
    end_point = revit_line.GetEndPoint(1)
    direction = (end_point - start_point).Normalize()
    normal = DB.XYZ.BasisZ.CrossProduct(direction)

    offset_line = revit_line.CreateOffset(offset_distance, normal)
    return offset_line

def get_orientation(grids):
    start = grids.Curve.GetEndPoint(0)
    end = grids.Curve.GetEndPoint(1)
    
    dx = abs(start.X - end.X)
    dy = abs(start.Y - end.Y)
  
    if dx > dy:  # More horizontal
        return "horizontal"
    elif dx < dy:  # More vertical
        return "vertical"
    elif dx != 0 and dy != 0:  
        return "diagonal"
    else:  
        return "unknown"

def get_grid_radius(grids):
    for grid in grids:
        arc = grid.Curve
        c = arc.Center
        point = arc.GetEndPoint(0)
        radius = point.DistanceTo(c)
        return radius

def convert_to_xyz(float_list):
    # Assuming the float_list contains coordinates in the order [X, Y, Z]
    if len(float_list) == 3:
        return DB.XYZ(float_list[0], float_list[1], float_list[2])
    else:
        forms.alert("List must contain exactly 3 float values.")
    
def change_to_2D_extents(grids):
    for grid in grids:
     grid.SetDatumExtentType(DB.DatumEnds.End0, view, DB.DatumExtentType.ViewSpecific)
     grid.SetDatumExtentType(DB.DatumEnds.End1, view, DB.DatumExtentType.ViewSpecific)
    return grids

def grid_points(grids, view):
    start_points = []
    end_points = []

    for grid in grids:
        if isinstance(grid, DB.Grid):
            grid_curve = grid.GetCurvesInView(DB.DatumExtentType.ViewSpecific, view)  # Get the 2D extent curve of the grid line
            if grid_curve:
                for curve in grid_curve:
                    start_points.append(curve.GetEndPoint(0))
                    end_points.append(curve.GetEndPoint(1))
            else:
                print("No curves found for the grid in the specified view.")

    # Ensure all points are XYZ objects
    start_points = [convert_to_xyz([point.X, point.Y, point.Z]) if not isinstance(point, DB.XYZ) else point for point in start_points]
    end_points = [convert_to_xyz([point.X, point.Y, point.Z]) if not isinstance(point, DB.XYZ) else point for point in end_points]

    '''# Print the points for debugging
    for point in start_points:
        print("Start Point: {}, Coordinates: ({}, {}, {})".format(point, point.X, point.Y, point.Z))
    for point in end_points:
        print("End Point: {}, Coordinates: ({}, {}, {})".format(point, point.X, point.Y, point.Z))'''

    return start_points, end_points, grid_curve

def gradient(grid):
    gr = None
    start = grid.Curve.GetEndPoint(0)
    end = grid.Curve.GetEndPoint(1)
    if round(start.X, 10) != round(end.X, 10):
        gr = round((1.0 / (start.X - end.X)) * (start.Y - end.Y), 10)
    return gr

def parallel(gridA, gridB):
    return gradient(gridA) == gradient(gridB)

def offset_line_plan(revit_line, offset_distance):
    # Calculate the direction & normal vector of the line
    normal = DB.XYZ.BasisZ
        
    # Create an offset line using the normal vector and the offset distance
    offset_line = revit_line.CreateOffset(offset_distance, normal)
    return offset_line

def offset_line_elev(revit_line, offset_distance):
    
    start_point = revit_line.GetEndPoint(0)
    end_point = revit_line.GetEndPoint(1)

    z1 = start_point.Z + offset_distance
    z2 = end_point.Z + offset_distance

    new_start_point = XYZ(start_point.X, start_point.Y, z1)
    new_end_point = XYZ(end_point.X, end_point.Y, z2)

    # Create an offset line using the normal vector and the offset distance
    offset_line = DB.Line.CreateBound(new_start_point, new_end_point)
    return offset_line


# t = DB.Transaction(doc, "CropView")
# t.Start()
# for view in selected_views:
#     ensure_view_is_cropped(view)
# t.Commit()

align_choice = forms.alert(
    "Are your grids aligned?",
    options= ['No, please align grids and dimension them!', 'Yes, only dimension grids!'],
    exitscript=True)
if not align_choice:
    forms.alert("No selction found. Exiting script.", exitscript=True)
if align_choice == 'No, please align grids and dimension them!':
    align_grids(doc, selected_views)

grid_list_length = 0
failed_data = []
view_issue_counts = {}

try:

    # Create transaction to create dimensions
    t = DB.Transaction(doc, "Dimension grids")
    t.Start()

    total_view_count = 0
    view_list_length =0
    total_grid_count = 0
    grid_list_length = 0

    # Get all grids on selected view
    for view in selected_views:
        view_name = view.Name
        view_type = view.ViewType
        view_id = view.Id
        view_key = output.linkify(view_id)
        if view_key not in view_issue_counts:
            view_issue_counts[view_key] = {
                        "Name": view_name,
                        "ViewType": view_type,
                        "DimensionError": 0
            }
        grids = get_grids_in_view(doc,view)
        #print(grids)
        gridGroups = {}
        excludedGrids = []
        flat_gr_list = None

        s_factor = view.Scale

        # Loop through all grids
        for grid in grids:
            grid_list = []
            gridName = grid.LookupParameter("Name").AsString()
            gridCurve = grid.Curve

            # Check if grid is already classified
            if gridName not in excludedGrids:
                # Check if the rest of the grids are parallel 
                for g in grids:
                    inter = gridCurve.Intersect(g.Curve)
                    gName = g.LookupParameter("Name").AsString()

                    # Check parallel grids and group them
                    if gName not in excludedGrids and inter == DB.SetComparisonResult.Disjoint and parallel(grid, g):
                        if gridName not in gridGroups.keys():
                            gridGroups[gridName] = [grid]
                            excludedGrids.append(gridName)
                        gridGroups[gridName].append(g)
                        excludedGrids.append(gName)
            grid_list_length += 1

        # Output the results for the current view
        #print("Processing view: {}".format(view.Name))
        #print("Number of grid groups found: {}".format(len(gridGroups)))

        sorted_groups = defaultdict(list)
        # Sort grids based on orientation
        for i, (group_name, group) in enumerate(gridGroups.items(), start=1):
            sorted_grids = []
        
            orientation = get_orientation(group[0])
            '''sorted_radial_grids = sorted(grids, key=lambda grid: get_grid_radius(grid))'''

            if orientation == "horizontal":
                sorted_grids = sorted(group, key=lambda grid: grid.Curve.GetEndPoint(0).Y)
            elif orientation == "vertical":
                sorted_grids  = sorted(group, key=lambda grid: grid.Curve.GetEndPoint(0).X, reverse=False)
            elif orientation == "diagonal":
                sorted_grids = sorted(group, key=lambda grid: grid.Curve.GetEndPoint(0).X)
            else:
                print("Non-standard grid found")
                group

            sorted_groups[i] = sorted_grids
            #print("No. of grids found in group {}: {}".format(i, len(group)))
            
            #for j, grid in enumerate(sorted_grids):
            #print("Grid {} in group {}: {}".format(j + 1, i, grid.LookupParameter("Name").AsString()))
        
        #print("Sorted_Grids: {}".format(len(sorted_grids)))
        #print("Sorted_Groups: {}".format(len(sorted_groups)))

        offset_distance1 = 1.5 * s_factor/100
        offset_distance2 = 3 * s_factor/100

        for k in sorted_groups.keys():
            
            lt = sorted_groups[k]
            gr = change_to_2D_extents(lt)
            start_points, end_points, grid_curve = grid_points(gr, view)
            line = refLine(end_points)
            ref = refArray(lt)
            orientation = get_orientation(gr[0])
            #print(orientation)

            if view.ViewType in (ViewType.Elevation, ViewType.Section ):
                offset_distance1 = -abs(offset_distance1)
                offset_distance2 = -abs(offset_distance2)

                line_offset1 = offset_line_elev(line, offset_distance1)
                line_offset2 = offset_line_elev(line, offset_distance2) 
                
            if view.ViewType in (ViewType.FloorPlan, ViewType.AreaPlan, ViewType.CeilingPlan):
                if orientation == "vertical":
                    offset_distance1 = abs(offset_distance1)
                    offset_distance2 =  abs(offset_distance2)
                if orientation == "horizontal":
                    offset_distance1 = -abs(offset_distance1)
                    offset_distance2 = -abs(offset_distance2)

                line_offset1 = offset_line_plan(line, offset_distance1)
                line_offset2 = offset_line_plan(line, offset_distance2) 

            #print("Offset1 : {}".format(offset_distance1))
            #print("Offset2 : {}".format(offset_distance2))
                
            n = len(lt) - 1
            end_grids = [lt[0], lt[n]]
            ref_overall = refArray(end_grids)

            d1 = doc.Create.NewDimension(view, line_offset2, ref)
            d2 = doc.Create.NewDimension(view, line_offset1, ref_overall)

            if not d1 or not d2:
                view_issue_counts[view_key]["DimensionError"] += 1
                grid_list_length -= len(gr)
        view_list_length+= 1                   
    total_grid_count += grid_list_length
    total_view_count += view_list_length

    #print("Successfully added dimensions to {} grids in {} views".format(total_grid_count, total_view_count))

    # Commit transaction
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

for view_key, data in view_issue_counts.iteritems(): 
    issue_description = []
    if data["DimensionError"] > 0:
        issue_description.append("DIMENSION ERROR")
    if issue_description:
        failed_room_data = [
            view_key,
            data["Name"],
            data["ViewType"],
            ", ".join(issue_description)
        ]
        failed_data.append(failed_room_data)

if failed_data:
    output.print_md(header_data)
    processed_grids = total_grid_count 
    processed_views = total_view_count - len(failed_data)
    output.print_md("##⚠️ Completed. {} grids in {} views dimensioned. {} Issues Found ".format(processed_grids, processed_views, len(failed_data)))
    output.print_table(table_data=failed_data, columns=["ELEMENT ID", "VIEW NAME", "VIEW TYPE", "EXCEPTION"]) 
    print("\n\n")
    output.print_md("---") 
    output.print_md("***❌ EXCEPTION REFERENCE***")
    output.print_md("---")
    output.print_md("**DIMENSION ERROR** - Dimension could not be created. Please check the view manually.\n")

elif total_grid_count > 0 and not failed_data:
    forms.alert("Successfully aligned & added dimensions to {} grids in {} views".format(total_grid_count, total_view_count))

elif total_grid_count > 0 and failed_data:
    output.print_md("##✅ {} Completed with issues.".format(__title__))


#print("Script runtime: {:.2f} seconds".format(runtime))


