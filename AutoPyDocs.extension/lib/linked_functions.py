# -*- coding: utf-8 -*-
# # Imports
import time
from Autodesk.Revit.DB  import *
from Extract.RunData    import get_run_data

def align_grids(doc, selected_views):

    def new_point(exisiting_point, direction, bbox_curves, start_point = None):

        projected_points = []
        for curve in bbox_curves:
            project = curve.Project(exisiting_point).XYZPoint
            projected_points.append(XYZ(project.X, project.Y, 0))

        possible_points = []

        exisiting_point = XYZ(exisiting_point.X, exisiting_point.Y, 0)

        try:
            for point in projected_points:
                if (exisiting_point - point).Normalize().IsAlmostEqualTo(direction) or (exisiting_point - point).Normalize().IsAlmostEqualTo(direction.Negate()):
                    possible_points.append(point)

            if start_point:
                if start_point.IsAlmostEqualTo(possible_points[0]):
                        new_point = possible_points[1]
                else:
                    new_point = possible_points[0]

            else:
                if point.DistanceTo(possible_points[0]) > point.DistanceTo(possible_points[1]):
                    new_point = possible_points[0]
                    
                else:
                    new_point = possible_points[1]
        
            return new_point

        except:
            return exisiting_point

    try:
        t = Transaction(doc, "Allign Grids")
        t.Start()

        start_time = time.time()

        manual_time = 108.4 

        total_grid_count = 0
        total_view_count = 0
        view_list_length =0    

        for view in selected_views:
            
            grids_collector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
        
            bbox = view.CropBox
        
            corner1 = XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z)
            corner2 = XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z)
            corner3 = XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z)
            corner4 = XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z)

            # Create lines representing the bounding box edges
            line1 = Line.CreateBound(corner1, corner2)
            line2 = Line.CreateBound(corner2, corner3)
            line3 = Line.CreateBound(corner3, corner4)
            line4 = Line.CreateBound(corner4, corner1)

            # Create model curves in the active view
            bbox_curves = [line1, line2, line3, line4]
            
            floor_plan_views = ["FloorPlan", "CeilingPlan", "EngineeringPlan", "AreaPlan"]
            front_views = ["Elevation", "Section"]

            if str(view.ViewType) in floor_plan_views:

                # Number of Grids
                grid_list_length = 0

                # Convert all Grids to ViewSpecific Grids
                for grid in grids_collector:
                    grid.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.ViewSpecific)
                    grid.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.ViewSpecific)

                    # Get the curves of the grids
                    curves = grid.GetCurvesInView(DatumExtentType.ViewSpecific, view)
                    for curve in curves:
                        grids_view_curve = curve
                        # point = curve.GetEndPoint(0)
                        # plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, point)
                        # sketch_plane = SketchPlane.Create(doc, plane)
                        # model_line = doc.Create.NewModelCurve(Line.CreateBound(point, curve.GetEndPoint(1)), sketch_plane)
                    
                    start_point = grids_view_curve.GetEndPoint(0)
                    end_point = grids_view_curve.GetEndPoint(1)
                    direction = (end_point - start_point).Normalize()
                
                    new_start_point = new_point(start_point, direction, bbox_curves)
                    new_end_point = new_point(end_point, direction, bbox_curves, new_start_point)

                    new_start_point = XYZ(new_start_point.X, new_start_point.Y, start_point.Z)
                    new_end_point = XYZ(new_end_point.X, new_end_point.Y, end_point.Z)

                    new_grid_line = Line.CreateBound(new_start_point, new_end_point)

                    if not int(new_start_point.X) == int(corner2.X):
                        new_grid_line = new_grid_line.CreateReversed()
                    
                    if not int(new_start_point.Y) == int(corner3.Y):
                        new_grid_line = new_grid_line.CreateReversed()


                    grid.SetCurveInView(DatumExtentType.ViewSpecific, view, new_grid_line)
                    grid.HideBubbleInView(DatumEnds.End0, view)
                    # grid.HideBubbleInView(DatumEnds.End1, view)
                    grid.ShowBubbleInView(DatumEnds.End1, view)

                    grid_list_length += 1

                total_grid_count += grid_list_length

                    # plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, new_start_point)
                    # sketch_plane = SketchPlane.Create(doc, plane)
                    # model_line = doc.Create.NewModelCurve(new_grid_line, sketch_plane)

            elif str(view.ViewType) in front_views:

                # Initialize an empty list to store Z-values from crop region curves
                end_point_z = []

                # Retrieve the crop region 
                crop_manager = view.GetCropRegionShapeManager()
                crop_region = crop_manager.GetCropShape()

                # Explode curve loop into individual lines
                for curve_loop in crop_region:
                    for curve in curve_loop:

                        # Get Z value of all end points
                        end_point_z.append(curve.GetEndPoint(0).Z)

                #Sort Z value of end points
                end_point_z = sorted(end_point_z)
            
                # Number of Grids
                grid_list_length = 0

                # Convert all Grids to ViewSpecific Grids
                for grid in grids_collector:
                    grid.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.ViewSpecific)
                    grid.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.ViewSpecific)
            
                    # Get the curves of the grids
                    curves = grid.GetCurvesInView(DatumExtentType.ViewSpecific, view)
                    for curve in curves:
                        grids_view_curve = curve

                    # Get start and end points of the grid curve 
                    start_point   = curve.GetEndPoint(0)
                    end_point     = curve.GetEndPoint(1)

                    # Modify the Z-values of the start and end points based on the crop region
                    start_point   = XYZ(start_point.X, start_point.Y, end_point_z[0])
                    end_point     = XYZ(end_point.X, end_point.Y, end_point_z[-1])

                    # Create a new grid line with the updated start and end points
                    new_grid_line = Line.CreateBound(start_point, end_point)

                    # Apply the new grid line to the grid in the view
                    grid.SetCurveInView(DatumExtentType.ViewSpecific, view, new_grid_line)

                    # Hide the bubble at one end of the grid in this view
                    grid.HideBubbleInView(DatumEnds.End0, view)

                    # Show bubble at the other end of the grid
                    grid.ShowBubbleInView(DatumEnds.End1, view)

                    grid_list_length += 1
                total_grid_count += grid_list_length

            view_list_length+= 1
        total_view_count += view_list_length

        #print("Successfully aligned {} grids in {} views".format(total_grid_count, total_view_count))

        t.Commit()

        end_time = time.time()
        runtime = end_time - start_time
                
        run_result = "Tool ran successfully"
        if total_grid_count:
            element_count = total_grid_count
        else:
            element_count = 0

        error_occured ="Nil"

        get_run_data('Align Grids link in Dimesnion Grids', runtime, element_count, manual_time, run_result, error_occured)
        
    except Exception as e:
    
        #t.RollBack()
        print("Error occurred: {}".format(str(e)))
        
        end_time = time.time()
        runtime = end_time - start_time

        error_occured = ("Error occurred: %s", str(e))    
        run_result = "Error"
        element_count = 0
        
        get_run_data('Align Grids link in Dimesnion Grids', runtime, element_count, manual_time, run_result, error_occured)