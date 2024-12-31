'''Spot Elevations'''
__title__ = "Spot Elevations"
__author__ = "prajwalbkumar"

# Imports
import clr
import os
import time
import xlrd

clr.AddReference("RevitAPI")
from pyrevit import revit, forms, script
from datetime import datetime
from Extract.RunData import get_run_data
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.Creation import *
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

view = doc.ActiveView
view_level_id = view.GenLevel.Id

all_floors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(view_level_id)).ToElements()

if not all_floors:
    script.exit()

t = Transaction(doc, "Spot Elevation")

t.Start()
for floor in all_floors:
    floor_grade = round(floor.LookupParameter("Height Offset From Level").AsDouble(), 2)
    if floor.LookupParameter("Area").AsDouble() < 40:
        continue

    floor_bbox = floor.get_BoundingBox(view)
    if floor_bbox is None:
        continue

    floor_outline = Outline(floor_bbox.Min, floor_bbox.Max)
    # Inflate the bounding box manually by scaling the min and max points
    scale_factor = 1.5  # Adjust scale factor as needed
    bbox_min = floor_bbox.Min
    bbox_max = floor_bbox.Max
    inflated_min = XYZ(
        bbox_min.X - (bbox_max.X - bbox_min.X) * (scale_factor - 1) / 2,
        bbox_min.Y - (bbox_max.Y - bbox_min.Y) * (scale_factor - 1) / 2,
        bbox_min.Z - (bbox_max.Z - bbox_min.Z) * (scale_factor * 10 - 1) / 2,
    )
    inflated_max = XYZ(
        bbox_max.X + (bbox_max.X - bbox_min.X) * (scale_factor - 1) / 2,
        bbox_max.Y + (bbox_max.Y - bbox_min.Y) * (scale_factor - 1) / 2,
        bbox_max.Z + (bbox_max.Z - bbox_min.Z) * (scale_factor * 10 - 1) / 2,
    )

    floor_inflated_outline = Outline(inflated_min, inflated_max)

    # adj_floors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(LogicalAndFilter(ElementLevelFilter(view_level_id), BoundingBoxIntersectsFilter(floor_inflated_outline))).ToElements()


    # filtered_by_level = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(view_level_id)).ToElements()
    # filtered_by_bbox = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(BoundingBoxIntersectsFilter(floor_inflated_outline)).ToElements()

    # # Combine manually for now
    # adj_floors = filtered_by_bbox

    # # print("Filtered by Level: {} floors".format(len(filtered_by_level)))
    # # print("Filtered by Bounding Box: {} floors".format(len(filtered_by_bbox)))
    # # print("Final Adjacent Floors: {} floors".format(len(adj_floors)))
    # print("New Floor")

    # for floor in filtered_by_bbox:
    #     print(floor.Id)

    # print("-"*50)

    # for floor in filtered_by_level:
    #     print(floor.Id)

    # print("-" * 100)

    # continue

    for adj_floor in all_floors:
        adj_floor_bbox = adj_floor.get_BoundingBox(view)
        if adj_floor_bbox is None:
            continue
        adj_floor_outline = Outline(adj_floor_bbox.Min, adj_floor_bbox.Max)
        if floor_inflated_outline.Intersects(adj_floor_outline, 0.1):
            # print("{} - Intersects - {}" .format(output.linkify(floor.Id), output.linkify(adj_floor.Id)))
            # print("Yes")
            if round(adj_floor.LookupParameter("Height Offset From Level").AsDouble(), 2) == floor_grade:
                continue
            else:
                options = Options()
                options.View = view
                options.IncludeNonVisibleObjects = True
                options.ComputeReferences = True

                geometry_faces = floor.get_Geometry(options)
                if not geometry_faces:
                    print("No geometry for floor:", floor.Id)
                    continue

                reference = None
                point = None

                for geom_obj in geometry_faces:
                    if hasattr(geom_obj, "Faces"):
                        solid = geom_obj
                        for face in solid.Faces:
                            if face.FaceNormal.Z == 1 and face.Reference:
                                reference = face.Reference
                                # point = face.Evaluate(UV(0.5, 0.5))
                                face_bbox = face.GetBoundingBox()
                                uv_point = (face_bbox.Min + face_bbox.Max)/2
                                point = face.Evaluate(uv_point)
                                break
                    if reference:
                        break

                # Ensure point aligns with the reference
                if reference is not None and point is not None:
                    projected_point = face.Project(point)
                    if projected_point and projected_point.XYZPoint:
                        point = projected_point.XYZPoint
                    else:
                        # print("Point does not project properly on reference. Skipping...")
                        continue

                    # print(f"Creating SpotElevation for Floor {floor.Id}")
                    doc.Create.NewSpotElevation(view, reference, point, point, point, point, False)
                # else:
                #     # print(f"No valid reference or point for Floor {floor.Id}. Skipping...")
t.Commit()
