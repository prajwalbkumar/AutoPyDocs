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

# Find all Floors that are hosted to the current floor. (Including Floors from the Level Pairs)


# Prompt for Selecting all floor items from Links

view = doc.ActiveView
view_level_id = view.GenLevel.Id

all_floors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(view_level_id)).ToElements()

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Outline, XYZ

# Collect all floors on a specific level
all_floors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(view_level_id)).ToElements()
if not all_floors:
    script.exit()

t = Transaction(doc, "Spot Elevation")

t.Start()
for floor in all_floors:
    floor_grade = floor.LookupParameter("Height Offset From Level").AsDouble()
    if floor.LookupParameter("Area").AsDouble() < 40:
        continue

    floor_bbox = floor.get_BoundingBox(view)
    if floor_bbox is None:
        continue

    floor_outline = Outline(floor_bbox.Min, floor_bbox.Max)
    # Inflate the bounding box manually by scaling the min and max points
    scale_factor = 1.1  # Adjust scale factor as needed
    bbox_min = floor_bbox.Min
    bbox_max = floor_bbox.Max
    inflated_min = XYZ(
        bbox_min.X - (bbox_max.X - bbox_min.X) * (scale_factor - 1) / 2,
        bbox_min.Y - (bbox_max.Y - bbox_min.Y) * (scale_factor - 1) / 2,
        bbox_min.Z  # Leave Z unscaled if only 2D inflation is required
    )
    inflated_max = XYZ(
        bbox_max.X + (bbox_max.X - bbox_min.X) * (scale_factor - 1) / 2,
        bbox_max.Y + (bbox_max.Y - bbox_min.Y) * (scale_factor - 1) / 2,
        bbox_max.Z  # Leave Z unscaled if only 2D inflation is required
    )

    floor_inflated_outline = Outline(inflated_min, inflated_max)
    adj_floors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(LogicalOrFilter(ElementLevelFilter(view_level_id), BoundingBoxIntersectsFilter(floor_inflated_outline))) 
    
    for adj_floor in adj_floors:
        if not adj_floor.LookupParameter("Height Offset From Level").AsDouble() == floor_grade:
            continue
        else:
            options = Options()
            options.View = view
            options.IncludeNonVisibleObjects = True
            options.ComputeReferences = True
            geometry_faces = floor.get_Geometry(options)

            reference = None

            for solid in geometry_faces:
                faces = solid.Faces
                for face in faces:
                    if int(face.FaceNormal.Z) == 1:
                        reference = face.Reference
                        break
                        
            point = floor_bbox.Max - floor_bbox.Min
            z_point = XYZ(point.X, point.Y, floor_bbox.Max.Z)
            print(view)
            print(reference)
            print(z_point) 

            if reference is None:
                continue
            doc.Create.NewSpotElevation(view, reference, z_point, z_point, z_point, z_point, False)
        
        break
t.Commit()