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

linked_instance = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
if linked_instance:
    documentation_file = forms.alert("Is this a Documentation File or a Live File", warn_icon=False, options=["Documentation File", "Live File"])

    if not documentation_file:
        forms.alert("No file option selected. Exiting script.", exitscript=True)

    if documentation_file == "Documentation File":
        link_name = []
        for link in linked_instance:
            link_name.append(link.Name)

        ar_instance_name = forms.SelectFromList.show(link_name, title = "Select Linked File Containing Floors", width=600, height=600, button_name="Select File", multiselect=False)

        if not ar_instance_name:
            script.exit()

        for link in linked_instance:
            if ar_instance_name == link.Name:
                ar_instance = link
                break

        ar_doc = ar_instance.GetLinkDocument()
        if not ar_doc:
            forms.alert("No instance found of the selected link.\n"
                        "Use Manage Links to Load the Link in the active document!", title = "Link Missing", warn_icon = True)
            script.exit()

        views_in_link = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()

        for link_view in views_in_link:
            if link_view.Name == "3D-Navisworks-Export":
                geom_view = link_view
                break
        
        # TODO: else: Choose any Random 3D View
        
        live_view_level_name = doc.GetElement(view.GenLevel.Id).Name
        link_levels = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

        for level in link_levels:
            if level.Name == live_view_level_name:
                view_level_id = level.Id
                break

        all_floors = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(view_level_id)).ToElements()
    
    
    else:
        linked_instance = None
        ar_doc = doc
        view_level_id = view.GenLevel.Id
        all_floors = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(view_level_id)).ToElements()
else:
    linked_instance = None
    ar_doc = doc
    view_level_id = view.GenLevel.Id
    all_floors = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(view_level_id)).ToElements()


# Take all Rooms, Loop and find the floors that are attached to that specific Room.


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
    scale_factor = 1.3  # Adjust scale factor as needed
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
                                if linked_instance:
                                    reference = face.Reference.CreateLinkReference(ar_instance)
                                else:
                                    reference = face.Reference
                                    
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
                    try:
                        doc.Create.NewSpotElevation(view, reference, point, point, point, point, False)
                    except:
                        print("Failed for Floor {}" .format(floor.Id))
                        continue
            break
t.Commit()