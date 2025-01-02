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
view_Scale = str(view.Scale)

linked_instance = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
if linked_instance:
    documentation_file = forms.alert("Is this a Documentation File or a Live File", warn_icon=False, options=["Documentation File", "Live File"])

    if not documentation_file:
        forms.alert("No file option selected. Exiting script.", exitscript=True)

    if documentation_file == "Documentation File":
        link_name = []
        for link in linked_instance:
            link_name.append(link.Name)

        # Collect all AI Floor Finishes
        ai_instance_name = forms.SelectFromList.show(link_name, title = "Select Linked AI File Containing Floor Finishes", width=600, height=600, button_name="Select File", multiselect=False)
        if not ai_instance_name:
            script.exit()

        for link in linked_instance:
            if ai_instance_name == link.Name:
                ai_instance = link
                break
        

        ai_doc = ai_instance.GetLinkDocument()
        if not ai_doc:
            forms.alert("No instance found of the selected link.\n"
                        "Use Manage Links to Load the Link in the active document!", title = "Link Missing", warn_icon = True)
            script.exit()

        views_in_ai_link = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()

        for link_view in views_in_ai_link:
            if link_view.Name == "3D-Navisworks-Export":
                geom_view = link_view
                break

        
        live_view_level_name = doc.GetElement(view.GenLevel.Id).Name
        
        link_levels = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

        for level in link_levels:
            if level.Name == live_view_level_name:
                ai_view_level_id = level.Id
                break
        
        ai_floor_finishes = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(ai_view_level_id)).ToElements()

        ######################################################
        # Collect all ST / SC Floors
        st_instance_name = forms.SelectFromList.show(link_name, title = "Select Linked ST File Containing Floor Slabs", width=600, height=600, button_name="Select File", multiselect=False)
        if not st_instance_name:
            script.exit()

        for link in linked_instance:
            if st_instance_name == link.Name:
                st_instance = link
                break
        

        st_doc = st_instance.GetLinkDocument()
        if not st_doc:
            forms.alert("No instance found of the selected link.\n"
                        "Use Manage Links to Load the Link in the active document!", title = "Link Missing", warn_icon = True)
            script.exit()

        views_in_st_link = FilteredElementCollector(st_doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()

        for link_view in views_in_st_link:
            if link_view.Name == "3D-Navisworks-Export":
                geom_view = link_view
                break

        
        link_levels = FilteredElementCollector(st_doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
        live_view_level_elevation = doc.GetElement(view.GenLevel.Id).Elevation
        for level in link_levels:
            if level.Elevation < live_view_level_elevation and level.Elevation > live_view_level_elevation - 2:
                st_view_level_id = level.Id
                break
        
        st_floor_finishes = FilteredElementCollector(st_doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(st_view_level_id)).ToElements()


        ######################################################
        # Collect all AR Rooms
        ar_instance_name = forms.SelectFromList.show(link_name, title = "Select Linked AR File Containing Rooms", width=600, height=600, button_name="Select File", multiselect=False)
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

        
        link_levels = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

        for level in link_levels:
            if level.Name == live_view_level_name:
                ar_view_level_id = level.Id
                break


        ar_rooms = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
        filtered_ar_rooms = [room for room in ar_rooms if room.LevelId == ar_view_level_id]

    else:
        linked_instance = None
        ai_doc = doc
        view_level_id = view.GenLevel.Id
        ai_floor_finishes = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(view_level_id)).ToElements()


else:
    linked_instance = None
    ai_doc = doc
    view_level_id = view.GenLevel.Id
    ai_floor_finishes = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(view_level_id)).ToElements()


if not ai_floor_finishes:
    script.exit()

if not st_floor_finishes:
    script.exit()

if not ar_rooms:
    script.exit()

finish_candidates = []

spot_dimension_collector = FilteredElementCollector(doc).OfClass(SpotDimensionType).WhereElementIsElementType()
spot_dimension_type_names = [spot_dimension.LookupParameter("Type Name").AsValueString() for spot_dimension in spot_dimension_collector]

ffl_type_names = [name for name in spot_dimension_type_names if "FFL" in name and view_Scale in name]
cl_type_names = [name for name in spot_dimension_type_names if "CL" in name and view_Scale in name]

ffl_spot_dimension_name = forms.SelectFromList.show(ffl_type_names, title = "Select FFL Tag Type", width=600, height=600, button_name="Select Tag Type", multiselect=False)
cl_spot_dimension_name = forms.SelectFromList.show(cl_type_names, title = "Select CL Tag Type", width=600, height=600, button_name="Select Tag Type", multiselect=False)

ffl_spot_dimension_type = [type for type in spot_dimension_collector if type.LookupParameter("Type Name").AsValueString() == ffl_spot_dimension_name]
cl_spot_dimension_type = [type for type in spot_dimension_collector if type.LookupParameter("Type Name").AsValueString() == cl_spot_dimension_name]

ffl_spot_dimension_type = ffl_spot_dimension_type[0]
cl_spot_dimension_type = cl_spot_dimension_type[0]

consumed_origins = []

t = Transaction(doc, "Spot Elevation")
t.Start()
for floor in ai_floor_finishes:
    floor_grade = round(floor.LookupParameter("Height Offset From Level").AsDouble(), 2)
    if floor.LookupParameter("Area").AsDouble() < 40:
        continue

    options = Options()
    options.View = view
    options.IncludeNonVisibleObjects = True
    options.ComputeReferences = True

    geometry_faces = floor.get_Geometry(options)
    if not geometry_faces:
        # print("No geometry for floor:", floor.Id)
        continue

    reference = None
    point = None

    for geom_obj in geometry_faces:
        if hasattr(geom_obj, "Faces"):
            solid = geom_obj
            for face in solid.Faces:
                try:
                    if face.FaceNormal.Z == 1 and face.Reference:
                        finish_candidates.append(face)
                except:
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

    for adj_floor in ai_floor_finishes:
        adj_floor_bbox = adj_floor.get_BoundingBox(view)
        if adj_floor_bbox is None:
            continue
        adj_floor_outline = Outline(adj_floor_bbox.Min, adj_floor_bbox.Max)
        if floor_inflated_outline.Intersects(adj_floor_outline, 0.1):
            if round(adj_floor.LookupParameter("Height Offset From Level").AsDouble(), 2) == floor_grade:
                continue
            else:
                geometry_faces = floor.get_Geometry(options)
                if not geometry_faces:
                    # print("No geometry for floor:", floor.Id)
                    continue

                reference = None
                point = None

                for geom_obj in geometry_faces:
                    if hasattr(geom_obj, "Faces"):
                        solid = geom_obj
                        for face in solid.Faces:
                            try:
                                if face.FaceNormal.Z == 1 and face.Reference:
                                    finish_candidates.append(face)
                                    if linked_instance:
                                        reference = face.Reference.CreateLinkReference(ai_instance)
                                    else:
                                        reference = face.Reference
                                        
                                    face_bbox = face.GetBoundingBox()
                                    uv_point = (face_bbox.Min + face_bbox.Max)/2
                                    point = face.Evaluate(uv_point)
                                    break
                            except:
                                continue

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
                        consumed_origins.append(point)
                        ffl_spot_elevation = doc.Create.NewSpotElevation(view, reference, point, point, point, point, False)
                        ffl_spot_elevation.DimensionType = ffl_spot_dimension_type

                    except:
                        # print("Failed for Floor {}" .format(floor.Id))
                        continue
            break

for room in ar_rooms:
    # print("------")

    room_finish = False
    room_point = room.Location.Point

    room_name = (room.LookupParameter("Name").AsValueString()).lower()
    # print(room_name)
    # print("-Processing-")


    # z_room_point = XYZ(room_point.X, room_point.Y, room_point.Z + 1)
    for face in finish_candidates:
        try:
            projecton = face.Project(room_point)
            if projecton and projecton.XYZPoint: # If valid projection is returned, then the room contains a finish. break the loop and go to other room
                room_finish = True
                break
        except:
            continue
            # print("None - None")
    
    else: 
        for floor in st_floor_finishes:
            options = Options()
            options.View = view
            options.IncludeNonVisibleObjects = True
            options.ComputeReferences = True

            geometry_faces = floor.get_Geometry(options)
            if not geometry_faces:
                # print("No geometry for floor:", floor.Id)
                continue
            
            face = None
            reference = None
            for geom_obj in geometry_faces:
                if hasattr(geom_obj, "Faces"):
                    solid = geom_obj
                    for face in solid.Faces:
                        if face.FaceNormal.Z == 1 and face.Reference:
                            if linked_instance:
                                reference = face.Reference.CreateLinkReference(st_instance)
                            else:
                                reference = face.Reference
                            break
                if reference:
                    break

            # Ensure point aligns with the reference
            if reference is not None:
                projected_point = face.Project(room_point)
                if projected_point and projected_point.XYZPoint:
                    point = projected_point.XYZPoint
                    # print(point)
                else:
                    # print("Point does not project properly on reference. Skipping...")
                    continue

                try:
                    cl_spot_elevation = doc.Create.NewSpotElevation(view, reference, point, point, point, point, False)
                    cl_spot_elevation.SpotDimensionType = cl_spot_dimension_type
                    break
                except:
                    # print("Failed for Floor {}" .format(floor.Id))
                    continue

t.Commit()