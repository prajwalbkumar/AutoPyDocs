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

def get_level_pairs(sorted_levels, room_z):
    count = len(sorted_levels)
    value_list = [sorted_levels[i][1] for i in range(count)]
    for i in range(count):
        if room_z > value_list[i] and room_z < value_list[i+1]:
            return [sorted_levels[i-1], sorted_levels[i], sorted_levels[i+1]]

def intersecting_geometries(element_list, options):
    intersecting_elements = []
    intersection_option = SolidCurveIntersectionOptions()
    for el in element_list:
        intersection_toggle = False 
        for solid in solids:
            if intersection_toggle == True:
                break             
            geometries = el.get_Geometry(options)
            if geometries:
                for geom in geometries:
                    if intersection_toggle == True:
                        break
                    for edge in geom.Edges:
                        curve = edge.AsCurve()
                        intersection_result = solid.IntersectWithCurve(curve,intersection_option)
                        if intersection_result.SegmentCount > 0:
                            intersecting_elements.append(el)
                            intersection_toggle = True
                            break

    return intersecting_elements

view = doc.ActiveView
view_scale = str(view.Scale)
view_type = view.ViewType

if not view_type == ViewType.Section:
    forms.alert("Active view must be a Section View", title="Script Exiting")
    script.exit()

view_crop_loop = view.GetCropRegionShapeManager().GetCropShape()
solids = []

options = Options()
options.View = view
options.IncludeNonVisibleObjects = True
options.ComputeReferences = True

t = Transaction(doc, "Dimension Section")
t.Start()

for loop in view_crop_loop:
    solid = GeometryCreationUtilities.CreateExtrusionGeometry([loop], view.ViewDirection, 1)

    # # Set the category for the DirectShape (use Generic Models as an example)
    # category = ElementId(BuiltInCategory.OST_GenericModel)
    
    # # Start a transaction to create the DirectShape
    # t = Transaction(doc, "Create DirectShape")
    # t.Start()
    
    # # Create the DirectShape
    # direct_shape = DirectShape.CreateElement(doc, category)
    # direct_shape.ApplicationId = "CustomAppId"
    # direct_shape.ApplicationDataId = "CustomDataId"
    
    # # Assign the solid geometry to the DirectShape
    # direct_shape.SetShape([solid])

    # t.Commit()

    solids.append(solid)

linked_instance = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
if linked_instance:
    # documentation_file = forms.alert("Is this a Documentation File or a Live File", warn_icon=False, options=["Documentation File", "Model File"])

    # if not documentation_file:
    #     forms.alert("No file option selected. Exiting script.", exitscript=True)

    documentation_file = "Documentation File"
    
    if documentation_file == "Documentation File":
        link_name = []
        for link in linked_instance:
            link_name.append(link.Name)


        # Collect all AI Floor Finishes and Ceilings
        ai_instance_name = forms.SelectFromList.show(link_name, title = "Select Linked AI File Containing Floor Finishes and Ceilings", width=600, height=600, button_name="Select File", multiselect=False)
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
            
        ai_floor_finishes = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements()
        filtered_ai_floor_finishes = intersecting_geometries(ai_floor_finishes, options)
       
        ai_ceiling_finishes = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Ceilings).ToElements()
        filtered_ai_ceiling_finishes = intersecting_geometries(ai_ceiling_finishes, options)

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
        
        st_floors = FilteredElementCollector(st_doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements()
        intersecting_st_floors = intersecting_geometries(st_floors, options)

        
        filtered_st_floors = [floor for floor in intersecting_st_floors if not "PT" in floor.Name.upper()]
        filtered_pt_floors = [floor for floor in intersecting_st_floors if "PT" in floor.Name.upper()]

        ######################################################
        # Collect all AR Rooms
        ar_instance_name = forms.SelectFromList.show(link_name, title = "Select Linked AR File Containing Rooms and Staircases", width=600, height=600, button_name="Select File", multiselect=False)
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

        ar_rooms = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
        filtered_ar_rooms = intersecting_geometries(ar_rooms, options)

active_levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
# Create a dictionary with level name as the key and elevation as the value
active_levels = {level.Name: level.Elevation for level in active_levels}

# Sort the dictionary by elevation values
sorted_levels = sorted(active_levels.items(), key=lambda item: item[1])



# Find the center point of the room. By Level.
for room in filtered_ar_rooms: 
    location = room.Location
    location = XYZ(location.Point.X, location.Point.Y, location.Point.Z + 2)
    dim_line = Line.CreateUnbound(location, XYZ(0,0,1))
    room_level_location = location.Z
    current_cl_level, current_ffl_level, next_cl_level = get_level_pairs(sorted_levels, room_level_location)

    skip_bottom_floor = False
    ceiling_in_room = False
    top_references = ReferenceArray()
    bottom_references = ReferenceArray()


    for ceiling in filtered_ai_ceiling_finishes:
        if ceiling.LookupParameter("Level").AsValueString() == current_ffl_level[0]:
            geometry_faces = ceiling.get_Geometry(options)
            if not geometry_faces:
                # print("No geometry for floor:", floor.Id)
                continue

            for geom_obj in geometry_faces:
                if hasattr(geom_obj, "Faces"):
                    solid = geom_obj
                    for face in solid.Faces:
                        try:
                            if face.FaceNormal.Z == 1 and face.Reference:
                                if face.Project(location):
                                    ceiling_in_room = True
                                    if linked_instance:
                                        top_references.Append(face.Reference.CreateLinkReference(ai_instance))
                                    else:
                                        top_references.Append(face.Reference)

                            if face.FaceNormal.Z == -1 and face.Reference:
                                if face.Project(location):
                                    ceiling_in_room = True
                                    if linked_instance:
                                        bottom_references.Append(face.Reference.CreateLinkReference(ai_instance))
                                    else:
                                        bottom_references.Append(face.Reference)
                        except:
                            continue

    for floor in filtered_ai_floor_finishes:
        if floor.LookupParameter("Level").AsValueString() == current_ffl_level[0]:
            geometry_faces = floor.get_Geometry(options)
            if not geometry_faces:
                # print("No geometry for floor:", floor.Id)
                continue

            for geom_obj in geometry_faces:
                if hasattr(geom_obj, "Faces"):
                    solid = geom_obj
                    for face in solid.Faces:
                        try:
                            if face.FaceNormal.Z == 1 and face.Reference:
                                if face.Project(location):
                                    if linked_instance:
                                        bottom_references.Append(face.Reference.CreateLinkReference(ai_instance))
                                    else:
                                        bottom_references.Append(face.Reference)
                                    skip_bottom_floor = True
                                    break
                        except:
                            continue


    for floor in filtered_pt_floors:
        if floor.LookupParameter("Level").AsValueString() == current_cl_level[0]:
            geometry_faces = floor.get_Geometry(options)
            if not geometry_faces:
                # print("No geometry for floor:", floor.Id)
                continue

            for geom_obj in geometry_faces:
                if hasattr(geom_obj, "Faces"):
                    solid = geom_obj
                    for face in solid.Faces:
                        try:
                            if face.FaceNormal.Z == 1 and face.Reference:
                                if face.Project(location):
                                    if linked_instance:
                                        bottom_references.Append(face.Reference.CreateLinkReference(st_instance))
                                    else:
                                        bottom_references.Append(face.Reference)
                                    skip_bottom_floor = True
                                    break
                        except:
                            pass


    for floor in filtered_st_floors:
        if not skip_bottom_floor:
            if floor.LookupParameter("Level").AsValueString() == current_cl_level[0]:
                geometry_faces = floor.get_Geometry(options)
                if not geometry_faces:
                    # print("No geometry for floor:", floor.Id)
                    continue

                for geom_obj in geometry_faces:
                    if hasattr(geom_obj, "Faces"):
                        solid = geom_obj
                        for face in solid.Faces:
                            try:
                                if face.FaceNormal.Z == 1 and face.Reference:
                                    if face.Project(location):
                                        if linked_instance:
                                            bottom_references.Append(face.Reference.CreateLinkReference(st_instance))
                                        else:
                                            bottom_references.Append(face.Reference)
                                        break
                            except:
                                pass

        if floor.LookupParameter("Level").AsValueString() == next_cl_level[0]:
            geometry_faces = floor.get_Geometry(options)
            if not geometry_faces:
                # print("No geometry for floor:", floor.Id)
                continue

            for geom_obj in geometry_faces:
                if hasattr(geom_obj, "Faces"):
                    solid = geom_obj
                    for face in solid.Faces:
                        try:
                            if face.FaceNormal.Z == -1 and face.Reference:
                                if face.Project(location):
                                    
                                    if linked_instance:
                                        if ceiling_in_room:
                                            top_references.Append(face.Reference.CreateLinkReference(st_instance))
                                        else:
                                            bottom_references.Append(face.Reference.CreateLinkReference(st_instance))
                                    else:
                                        if ceiling_in_room:
                                            top_references.Append(face.Reference)
                                        else:
                                            bottom_references.Append(face.Reference)
                                    break
                        except:
                            continue



    if ceiling_in_room:
        doc.Create.NewDimension(view, dim_line, top_references)
        
    doc.Create.NewDimension(view, dim_line, bottom_references)

t.Commit()