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

def calculate_triangle_area(pt1, pt2, pt3):
    AB = pt2 - pt1
    AC = pt3 - pt1
    cross_product = AB.CrossProduct(AC)
    return 0.5 * cross_product.GetLength()
    
def triangulate_point(face):
    triangulated_mesh = face.Triangulate()
    triangle_count = triangulated_mesh.NumTriangles
    largest_area = 0
    for i in range(triangle_count):
        triangle = triangulated_mesh.get_Triangle(i)
        pt1 = triangle.get_Vertex(0)
        pt2 = triangle.get_Vertex(1)
        pt3 = triangle.get_Vertex(2)
        area = calculate_triangle_area(pt1, pt2, pt3)
        if area > largest_area:
            centroid = (pt1 + pt2 + pt3) / 3

    centroid = XYZ(centroid.X, centroid.Y - 6, centroid.Z)

    return centroid

def get_faces(solid):
    faces = []
    all_faces = solid.Faces
    for face in all_faces:
        faces.append(face)
    return faces

def get_upper_faces(stair, stair_geometry):
    upper_faces = []
    if (stair.LookupParameter("Family").AsValueString() == "Assembled Stair"):
        for components in stair_geometry:
            for geometry in components:
                if (geometry.ToString() == "Autodesk.Revit.DB.GeometryInstance"):
                    geometry_instance = geometry.GetInstanceGeometry()
                    for solid in geometry_instance:
                        faces = get_faces(solid)
                        if faces:
                            for face in faces:
                                vector_z = int(face.FaceNormal.Z)
                                if vector_z == 1:
                                    upper_faces.append(face)
        return upper_faces

    else:
        for component in stair_geometry:
            for geometry in component:
                if (geometry.ToString() == "Autodesk.Revit.DB.Solid"):
                    # print(geometry.Volume) ## This is a solid as well
                    faces = get_faces(geometry)
                    if faces:
                            for face in faces:
                                vector_z = int(face.FaceNormal.Z)
                                if vector_z == 1:
                                    upper_faces.append(face)
        return upper_faces
     

# Define all possible view types for selection
all_view_types = {
    'Floor Plans': ViewType.FloorPlan,
    'Reflected Ceiling Plans': ViewType.CeilingPlan,
    'Section' : ViewType.Section
}

# Show a dialog box for selecting desired view types
selected_view_type_names = forms.SelectFromList.show(
    sorted(all_view_types.keys()),            
    title='Select View Types',                
    multiselect=True                           
)
if not selected_view_type_names:
    forms.alert("No view type was selected. Exiting script.", exitscript=True)
selected_view_types = [all_view_types[name]
                        for name in selected_view_type_names
                        if name in all_view_types]

# Collect all views in the document
views_collector = FilteredElementCollector(doc).OfClass(View)

# Filter views by the selected types
filtered_views = [view for view in views_collector
                    if view.ViewType in selected_view_types
                    and not view.IsTemplate]
if not filtered_views:
    forms.alert("No views of the selected types were found in the document.", exitscript=True)

# Collect all Viewport elements in the document
viewports_collector = FilteredElementCollector(doc).OfClass(Viewport)
views_on_sheets_ids = {viewport.ViewId for viewport in viewports_collector}
filtered_views_on_sheets = [
    view for view in filtered_views
    if view.Id in views_on_sheets_ids
]
if not filtered_views_on_sheets:
    forms.alert("No views of the selected types were found on sheets in the document.", exitscript=True)

view_dict = {view.Name: view for view in filtered_views_on_sheets}

# Show the selection window for the user to choose views
selected_view_names = forms.SelectFromList.show(
    sorted(view_dict.keys()),         
    title='Select Views',             
    multiselect=True                  
)

if selected_view_names:
    selected_views = [view_dict[name] for name in selected_view_names]
else:
    forms.alert("No views were selected.", exitscript=True)


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
            
        # AI LINK DETAILS
        ai_instance_name = forms.SelectFromList.show(link_name, title = "Select Linked AI File Containing Ceiling", width=600, height=600, button_name="Select File", multiselect=False)
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

        ai_link_levels = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()


        # ST LINK DETAILS
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
            
        st_link_levels = FilteredElementCollector(st_doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

        # AR LINK DETAILS
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

        ar_link_levels = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

spot_dimension_collector = FilteredElementCollector(doc).OfClass(SpotDimensionType).WhereElementIsElementType()
spot_dimension_type_names = [spot_dimension.LookupParameter("Type Name").AsValueString() for spot_dimension in spot_dimension_collector]

ffl_type_names = [name for name in spot_dimension_type_names if "FFL" in name]
cl_type_names = [name for name in spot_dimension_type_names if "CL" in name]

ffl_spot_dimension_name = forms.SelectFromList.show(ffl_type_names, title = "Select FFL Tag Type", width=600, height=600, button_name="Select Tag Type", multiselect=False)
cl_spot_dimension_name = forms.SelectFromList.show(cl_type_names, title = "Select CL Tag Type", width=600, height=600, button_name="Select Tag Type", multiselect=False)

ffl_spot_dimension_type = [type for type in spot_dimension_collector if type.LookupParameter("Type Name").AsValueString() == ffl_spot_dimension_name]
cl_spot_dimension_type = [type for type in spot_dimension_collector if type.LookupParameter("Type Name").AsValueString() == cl_spot_dimension_name]

ffl_spot_dimension_type = ffl_spot_dimension_type[0]
cl_spot_dimension_type = cl_spot_dimension_type[0]


t = Transaction(doc, "Spot Elevation")
t.Start()

for view in selected_views:

    view_scale = str(view.Scale)
    view_type = view.ViewType

    options = Options()
    options.View = view
    options.IncludeNonVisibleObjects = True
    options.ComputeReferences = True



    if linked_instance:
        if documentation_file == "Documentation File":
            if view_type == ViewType.CeilingPlan:

                live_view_level_name = doc.GetElement(view.GenLevel.Id).Name
                live_view_level_elevation = doc.GetElement(view.GenLevel.Id).Elevation
                
                # Filtering AI Ceiling Finishes

                for level in ai_link_levels:
                    if level.Name == live_view_level_name:
                        ai_view_level_id = level.Id
                        break
                
                ai_ceiling_finishes = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Ceilings).WherePasses(ElementLevelFilter(ai_view_level_id)).ToElements()

            if view_type == ViewType.FloorPlan:

                live_view_level_name = doc.GetElement(view.GenLevel.Id).Name
                live_view_level_elevation = doc.GetElement(view.GenLevel.Id).Elevation

                # Collect all AI Floor Finishes 
                for level in ai_link_levels:
                    if level.Name == live_view_level_name:
                        ai_view_level_id = level.Id
                        break
                
                ai_floor_finishes = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(ai_view_level_id)).ToElements()

                # Collect all ST / SC Floors
                for level in st_link_levels:
                    if level.Elevation < live_view_level_elevation and level.Elevation > live_view_level_elevation - 2:
                        st_view_level_id = level.Id
                        break
                
                st_floors = FilteredElementCollector(st_doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(st_view_level_id)).ToElements()
                
                st_floor_finishes = [floor for floor in st_floors if not "PT" in floor.Name.upper()]
                pt_floor_finishes = [floor for floor in st_floors if "PT" in floor.Name.upper()]

        
                # Collect all AR Rooms
                for level in ar_link_levels:
                    if level.Name == live_view_level_name:
                        ar_view_level_id = level.Id
                        break

                ar_rooms = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
                filtered_ar_rooms = [room for room in ar_rooms if room.LevelId == ar_view_level_id]

            if view_type == ViewType.Section:
                view_crop_loop = view.GetCropRegionShapeManager().GetCropShape()
                solids = []

                for loop in view_crop_loop:
                    solid = GeometryCreationUtilities.CreateExtrusionGeometry([loop], view.ViewDirection, 1)

                    solids.append(solid)

                    
                ai_floor_finishes = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements()
                filtered_ai_floor_finishes = intersecting_geometries(ai_floor_finishes, options)
            
                ai_ceiling_finishes = FilteredElementCollector(ai_doc).OfCategory(BuiltInCategory.OST_Ceilings).ToElements()
                filtered_ai_ceiling_finishes = intersecting_geometries(ai_ceiling_finishes, options)
                
                st_floors = FilteredElementCollector(st_doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements()
                intersecting_st_floors = intersecting_geometries(st_floors, options)

                filtered_st_floors = [floor for floor in intersecting_st_floors if not "PT" in floor.Name.upper()]
                filtered_pt_floors = [floor for floor in intersecting_st_floors if "PT" in floor.Name.upper()]

                # Collect all AR Rooms

                ar_rooms = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
                filtered_ar_rooms = intersecting_geometries(ar_rooms, options)

        else:
            linked_instance = None
            ai_doc = doc
            ar_doc = doc
            view_level_id = view.GenLevel.Id
            ai_floor_finishes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(view_level_id)).ToElements()
            stair_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Stairs).WhereElementIsNotElementType().ToElements()
            filtered_stairs = [stair for stair in stair_collector if stair.LookupParameter("Base Level").AsElementId() == view_level_id] 

    else:
        linked_instance = None
        ai_doc = doc
        view_level_id = view.GenLevel.Id
        ai_floor_finishes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WherePasses(ElementLevelFilter(view_level_id)).ToElements()


    if view_type == ViewType.FloorPlan:

        finish_candidates = []

        consumed_origins = []

        options = Options()
        options.View = view
        options.IncludeNonVisibleObjects = True
        options.ComputeReferences = True
        
        # for stair in filtered_stairs:
        #     # Extract Landing Faces
        #     stair_geometry = []
        #     stair_landing_ids = stair.GetStairsLandings()
        #     landing_faces = []
        #     for landing_id in stair_landing_ids:
        #         landing = ar_doc.GetElement(landing_id)
        #         stair_geometry.append(landing.get_Geometry(options))

        #     landing_faces = get_upper_faces(stair, stair_geometry)
            
        #     for face in landing_faces:
        #         reference = Reference(stair)
        #         # if linked_instance:
        #         #     reference = face.Reference.CreateLinkReference(ar_instance)
        #         # else:
        #         #     reference = face.Reference
        #         print(reference.ElementReferenceType)
                    
        #         face_bbox = face.GetBoundingBox()
        #         point = triangulate_point(face)
        #         point = XYZ(point.X, point.Y + 6, point.Z)
                


        #         # Ensure point aligns with the reference
        #         if reference is not None and point is not None:
        #             projected_point = face.Project(point)
        #             if projected_point and projected_point.XYZPoint:
        #                 point = projected_point.XYZPoint
        #                 print(point)
        #             else:
        #                 print("Point does not project properly on reference. Skipping...")
        #                 continue

        #             # print(f"Creating SpotElevation for Floor {floor.Id}")
        #             # try:
        #             print("Spotting")
        #             ffl_spot_elevation = doc.Create.NewSpotElevation(view, reference, point, point, point, point, False)
        #             ffl_spot_elevation.DimensionType = ffl_spot_dimension_type

        #             # except:
        #             #     print("Failed for Floor")
        #             #     continue
        

        for floor in ai_floor_finishes:
            floor_grade = round(floor.LookupParameter("Height Offset From Level").AsDouble(), 2)
            if floor.LookupParameter("Area").AsDouble() < 40:
                continue


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
                                            # uv_point = (face_bbox.Min + face_bbox.Max)/2
                                            # point = face.Evaluate(uv_point)
                                            point = triangulate_point(face)
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

        if linked_instance:

            for room in filtered_ar_rooms:
                # print("------")

                room_finish = False
                room_point = room.Location.Point

                room_name = (room.LookupParameter("Name").AsValueString()).lower()

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

            for floor in pt_floor_finishes:

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
                                    if linked_instance:
                                        reference = face.Reference.CreateLinkReference(st_instance)
                                    else:
                                        reference = face.Reference
                                        
                                    face_bbox = face.GetBoundingBox()
                                    # uv_point = (face_bbox.Min + face_bbox.Max)/2
                                    # point = face.Evaluate(uv_point)
                                    point = triangulate_point(face)
                                    break
                            except:
                                print("error")
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
                        cl_spot_elevation = doc.Create.NewSpotElevation(view, reference, point, point, point, point, False)
                        cl_spot_elevation.DimensionType = cl_spot_dimension_type

                    except:
                        print("Failed for Floor {}" .format(floor.Id))
                        continue
        
    if view_type == ViewType.CeilingPlan:

        if not ai_ceiling_finishes:
            script.exit()

        for ceiling in ai_ceiling_finishes:

            options = Options()
            options.View = view
            options.IncludeNonVisibleObjects = True
            options.ComputeReferences = True

            geometry_faces = ceiling.get_Geometry(options)
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
                                if linked_instance:
                                    reference = face.Reference.CreateLinkReference(ai_instance)
                                else:
                                    reference = face.Reference
                                    
                                face_bbox = face.GetBoundingBox()
                                # uv_point = (face_bbox.Min + face_bbox.Max)/2
                                # point = face.Evaluate(uv_point)

                                point = triangulate_point(face)
                                # print(point)

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
                    ceiling_spot_elevation = doc.Create.NewSpotElevation(view, reference, point, point, point, point, False)
                    ceiling_spot_elevation.DimensionType = ffl_spot_dimension_type

                except:
                    print("Failed for Floor {}" .format(floor.Id))
                    continue

            else:
                print("failed ceiling reference")

    if view_type == ViewType.Section:

        all_floors = []
        all_floors += filtered_ai_floor_finishes
        all_floors += filtered_ai_ceiling_finishes
        all_floors += filtered_pt_floors
        all_floors += filtered_st_floors

        def spot_elevation(slabs, instance):
            for slab in slabs:
                if not slab.LookupParameter("Height Offset From Level").AsDouble() == 0:
                    options = Options()
                    options.View = view
                    options.IncludeNonVisibleObjects = True
                    options.ComputeReferences = True

                    geometry_faces = slab.get_Geometry(options)
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
                                        edgeloops = face.EdgeLoops
                                        for loop in edgeloops:
                                            for edge in loop:
                                                # if edge.AsCurve().Direction.IsAlmostEqualTo(view.RightDirection):
                                                # if edge.AsCurve().Direction.IsAlmostEqualTo(XYZ(0,0,1).CrossProduct(view.RightDirection)):
                                                if edge.AsCurve().Direction.IsAlmostEqualTo(view.RightDirection.CrossProduct(XYZ(0,0,1))):
                                                    if linked_instance:
                                                        reference = edge.Reference.CreateLinkReference(instance)
                                                        point = edge.AsCurve().GetEndPoint(1) - edge.AsCurve().GetEndPoint(0)
                                                        # print(edge.AsCurve().Direction)
                                                        # print(view.RightDirection)
                                                        break
                                        break
                                except:
                                    continue

                        if reference:
                            break

                    # Ensure point aligns with the reference
                    # try:
                    bendpt = point + XYZ(0,0,5)
                    endpt = bendpt + view.RightDirection * 3
                    ceiling_spot_elevation = doc.Create.NewSpotElevation(view, reference, point, point, point, point, False)
                    # print(point, bendpt, endpt)
                    # ceiling_spot_elevation.DimensionType = ffl_spot_dimension_type

                    # except:
                    #     print("Failed for Floor {}" .format(floor.Id))
                    #     continue

        spot_elevation(filtered_ai_floor_finishes, ai_instance)
        spot_elevation(filtered_ai_ceiling_finishes, ai_instance)
        spot_elevation(filtered_pt_floors, st_instance)
        spot_elevation(filtered_st_floors, st_instance)

t.Commit()