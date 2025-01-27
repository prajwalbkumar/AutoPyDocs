# -*- coding: utf-8 -*-
'''Center Room Tags'''
__title__ = "Center Room Tags"
__author__ = "abhiramnair"

# Imports
import math
import os
import time
from datetime import datetime
from Extract.RunData import get_run_data
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import WorksharingUtils
from System.Collections.Generic import List



# Record the start time
start_time = time.time()
manual_time = 45

script_dir = os.path.dirname(__file__)
ui_doc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document  # Get the Active Document
tool_name = __title__ 
model_name = doc.Title
app = __revit__.Application  # Returns the Revit Application Object
rvt_year = "Revit " + str(app.VersionNumber)
output = script.get_output()
user_name = app.Username
header_data = " | ".join([
    datetime.now().strftime("%Y-%m-%d %H:%M:%S"), rvt_year, tool_name, model_name, user_name
])


centered_tags_count = 0  # Add a counter for centered room tags

# Define all possible view types for selection
all_view_types = {
    'Floor Plans': ViewType.FloorPlan,
    'Reflected Ceiling Plans': ViewType.CeilingPlan,
    'Area Plans': ViewType.AreaPlan,
    'Structural Plans' : ViewType.EngineeringPlan,
    'Sections' : ViewType.Section
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


def move_room_and_tag(tag, room, new_pt):
        if tag.GroupId == ElementId(-1):
            tag.Location.Point = new_pt

def create_room_tag(doc, room, new_pt, view):
    tag_type = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RoomTags).WhereElementIsElementType().FirstElement()
    if tag_type:
        room_tag = IndependentTag.Create(doc, view.Id, room.Id, False, TagOrientation.Horizontal, new_pt)
        return room_tag
    return None

def calculate_triangle_area(pt1, pt2, pt3):
    AB = pt2 - pt1
    AC = pt3 - pt1
    cross_product = AB.CrossProduct(AC)
    return 0.5 * cross_product.GetLength()

def get_triangle_center(pt1, pt2, pt3):
    return XYZ((pt1.X + pt2.X + pt3.X) / 3,
                (pt1.Y + pt2.Y + pt3.Y) / 3,
                (pt1.Z + pt2.Z + pt3.Z) / 3)

def draw_triangle_lines(doc, view, pt1, pt2, pt3):
    line1 = Line.CreateBound(pt1, pt2)
    line2 = Line.CreateBound(pt2, pt3)
    line3 = Line.CreateBound(pt3, pt1)
    doc.Create.NewDetailCurve(view, line1)
    doc.Create.NewDetailCurve(view, line2)
    doc.Create.NewDetailCurve(view, line3)

def calculate_average_area(areas):
    return sum(areas) / len(areas)

def is_similar_area(areas, threshold=0.1):
    avg_area = calculate_average_area(areas)
    return all(abs(area - avg_area) / avg_area < threshold for area in areas)


def get_rooms_from_linked_models(doc):
    """Collect rooms from all linked Revit models."""
    linked_rooms = []
    link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    for link_instance in link_instances:
        link_doc = link_instance.GetLinkDocument()
        if link_doc:
            linked_rooms.extend(
                FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
            )
    return linked_rooms

def find_linked_room_for_tag(tag_point, linked_room):

    closed_shell = linked_room.ClosedShell
    for geom in closed_shell:
        for face in geom.Faces:
            normal = face.FaceNormal
            if normal.Z == -1:
                if face.Project(tag_point):
                    return linked_room

    return None

                    



def get_room_boundary_as_polygon(room):

    boundary_options = SpatialElementBoundaryOptions()
    bonudary_segment_length = room.GetBoundarySegments(boundary_options)

    if not bonudary_segment_length:
        return None  # Room has no boundary

    boundary_points = []
    for segment_list in bonudary_segment_length:
        for segment in segment_list:
            curve = segment.GetCurve()
            # Add the start points of the curve to define a closed boundary
            boundary_points.append(curve.GetEndPoint(0))

    # Ensure the boundary is closed by adding the first point at the end
    if boundary_points[0] != boundary_points[-1]:
        boundary_points.append(boundary_points[0])

    return boundary_points


def is_point_inside_boundary(point, boundary_points):

    # Project points to 2D plane (XY) for boundary checks
    px, py = point.X, point.Y
    for pt in boundary_points:
        print(pt)
    vertices = [(pt.X, pt.Y) for pt in boundary_points]

    inside = False
    n = len(vertices)
    print(n)
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        if ((y1 > py) != (y2 > py)) and (px < (x2 - x1) * (py - y1) / (y2 - y1) + x1):
            inside = inside
        else:
            inside = not inside
    return inside


# Collect all linked instances in the model
linked_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
linked_instance_dict = {link.Name: link for link in linked_instances}

if linked_instances:
    documentation_file = forms.alert("Is this a Documentation File or a Live File", warn_icon=False, options=["Documentation File", "Live File"])

    if not documentation_file:
        forms.alert("No file option selected. Exiting script.", exitscript=True)

    if documentation_file == "Documentation File":
        # Display a selection dialog for linked models
        selected_link_name = forms.SelectFromList.show(sorted(linked_instance_dict.keys()),title="Select Linked Model Containing Rooms",multiselect=False,button_name="Select")

        if not selected_link_name:
            # Exit if no link is selected
            forms.alert("No linked model selected. Exiting script.", title="Selection Cancelled", exitscript=True)
        else:
            # Get the selected linked instance
            selected_link_instance = linked_instance_dict[selected_link_name]

            linked_instance = selected_link_instance
            link_doc = selected_link_instance.GetLinkDocument()

    
    else:
        linked_instance = None
        link_doc = None
else:
    linked_instance = None
    link_doc = None




tg = TransactionGroup(doc, "Place Room Tags")
tg.Start()

ignored_rooms = []
for view in selected_views:
    all_room_tags = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_RoomTags).WhereElementIsNotElementType().ToElements()

    owned_elements = []
    unowned_elements = []
    processed_element_ids = set()  # Use a set to track processed elements and avoid duplicates

    elements_to_checkout = List[ElementId]()

    for tag in all_room_tags:
        # Only add to the checkout list if the element hasn't been processed yet
        if tag.Id not in processed_element_ids:
            elements_to_checkout.Add(tag.Id)
            processed_element_ids.add(tag.Id)  # Mark the element as processed

    if doc.IsWorkshared:
        
        # Attempt to checkout elements
        WorksharingUtils.CheckoutElements(doc, elements_to_checkout)
        # Check ownership of each element
        for elementid in elements_to_checkout:
            worksharingStatus = WorksharingUtils.GetCheckoutStatus(doc, elementid)
            if worksharingStatus != CheckoutStatus.OwnedByOtherUser:
                owned_elements.append(doc.GetElement(elementid))  # Add to owned elements
            else:
                if elementid not in [elem.Id for elem in unowned_elements]:  # Prevent duplicate additions
                    unowned_elements.append(doc.GetElement(elementid))  # Add to unowned element

    else:
            owned_elements = all_room_tags

    # Collect all room tags in the active view
    for tag in owned_elements:
        t = Transaction(doc, "delete tags")
        t.Start()

        if tag.HasLeader:
            continue
        if not tag.Room:
            doc.Delete(tag.Id)

        t.Commit()




for view in selected_views:
    if view.ViewType == ViewType.Section:
        section_bbox = view.CropBox
        section_transform = view.CropBox.Transform


        # Collect all rooms visible in the active view from the current document
        current_rooms = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

        # Transform the section bounding box to the linked model's coordinate system
        transform_to_link = linked_instance.GetTotalTransform().Inverse
        transformed_bbox = BoundingBoxXYZ()
        transformed_bbox.Transform = section_transform.Multiply(transform_to_link)
        transformed_bbox.Min = transform_to_link.OfPoint(section_bbox.Min)
        transformed_bbox.Max = transform_to_link.OfPoint(section_bbox.Max)


        # Collect all rooms visible in the active view from linked models
        linked_rooms = []
        if linked_instance:
            linkedrooms = FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

            for room in linkedrooms:
                room_bbox = room.get_BoundingBox(None)
                if room_bbox:
                    room_outline = Outline(room_bbox.Min, room_bbox.Max)
                    section_outline = Outline(transformed_bbox.Min, transformed_bbox.Max)
                    if room_outline.Intersects(section_outline, 0.001):
                        linked_rooms.append(room)

        
        all_rooms = list(current_rooms) + linked_rooms
        print(all_rooms)


        # Create a set of already tagged room IDs

        updated_room_tags = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_RoomTags).WhereElementIsNotElementType().ToElements()
        tagged_room_ids = set(tag.Room.Id for tag in updated_room_tags if hasattr(tag, "Room") and tag.Room)

        t = Transaction(doc, "place tags")
        t.Start()

        for room in current_rooms:

            # Skip if the room is already tagged
            if room.Id in tagged_room_ids:

                continue

            room_bbox = room.get_BoundingBox(view)
            bbox_center = (room_bbox.Min + room_bbox.Max) / 2
            bbox_center_uv = UV(bbox_center.X, bbox_center.Y)
            height = bbox_center.Z
            room_tag = doc.Create.NewRoomTag(LinkElementId(room.Id),bbox_center_uv,view.Id)
            room_tag.Location.Move(XYZ(0,0,height))

        for room in linked_rooms:
            print(room)
        for room in linked_rooms:
            area = room.LookupParameter("Area").AsDouble()
            print(area)
            if room.LookupParameter("Area").AsDouble() < 0.1:
                continue
            location = room.Location

            print(location)

            room_bbox = room.get_BoundingBox(view)
            bbox_center = (room_bbox.Min + room_bbox.Max) / 2
            height = bbox_center.Z
            room_center = location.Point
            room_center_uv = UV(room_center.X, room_center.Y)
            if linked_instance:

                has_tag_inside = False
                for tag in updated_room_tags:
                    if tag.HasLeader:  # For tags with a leader
                        leader_end_point = tag.LeaderEnd


                        # Get room boundary as a polygon
                        boundary_points = get_room_boundary_as_polygon(room)
                        if not boundary_points:

                            continue  # Skip rooms without boundaries

                        # Check if the leader endpoint is inside the room boundary
                        if is_point_inside_boundary(leader_end_point, boundary_points):
                            has_tag_inside = True
                if has_tag_inside:
                        continue
                else:
                    room_tag = doc.Create.NewRoomTag(LinkElementId(linked_instance.Id, room.Id),room_center_uv,view.Id)
                    room_tag.Location.Move(XYZ(0,0,height))
        t.Commit()

    elif view.ViewType == ViewType.FloorPlan:

        # Collect all rooms visible in the active view from the current document
        current_rooms = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

        # Collect all rooms visible in the active view from linked models
        linked_rooms = []
        if linked_instance:
            linkedrooms = FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

            for room in linkedrooms:
                if link_doc.GetElement(room.LevelId).Name != view.GenLevel.Name:
                    continue
                linked_rooms.append(room)


        all_rooms = list(current_rooms) + linked_rooms


        # Create a set of already tagged room IDs

        updated_room_tags = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_RoomTags).WhereElementIsNotElementType().ToElements()
        tagged_room_ids = set(tag.Room.Id for tag in updated_room_tags if hasattr(tag, "Room") and tag.Room)

        t = Transaction(doc, "place tags")
        t.Start()
        for room in current_rooms:

            # Skip if the room is already tagged
            if room.Id in tagged_room_ids:

                continue

            location = room.Location
            room_center = location.Point
            room_center_uv = UV(room_center.X, room_center.Y)
            room_tag = doc.Create.NewRoomTag(LinkElementId(room.Id),room_center_uv,view.Id)

        for room in linked_rooms:

            location = room.Location
            if location:
                room_center = location.Point
                room_center_uv = UV(room_center.X, room_center.Y)
                if linked_instance:

                    has_tag_inside = False
                    for tag in updated_room_tags:
                        if tag.HasLeader:  # For tags with a leader
                            leader_end_point = tag.LeaderEnd


                            # Get room boundary as a polygon
                            boundary_points = get_room_boundary_as_polygon(room)
                            if not boundary_points:

                                continue  # Skip rooms without boundaries

                            # Check if the leader endpoint is inside the room boundary
                            if is_point_inside_boundary(leader_end_point, boundary_points):
                                has_tag_inside = True
                    if has_tag_inside:
                            continue
                    else:
                        room_tag = doc.Create.NewRoomTag(LinkElementId(linked_instance.Id, room.Id),room_center_uv,view.Id)


        t.Commit()


t = Transaction(doc, "Center Room Tags")
t.Start()

for view in selected_views:
    if view.ViewType == ViewType.Section:
        continue

    all_room_tags = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_RoomTags).WhereElementIsNotElementType().ToElements()

    for tag in all_room_tags:

        if tag.HasLeader:  # Skip tags with leaders
            continue

        room = tag.Room
        tag_location = tag.Location.Point

        if room is None:  # Attempt to find a linked room
            for linked_room in linked_rooms:
                if find_linked_room_for_tag(tag_location, linked_room):
                    room = linked_room
                    room_name = room.LookupParameter("Name").AsString()

                    closed_shell = room.ClosedShell
                    for geom in closed_shell:
                        for face in geom.Faces:
                            normal = face.FaceNormal
                            if normal.Z == -1:
                                edges = face.EdgeLoops
                                for edge in edges:
                                    bonudary_segment_length = edge.Size

                                    # Handling small boundary rooms
                                    if bonudary_segment_length <= 5:
                                        room_bb = room.get_BoundingBox(view)
                                        room_center = (room_bb.Min + room_bb.Max) / 2

                                        if room.IsPointInRoom(room_center):
                                            move_room_and_tag(tag, room, room_center)
                                        location = room.Location
                                        if not linked_instance:
                                            if location:
                                                location.Point = room_center
                                            else:

                                                ignored_rooms.append(room)
                                        continue

                                    # Handling larger polygonal rooms
                                    try:
                                        closed_shell = room.ClosedShell
                                        for geom in closed_shell:
                                            for face in geom.Faces:
                                                normal = face.FaceNormal
                                                if normal.Z == -1:
                                                    bottom_face = face

                                        if bottom_face:
                                            mesh = bottom_face.Triangulate()
                                            largest_triangle_center = None
                                            largest_area = 0

                                            for i in range(mesh.NumTriangles):
                                                triangle = mesh.get_Triangle(i)
                                                pt1, pt2, pt3 = triangle.get_Vertex(0), triangle.get_Vertex(1), triangle.get_Vertex(2)
                                                area = calculate_triangle_area(pt1, pt2, pt3)

                                                if area > largest_area:
                                                    largest_area = area
                                                    largest_triangle_center = get_triangle_center(pt1, pt2, pt3)

                                            if largest_triangle_center and room.IsPointInRoom(largest_triangle_center):
                                                move_room_and_tag(tag, room, largest_triangle_center)

                                            else:

                                                ignored_rooms.append(room)

                                    except Exception as e:
                                        print("Error processing room: {0}, {1}".format(room.Id, e))
                                        ignored_rooms.append(room)
                    break
        else:

                    boundary_segments = room.GetBoundarySegments(SpatialElementBoundaryOptions())

                    if not boundary_segments:
                        print("no boundary segements for room", room.Id)
                        continue

                    if len(boundary_segments[0]) <= 5:
                        room_bb = room.get_BoundingBox(view)
                        room_center = (room_bb.Min + room_bb.Max) / 2

                        if room.IsPointInRoom(room_center):
                            move_room_and_tag(tag, room, room_center)
                            room.Location.Point = room_center
                            centered_tags_count += 1  # Increment the counter
                        else:
                            ignored_rooms.append(room)

                    else:
                        curve_loop = CurveLoop()
                        outer_loop = boundary_segments[0]

                        for segment in outer_loop:
                            curve = segment.GetCurve()
                            curve_loop.Append(curve)

                        height = 0.1
                        extrusion_direction = XYZ(0, 0, 1)
                        solid = GeometryCreationUtilities.CreateExtrusionGeometry([curve_loop], extrusion_direction, height)

                        bottom_face = None
                        for face in solid.Faces:
                            normal = face.ComputeNormal(UV(0.5, 0.5))
                            if normal.IsAlmostEqualTo(XYZ(0, 0, -1)):
                                bottom_face = face
                                break

                        if bottom_face:
                            mesh = bottom_face.Triangulate()

                            triangle_areas = []
                            largest_area = 0
                            largest_triangle_center = None

                            for i in range(mesh.NumTriangles):
                                try:
                                    triangle = mesh.get_Triangle(i)

                                    pt1 = triangle.get_Vertex(0)
                                    pt2 = triangle.get_Vertex(1)
                                    pt3 = triangle.get_Vertex(2)

                                    area = calculate_triangle_area(pt1, pt2, pt3)
                                    triangle_areas.append(area)

                                    if area > largest_area:
                                        largest_area = area
                                        largest_triangle_center = get_triangle_center(pt1, pt2, pt3)
                                except Exception as e:
                                    print("error processing mesh for room ")
                                    

                            if is_similar_area(triangle_areas):
                                room_bb = room.get_BoundingBox(view)
                                room_center = (room_bb.Min + room_bb.Max) / 2
                                if room.IsPointInRoom(room_center):

                                    move_room_and_tag(tag, room, room_center)
                                    room.Location.Point = room_center
                                    centered_tags_count += 1  # Increment the counter
                                else:
                                    ignored_rooms.append(room)
                            else:
                                if largest_triangle_center and room.IsPointInRoom(largest_triangle_center):

                                    move_room_and_tag(tag, room, largest_triangle_center)
                                    room.Location.Point = largest_triangle_center
                                    centered_tags_count += 1  # Increment the counter
                                else:
                                    ignored_rooms.append(room)

                       
t.Commit()

tg.Assimilate()

if unowned_elements or ignored_rooms:
    output.print_md(header_data)

if unowned_elements:
    unowned_element_data = []
    for element in unowned_elements:
        try:
            unowned_element_data.append([output.linkify(element.Id), element.Category.Name.upper(), "REQUEST OWNERSHIP", WorksharingUtils.GetWorksharingTooltipInfo(doc, element.Id).Owner])
        except Exception as e:
            pass
    
    
    output.print_md("##⚠️ Elements Skipped ☹️")  # Markdown Heading 2
    output.print_md("---")  # Markdown Line Break
    output.print_md("❌ Make sure you have Ownership of the Elements - Request access. Refer to the **Table Report** below for reference")  # Print a Line
    output.print_table(table_data=unowned_element_data, columns=["ELEMENT ID", "CATEGORY", "TO-DO", "CURRENT OWNER"])  # Print a Table
    print("\n\n")
    output.print_md("---")  # Markdown Line Break


# Display ignored rooms
if ignored_rooms:
    ignored_room_info = []
    for room in ignored_rooms:
        room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
        room_number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
        room_id = room.Id
        ignored_room_info.append([room_name, room_number, output.linkify(room_id)])

if ignored_rooms:   
    output.print_md("## ⚠️ Rooms Skipped ☹️")
    output.print_md("---")
    output.print_md("❌ Rooms skipped : Point outside room boundary. Refer to the **Table Report** below for reference.")

    output.print_table(table_data=ignored_room_info, columns=["ROOM NAME", "ROOM NUMBER", "ELEMENT ID"])

    print("\n\n")
    output.print_md("---")

