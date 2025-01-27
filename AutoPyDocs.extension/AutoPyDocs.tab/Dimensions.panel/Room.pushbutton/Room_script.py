# -*- coding: utf-8 -*-

__title__ = "Test walls"
__author__ = "roma ramnani"

import clr
import time
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB      import *
from pyrevit                import revit, forms, script
from datetime               import datetime

from doc_functions          import get_view_on_sheets
from dim_functions          import dim_room_wall, segment_reference, filter_linear_walls
from Extract.RunData        import get_run_data

doc = revit.doc
view = doc.ActiveView
output = script.get_output()
app = __revit__.Application 
rvt_year = app.SubVersionNumber
model_name = doc.Title
tool_name = __title__ 
user_name = app.Username

start_time = time.time()
manual_time = 57

op = app.Create.NewGeometryOptions()
op.ComputeReferences = True
op.IncludeNonVisibleObjects = True

def isParallel(v1, v2):
    cross_product = v1.CrossProduct(v2)
    return cross_product.IsAlmostEqualTo(XYZ(0, 0, 0))

def isCollinear(l0, l1, tolerance=1e-9):
    a = l0.GetEndPoint(0)
    b = l0.GetEndPoint(1)
    c = l1.GetEndPoint(0)
    d = l1.GetEndPoint(1)
    
    # Define vectors
    v0 = b - a
    v1 = d - c
    v2 = c - a
    
    # Check if vectors are parallel
    if not isParallel(v0, v1):
        return False

    # Check if the lines lie on the same infinite line
    return v0.CrossProduct(v2).IsAlmostEqualTo(XYZ(0, 0, 0), tolerance)

def is_almost_equal(value1, value2, tolerance=1e-9):
    return abs(value1 - value2) < tolerance

def normal_line(seg, offset_distance):
    scale_factor = view.Scale
    segment_line = seg.GetCurve()

    if isinstance(segment_line, Arc):
        room_issue_counts[room_key]["CurvedBoundaryCount"] += 1
        return None
    elif isinstance(segment_line, Line):
        seg1_pt1 = segment_line.GetEndPoint(0)
        seg1_pt2 = segment_line.GetEndPoint(1)

        if is_almost_equal(seg1_pt1.X, seg1_pt2.X, tolerance=1e-9):
            if seg1_pt1.Y > seg1_pt2.Y:
                s_direction  = seg1_pt1 - seg1_pt2
                first_pt = seg1_pt1 
            else:
                s_direction  = seg1_pt2 - seg1_pt1
                first_pt = seg1_pt2 
        elif seg1_pt1.X > seg1_pt2.X:
            s_direction  = seg1_pt1 - seg1_pt2
            first_pt = seg1_pt1 
        else:
            s_direction  = seg1_pt2 - seg1_pt1
            first_pt = seg1_pt2 
    
        # Create a line from the midpoint in the direction of the normal vector
        direction = s_direction.Normalize()
        normal_vector = XYZ(-direction.Y, direction.X, 0) * 0.25
        o_direction = (direction * offset_distance) * scale_factor / 100

        pt1 = first_pt - o_direction 
        pt2 = normal_vector + pt1
        if room.IsPointInRoom(pt2):
            nl = Line.CreateBound(pt1, pt2)
        else:
            nl = Line.CreateBound(pt1, pt1 - normal_vector)
        return nl

    else:
        room_issue_counts[room_key]["InvalidReference"] += 1
        return None

def cs_normal_line(room, seg, offset_distance, cs=None):    
    scale_factor = view.Scale
    if cs and len(cs) > 1:
        #print("cs found")
        seg1_pt1 = cs[0].GetCurve().GetEndPoint(0)
        seg1_pt2 = cs[0].GetCurve().GetEndPoint(1)
        seg2_pt1 = cs[-1].GetCurve().GetEndPoint(0)
        seg2_pt2 = cs[-1].GetCurve().GetEndPoint(1)

        points_list = [seg1_pt1, seg1_pt2, seg2_pt1, seg2_pt2]
        sorted_points = sorted(points_list, key=lambda point: (point[0], point[1]))
        # if seg1_pt1.X == seg1_pt2.X:
        #     sorted_points = sorted(points_list, key=lambda point: (point[1], point[0]), reverse=True)
        # elif seg1_pt1.Y == seg1_pt2.Y:
        #     sorted_points = sorted(points_list, key=lambda point: (point[0], point[1]))
        # else:
        #     sorted_points = sorted(points_list, key=lambda point: (point[0], point[1]))

        start_pt_index = points_list.index(sorted_points[0])
        end_pt_index = points_list.index(sorted_points[-1])
        s_line_start_pt = points_list[start_pt_index]
        s_line_end_pt = points_list[end_pt_index]

        #print(s_line_start_pt)
        #print(s_line_end_pt)
        #print(s_line_start_pt>s_line_end_pt)
        if s_line_start_pt > s_line_end_pt:
            s_direction = s_line_start_pt - s_line_end_pt 
            first_pt = s_line_start_pt
        else:
            s_direction = s_line_end_pt - s_line_start_pt
            first_pt = s_line_end_pt
        s_length = s_direction.GetLength()
        if s_length == 0:
            #print("Warning: Collinear set has zero length. Check the start and end points.")
            return None

    else:
        segment_line = seg.GetCurve()
        if isinstance(segment_line, Arc):
            #room_issue_counts[room_key]["CurvedBoundaryCount"] += 1
            return None
        elif isinstance(segment_line, Line):
            #s_direction = segment_line.Direction
            s_length = segment_line.Length
            s_line_start_pt = segment_line.GetEndPoint(0)   
            s_line_end_pt = segment_line.GetEndPoint(1)
            if s_line_start_pt > s_line_end_pt:
                s_direction = s_line_start_pt - s_line_end_pt 
                first_pt = s_line_start_pt
            else:
                s_direction = s_line_end_pt - s_line_start_pt
                first_pt = s_line_end_pt
            
        else:
            #room_issue_counts[room_key]["InvalidReference"] += 1
            return None

    direction = s_direction.Normalize()
    normal_vector = XYZ(-direction.Y, direction.X, 0) * 0.25
    o_direction = (direction * offset_distance) * scale_factor / 100
    
    pt1 = first_pt - o_direction 
    pt2 = normal_vector + pt1

    if room.IsPointInRoom(pt2):
    # Create a line from the midpoint in the direction of the normal vector
        nl = Line.CreateBound(pt1, pt2)
    else:
        nl = Line.CreateBound(pt1, pt1 - normal_vector)
    return nl

def model_rooms(doc):
    rooms = []
    if doc_opted == 'Model File':
        model_rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
        rooms.extend(model_rooms)
    if doc_opted =='Documentation File':
        linked_rooms = FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms)
        rooms.extend(linked_rooms)

    rooms_in_view = list(rooms)
    #print(rooms_in_view)
    filtered_rooms = [room for room in rooms_in_view if room.Area>0]

    return filtered_rooms

def view_rooms(view):
    rooms = []
    if doc_opted == 'Model File':
        model_rooms = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Rooms)
        rooms.extend(model_rooms)
    if doc_opted =='Documentation File':
        filtered = []
        linked_rooms = FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms)
        for room in linked_rooms:
            if link_doc.GetElement(room.LevelId).Name != view.GenLevel.Name:
                filtered.append(room)
        rooms.extend(filtered)

    rooms_in_view = list(rooms)
    #print(rooms_in_view)
    filtered_rooms = [room for room in rooms_in_view if room.Area>0]

    return filtered_rooms

def selected_model_rooms(doc):
    rooms_to_dim = []
    room_lists = []
    room_dict = {}
    filtered_rooms = model_rooms(doc)
    #print(filtered_rooms)
    unique_rooms = {room.Id: room for room in filtered_rooms}.values()
    
    for room in unique_rooms:
        room_name = room.LookupParameter("Name").AsString()
        room_number = room.LookupParameter("Number").AsString()
        room_entry = "{}: {}".format(room_name, room_number)
        room_lists.append(room_entry)
        room_dict[room_entry] = room
    
    rooms_selected = forms.SelectFromList.show(room_lists,
                                               title='Select rooms to dimension',
                                               multiselect=True,
                                               button_name='Select')
    
    for room_entry in rooms_selected:
        room = room_dict[room_entry]
        rooms_to_dim.append(room)
    #print ([(room.LookupParameter("Number").AsString()) for room in rooms_to_dim])
    return rooms_to_dim

def swap_collinear_segments(collinear_sets):
    for cs in collinear_sets:
        if len(cs) > 1:
            cs[0], cs[-1] = cs[-1], cs[0]

def get_hosted_elements(wall):
    # Find all inserts (doors, windows, etc.) in the wall
    insert_ids = wall.FindInserts(True, True, True, True)
    
    hosted_elements = []
    for insert_id in insert_ids:
        element = doc.GetElement(insert_id)
        hosted_elements.append(element)
    
    return hosted_elements

def get_wall_reference(wall):
    opt = Options()  # Create an options object for retrieving the geometry
    geom_elem = wall.get_Geometry(opt)  # Get the geometry of the wall
    
    for geom_obj in geom_elem:
        if hasattr(geom_obj, 'Faces'):
            for face in geom_obj.Faces:
                # Return the reference to the first face found
                return face.Reference
    return None

doc_ops = ['Model File', 'Documentation File']
doc_opted = forms.alert("This Revit model is:", 
                        warn_icon=False, 
                        options = doc_ops)

if not doc_opted:
    forms.alert ("File type not selected", exitscript = True)

def select_link_doc(doc):
    
    linked_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    if not linked_instances:
        forms.alert("No linked document", exitscript=True)

    link_names = [link.Name for link in linked_instances]

    target_instance_names = forms.SelectFromList.show(link_names, title="Select File with Rooms", width=600, height=600, button_name="Select File", multiselect=False, exitscript = True)

    if not target_instance_names:
        script.exit()
    
    for link_instance in linked_instances:
        if link_instance.Name in target_instance_names:    
            link_doc = link_instance.GetLinkDocument()  
            transform = link_instance.GetTotalTransform() 
            selected_link = link_instance

    return selected_link, link_doc

if doc_opted =='Documentation File':
    selected_link, link_doc = select_link_doc(doc)
else:
    selected_link = None

dim_ops = ['Length & Width only', 'Least width', 'All wall segments']
dim_opted = forms.SelectFromList.show(dim_ops,
                                      title='Plan to dimension is for:',
                                      multiselect=True,
                                      button_name='Select',
                                    )

if not dim_opted:
    forms.alert ("Dimension criteria not selected", exitscript = True)

view_types = {
        'Floor Plans': ViewType.FloorPlan,
        'Reflected Ceiling Plans': ViewType.CeilingPlan,
        'Area Plans': ViewType.AreaPlan, 
    }
selected_views = get_view_on_sheets(doc, view_types)

# Create SpatialElementBoundaryOptions and set boundary location
options = SpatialElementBoundaryOptions()
options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish

offset_distance = 1.2
room_issue_counts = {}
failed_data = []
#room = filtered_rooms[9] 
selected_rooms = selected_model_rooms(doc)
filtered_rooms = model_rooms(doc)
sel_room_ids = []
for room in selected_rooms:
    room_id = room.Id
    sel_room_ids.append(room_id)

try:
    t = Transaction(doc, "Dimension Room")
    t.Start()

    total_dim_count = 0
    total_room_count = 0

    for view in selected_views:
        #print(view.Name)
        rooms_in_view = view_rooms(view)
        #print(rooms_in_view)
        for room in rooms_in_view:
            #print("room is in view")
            room_id = room.Id
            if room_id in sel_room_ids:
                #print("Yes room found")
                room_id = room.Id
                room_key = output.linkify(room_id)
                if room_key not in room_issue_counts:
                    level_param = room.LookupParameter("Level")
                    name_param = room.LookupParameter("Name")
                    number_param = room.LookupParameter("Number")

                    room_issue_counts[room_key] = {
                        "Level": level_param.AsValueString().upper() if level_param else "N/A",
                        "Name": name_param.AsString().upper() if name_param else "N/A",
                        "Number": number_param.AsString().upper() if number_param else "N/A",
                        "NonParallelCount": 0,
                        "CurvedBoundaryCount": 0,
                        "InvalidReference": 0, 
                    }

                segments = []
                #print("Processing room: {}-{}".format(room.LookupParameter("Name").AsString(), room.LookupParameter("Number").AsString()))

                # Get boundary segments for the room
                loops = room.GetBoundarySegments(options)
                # Collect all segments
                for loop in loops:
                    for segment in loop:
                        segments.append(segment)
                #print(loops)
                #print("Segments:{}".format(segments))

                directions = []
                segment_groups = []
                # Group segments based on direction

                for segment in segments:
                    curve = segment.GetCurve()

                    d = curve.GetEndPoint(1) - curve.GetEndPoint(0)  # Direction vector of the segment

                    idx = -1
                    for i, direction in enumerate(directions):
                        if isParallel(d, direction):
                            idx = i
                            break

                    if idx != -1:
                        # Append to existing group
                        segment_groups[idx].append(segment)
                    else:
                        # Create new group
                        directions.append(d)
                        new_group = [segment]
                        segment_groups.append(new_group)

                diagonal_walls = []
                for i, group in enumerate(segment_groups):
                    csets = [] 
                    segment_list = []
                    length_dir = {}
                    if len(group) >= 2: 
                        #print("Group{} : {}".format(i, len(group)))
                        for segment in group:
                            for cs in csets:
                                if len(cs) > 0:
                                    if cs:
                                        l0 = segment.GetCurve()
                                        l1 = cs[0].GetCurve()
                                        if isCollinear(l0, l1):
                                            cs.append(segment)
                                            break
                            else:
                                # No collinear group found, so create a new one
                                csets.append([segment])

                        for cs in csets:
                            #print (cs)
                            #print("Set length: {}".format(len(cs)))
                            if len(cs) > 1:
                                #print("cs")
                                total_length = sum(seg.GetCurve().Length for seg in cs)
                                length_dir[cs[0]] = total_length
                                # for seg in cs:
                                #     doc.Create.NewDetailCurve(view, seg.GetCurve())
                            elif len(cs) == 1:
                                #print("seg")
                                for segment in cs:
                                    length_dir[segment] = segment.GetCurve().Length

                        sorted_length_dir = dict(sorted(length_dir.items(), key=lambda item: item[1]))
                        #print(i, sorted_length_dir.items())

                        #Pair segments with the distance between them
                        set_pairs = []
                        
                        for i, [s0, length0] in enumerate(sorted_length_dir.items()):
                            c0 = s0.GetCurve()
                            for [s1, length1] in list(sorted_length_dir.items())[i+1:]:
                                c1 = s1.GetCurve()
                                d = c0.Distance(c1.GetEndPoint(0))
                                set_pairs.append([d, s0, s1])
                        #print(set_pairs)
                        
                        sorted_by_distance = sorted(set_pairs, key=lambda x: x[0], reverse=True)

                        #print(sorted_by_distance)
                        if sorted_by_distance:
                            #Add dimension to segments farthest apart
                            if 'Length & Width only' in dim_opted:
                                first = sorted_by_distance[0][1]
                                second = sorted_by_distance[0][2]
                                #print(type(first))
                                refArray = ReferenceArray() 
                                refArray.Append(segment_reference(doc_opted, first, selected_link))
                                refArray.Append(segment_reference(doc_opted, second, selected_link)) 
                                
                                length_first = sorted_length_dir[first]
                                length_second = sorted_length_dir[second]

                                first_cs = None
                                second_cs = None
                                for cs in csets:
                                    if len(cs) > 1:
                                        if first in cs:
                                            first_cs = cs
                                        if second in cs:
                                            second_cs = cs

                                if length_first >= length_second:
                                    #print("Length: {}".format(length_second))
                                    if second_cs and len(second_cs) > 1:
                                        #print("second_cs")
                                        nl = cs_normal_line(room, second, offset_distance, second_cs)
                                    else:
                                        nl = normal_line(second, offset_distance)
                                else:
                                    #print("Length: {}".format(length_first))
                                    if first_cs and len(first_cs) > 1:
                                        #print("first_cs")
                                        nl = cs_normal_line(room, first, offset_distance, first_cs)
                                    else:
                                        nl = normal_line(first, offset_distance)
                                
                                #print(nl)
                                # if nl is not None:
                                #     doc.Create.NewDetailCurve(view, nl)
                                
                                ref1 = segment_reference(doc_opted, first, selected_link)
                                ref2 = segment_reference(doc_opted, second, selected_link)

                                # print("RefSize: {}".format(refArray.Size))
                                # print("Ref1: {}".format(ref1))
                                # print("Ref2: {}".format(ref2))

                                if not ref1 or not ref2:
                                    room_issue_counts[room_key]["InvalidReference"] += 1
                                
                                if ref1 and ref2:
                                    #try:
                                        if refArray.Size>=2:
                                            d1 = doc.Create.NewDimension(view, nl, refArray)
                                            # if d1:
                                            #     print("Dimensioned")
                                            total_dim_count += 1
                                        else:
                                            room_issue_counts[room_key]["InvalidReference"] += 1
                                    #except Exception as e:
                                        #print("Error in LxW dimension: {}".format(str(e)))
                                #print(d1)
                                #Add dimension to segments closest to each other
                                if 'Least width' in dim_opted:
                                    if len(set_pairs) >= 2:
                                        first_sg = sorted_by_distance[-1][1]
                                        second_sg = sorted_by_distance[-1][2]
                                        #print(type(first))
                                        curve1 = first_sg.GetCurve()
                                        curve2 = second_sg.GetCurve()
                                        if curve1.Length >= curve2.Length: 
                                            nl2 = normal_line(second_sg, offset_distance)
                                        else:
                                            nl2 = normal_line(first_sg, offset_distance)
                                        if nl2 is not None:
                                            #doc.Create.NewDetailCurve(view, nl2)
                                            refArray2 = ReferenceArray()
                                            refArray2.Append(segment_reference(doc_opted, first_sg, selected_link))  
                                            refArray2.Append(segment_reference(doc_opted, second_sg, selected_link))
                                            if segment_reference(doc_opted, first_sg) is None or segment_reference(doc_opted, second_sg, selected_link) is None:
                                                room_issue_counts[room_key]["InvalidReference"]+= 1
                                            else:
                                                # print(refArray2.Size)
                                                # print(refArray2)
                                                if refArray2.Size>=2:
                                                    d2 = doc.Create.NewDimension(view, nl2, refArray2)
                                                    total_dim_count += 1
                                                else:
                                                    room_issue_counts[room_key]["InvalidReference"] += 1
                                
                                #Add dimension to all segments
                                if 'All wall segments' in dim_opted:
                                    if len(set_pairs) >=3:
                                        second = sorted_by_distance[0][2]
                                        nl3 = normal_line(second, (offset_distance*2))
                                        if nl3 is not None:
                                            #nl3 = normal_line(sorted_dir[0], (offset_distance*2))
                                            #doc.Create.NewDetailCurve(view, nl3)
                                            refArray3 = ReferenceArray()
                                            for sorted_segment in sorted_length_dir:
                                                refArray3.Append(segment_reference(doc_opted, sorted_segment, selected_link))
                                            #print(refArray3.Size) 
                                            if segment_reference(doc_opted, sorted_segment) is None:
                                                room_issue_counts[room_key]["InvalidReference"] += 1
                                            else:
                                                if refArray3.Size>=2: 
                                                    d3 = doc.Create.NewDimension(view, nl3, refArray3)
                                                    total_dim_count += (refArray3.Size - 1)
                                                else:
                                                    room_issue_counts[room_key]["InvalidReference"] += 1
                        
                    elif len(group) == 1:
                        seg = group[0]
                        wall = doc.GetElement(seg.ElementId)
                        if wall:
                            wall_id = wall.Id
                            #seg_curve = seg.Curve
                            wall_curve = wall.Location.Curve

                            if isinstance (wall_curve, Line) :
                                if wall_id not in [doc.GetElement(d_wall.ElementId).Id for d_wall in diagonal_walls]:
                                    diagonal_walls.append(seg)
                                    #print("Segment added: {}".format(wall.Id))
                             
                            else:   
                                room_issue_counts[room_key]["CurvedBoundaryCount"] += 1
                                #forms.alert("Room needs to have atleast 2 parallel sides to make linear dimension")
                if diagonal_walls:
                    collected_walls = FilteredElementCollector(doc, view.Id).OfClass(Wall).WhereElementIsNotElementType().ToElements()  
                    #linear_walls = filter_linear_walls(collected_walls)
                    for diagonal_seg in diagonal_walls:
                        #print("line dim")
                        dim_room_wall(view, room, diagonal_seg, collected_walls, offset_distance)

                total_room_count += 1
            else:
                room_number = room.LookupParameter("Number").AsString()
                #print("Room {} not in selectd rooms: {}".format(room_number, view.Name))
                continue 

    t.Commit()

    end_time = time.time()
    runtime = end_time - start_time
            
    run_result = "Tool ran successfully"
    element_count = total_dim_count
    error_occured ="Nil"
    #print("Runtime in seconds: {}".format(runtime))
    get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)

except Exception as e:
    print("Error occurred: {}".format(str(e)))
    t.RollBack()

    end_time = time.time()
    runtime = end_time - start_time
    print("Runtime in seconds: {}".format(runtime))
    error_occured = ("Error occurred: %seg", str(e))    
    run_result = "Error"
    element_count = 0
    
    get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)

for room_key, data in room_issue_counts.iteritems(): 
    issue_description = []
    if data["CurvedBoundaryCount"] > 0:
        issue_description.append("CURVED BOUNDARY")
    if data["NonParallelCount"] > 0:
        issue_description.append("NON-PARALLEL LINES")
    if data["InvalidReference"] > 0:
        issue_description.append("INVALID REFERENCE")
    if issue_description:
        failed_room_data = [
            room_key,
            data["Level"],
            data["Name"],
            data["Number"],
            ", ".join(issue_description)
        ]
        failed_data.append(failed_room_data)

if failed_data:
    header_data = datetime.now().strftime("%Y-%m-%d %H:%M:%S"), rvt_year, tool_name, model_name, user_name 
    print(header_data)
    processed_data = total_room_count - len(failed_data)
    output.print_md("##⚠️ Completed. {} rooms dimensioned. {} Issues Found ".format(processed_data, len(failed_data)))
    output.print_table(table_data=failed_data, columns=["ELEMENT ID", "LEVEL", "ROOM NAME", "ROOM NUMBER", "EXCEPTION"]) 
    print("\n\n")
    output.print_md("---") 
    output.print_md("***❌ EXCEPTION REFERENCE***")
    output.print_md("---")
    output.print_md("**CURVED BOUNDARY** - Room has curved room boundary\n")
    output.print_md("**NON-PARALLEL LINES** - Room has bounding elements/lines that are not parallel to any other room boundary \n")
    output.print_md("**INVALID REFERNCE** - Unable to get a valid start and end reference for dimension\n")
else:
    forms.alert("Completed. {} rooms dimensioned.".format(total_room_count))
