# -*- coding: utf-8 -*-
'''dim_functions'''
__Title__ = " dim_functions"
__author__ = "romaramnani"

import clr
import os

clr.AddReference("System.Windows.Forms")
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Forms           import FolderBrowserDialog
from System.Windows                 import Window
from System.Windows.Controls        import Button
from Autodesk.Revit.DB              import *
from Autodesk.Revit.UI              import *
from pyrevit                        import revit, forms, script
from pyrevit.framework              import List

from g_curve_functions              import translate_curve, isParallel, isPerpendicular, isAlmostEqualTo

import Autodesk.Revit.DB            as DB

script_dir  = os.path.dirname(__file__)
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application
doc         = __revit__.ActiveUIDocument.Document  # type: Document
uiapp       = __revit__

opts = Options()
#without compute references, none of this works
opts.ComputeReferences = True
opts.IncludeNonVisibleObjects = True
opts.View = doc.ActiveView

op = app.Create.NewGeometryOptions()
op.ComputeReferences = True
op.IncludeNonVisibleObjects = True

class MultipleSlidersForm:
    def __init__(self):
        # Load the XAML window
        xaml_path = os.path.join(script_dir, 'Slider.xaml')
        self.window = forms.WPFWindow(xaml_path)

        # Attach event handlers
        self.window.equalValuesCheckBox.Checked += self.OnCheckBoxChecked
        self.window.equalValuesCheckBox.Unchecked += self.OnCheckBoxUnchecked

        try:
            self.window.Sl_left.ValueChanged += self.OnSliderValueChanged
            self.window.Sl_right.ValueChanged += self.OnSliderValueChanged
            self.window.Sl_top.ValueChanged += self.OnSliderValueChanged
            self.window.Sl_bottom.ValueChanged += self.OnSliderValueChanged
        except Exception as ex:
            print("Error while binding ValueChanged event:", ex)

        submit_button = self.window.FindName('Submit')
        submit_button.Click += self.submit_button_click

        self.window.ShowDialog()

    def OnCheckBoxChecked(self, sender, e):
        # Set all sliders to the value of Sl_left and disable other sliders
        value = self.window.Sl_left.Value
        self.window.Sl_right.Value = value
        self.window.Sl_top.Value = value
        self.window.Sl_bottom.Value = value

        self.window.Sl_right.IsEnabled = False
        self.window.Sl_left.IsEnabled = True
        self.window.Sl_top.IsEnabled = False
        self.window.Sl_bottom.IsEnabled = False

    def OnCheckBoxUnchecked(self, sender, e):
        # Enable all sliders
        self.window.Sl_right.IsEnabled = True
        self.window.Sl_top.IsEnabled = True
        self.window.Sl_bottom.IsEnabled = True

    def OnSliderValueChanged(self, sender, e):
        # If the checkbox is checked, synchronize all sliders
        if self.window.equalValuesCheckBox.IsChecked:
            value = sender.Value
            self.window.Sl_left.Value = value
            self.window.Sl_right.Value = value
            self.window.Sl_top.Value = value
            self.window.Sl_bottom.Value = value

    def submit_button_click(self, sender, e):
        # Close the window
        self.window.Close()

class DoubleSlidersForm:
    def __init__(self):
        # Load the XAML window
        xaml_path = os.path.join(script_dir, 'Double_slider.xaml')
        self.window = forms.WPFWindow(xaml_path)

        # Attach event handlers
        self.window.equalValuesCheckBox.Checked += self.OnCheckBoxChecked
        self.window.equalValuesCheckBox.Unchecked += self.OnCheckBoxUnchecked

        try:
            self.window.Sl_left.ValueChanged += self.OnSliderValueChanged
            self.window.Sl_right.ValueChanged += self.OnSliderValueChanged

        except Exception as ex:
            print("Error while binding ValueChanged event:", ex)

        submit_button = self.window.FindName('Submit')
        submit_button.Click += self.submit_button_click

        self.window.ShowDialog()

    def OnCheckBoxChecked(self, sender, e):
        # Set all sliders to the value of Sl_left and disable other sliders
        value = self.window.Sl_left.Value
        self.window.Sl_right.Value = value

        self.window.Sl_right.IsEnabled = False

    def OnCheckBoxUnchecked(self, sender, e):
        # Enable all sliders
        self.window.Sl_right.IsEnabled = True
        self.window.Sl_left.IsEnabled = True

    def OnSliderValueChanged(self, sender, e):
        # If the checkbox is checked, synchronize all sliders
        if self.window.equalValuesCheckBox.IsChecked:
            value = sender.Value
            self.window.Sl_left.Value = value
            self.window.Sl_right.Value = value

    def submit_button_click(self, sender, e):
        # Close the window
        self.window.Close()

class SliderData:
    def __init__(self, name, min_value, max_value, default_value, tick_frequency=1):
        self.Name = name
        self.Min = min_value
        self.Max = max_value
        self.Value = default_value
        self.TickFrequency = tick_frequency

class DynamicSliderForm:
    def __init__(self, slider_definitions):
        #Load the XAML window
        xaml_path = os.path.join(script_dir, 'Slider.xaml')
        self.window = forms.WPFWindow(xaml_path)

        self.sliders_container = self.window.FindName("SlidersContainer")
        if self.sliders_container is None:
            raise ValueError("SlidersContainer not found in XAML")

        self.sliders = []
        self.sliders_container.ItemsSource = self.sliders

        self.load_sliders(slider_definitions)

        #Attach event handlers
        self.window.equalValuesCheckBox.Checked += self.OnCheckBoxChecked
        self.window.equalValuesCheckBox.Unchecked += self.OnCheckBoxUnchecked

        submit_button = self.window.FindName('Submit')
        submit_button.Click += self.submit_button_click

        self.window.ShowDialog()

    def add_slider(self, name, min_value, max_value, default_value):
        self.sliders.Add(SliderData(name, min_value, max_value, default_value))

    def load_sliders(self, slider_definitions):
        for name, settings in slider_definitions.items():
            self.add_slider(name, settings["min"], settings["max"], settings["default"])

        slider_definitions = {
            "Left Offset": {"min": 0, "max": 100, "default": 15},
            "Right Offset": {"min": 0, "max": 100, "default": 20},
            "Top Offset": {"min": 0, "max": 100, "default": 10},
            "Bottom Offset": {"min": 0, "max": 100, "default": 25},
        }

    def on_checkbox_checked(self, sender, event):
        #Set all sliders to the same value
        if self.sliders.Count > 0:
            value = self.sliders[0].Value
            for slider in self.sliders:
                slider.Value = value

    def on_checkbox_unchecked(self, sender, event):
        pass

    def submit_button_click(self, sender, e):
        slider_values = {slider.Name: slider.Value for slider in self.sliders}
        #Close the window
        self.window.Close()

def isclose(a, b, rel_tol=1e-9, abs_tol=0.0):
    return abs(a - b) <= max(rel_tol * max(abs(a), abs(b)), abs_tol)

def is_vector_equal(v1, v2, tolerance=1e-6):
    return (v1 - v2).GetLength() < tolerance

def convert_to_xyz(float_list):
    # Assuming the float_list contains coordinates in the order [X, Y, Z]
    if len(float_list) == 3:
        return DB.XYZ(float_list[0], float_list[1], float_list[2])
    else:
        forms.alert("List must contain exactly 3 float values.")

def datum_points(datum_elements, view):
    start_points = []
    end_points = []
    datum_curves = []
    for datum_element in datum_elements:
        grid_curve = datum_element.GetCurvesInView(DB.DatumExtentType.ViewSpecific, view)  # Get the 2D extent curve of the grid line

        if grid_curve:
            for curve in grid_curve:
                start_points.append(curve.GetEndPoint(0))
                end_points.append(curve.GetEndPoint(1))
                datum_curves.append(curve)
        else:
            print("No curves found for the grid in the specified view.")

    # Ensure all points are XYZ objects
    start_points = [convert_to_xyz([point.X, point.Y, point.Z]) if not isinstance(point, DB.XYZ) else point for point in start_points]
    end_points = [convert_to_xyz([point.X, point.Y, point.Z]) if not isinstance(point, DB.XYZ) else point for point in end_points]

    return start_points, end_points, datum_curves

def process_wall_faces(link_doc, wall, seg_curve, link_instance):
    ref = None
    try:
        rExt = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior)[0]
        rInt = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior)[0]
        #print(rExt)
        rExt_linked = rExt.CreateLinkReference(link_instance)
        rInt_linked = rInt.CreateLinkReference(link_instance)
        #print(rExt_linked)
        elemExt = link_doc.GetElement(rExt)
        elemInt = link_doc.GetElement(rInt)
        #print(elemExt)
        #print(elemInt)
        fExt = elemExt.GetGeometryObjectFromReference(rExt)
        fInt = elemInt.GetGeometryObjectFromReference(rInt)

        if isinstance(fExt, Face) and fExt.Intersect(seg_curve) in [SetComparisonResult.Overlap, SetComparisonResult.Subset]:
            #print("fExt")
            ref = rExt_linked
        elif isinstance(fInt, Face) and fInt.Intersect(seg_curve) in [SetComparisonResult.Overlap, SetComparisonResult.Subset]:
            #print("fInt")
            ref = rInt_linked
        return ref
    except Exception as e:
        print("Error processing wall faces: {}".format(e))
        return ref

def process_curtain_wall(wall, link_instance):
    ref = None
    try:
        wall_elements = wall.get_Geometry(op)
        for wall_element in wall_elements:
            if isinstance(wall_element, Solid):
                for face in wall_element.Faces:
                    ref = face.Reference.CreateLinkReference(link_instance)
        return ref
    except Exception as e:
        print("Error processing curtain wall: {}".format(e))
        return ref
          
def segment_reference(doc_opted, seg, link_instance = None):
    ref = None
    if doc_opted == 'Model File':
        se = doc.GetElement(seg.ElementId)
        if se:
            #print("Processing element {} of {}".format(se.Id, type(se)))
            seg_curve = seg.GetCurve()

        # If the element is a model line (room separator)
        if isinstance(se, ModelLine):
            return se.GeometryCurve.Reference

        # If the element is a wall
        elif isinstance(se, Wall):
            # Get the exterior and interior faces
            rExt = HostObjectUtils.GetSideFaces(se, ShellLayerType.Exterior)[0]
            rInt = HostObjectUtils.GetSideFaces(se, ShellLayerType.Interior)[0]

            elemExt = doc.GetElement(rExt)
            elemInt = doc.GetElement(rInt)

            fExt = elemExt.GetGeometryObjectFromReference(rExt)
            fInt = elemInt.GetGeometryObjectFromReference(rInt)

            # Check if the curve intersects with the faces
            if isinstance(fExt, Face) and fExt.Intersect(seg_curve) in [SetComparisonResult.Overlap, SetComparisonResult.Subset]:
                return rExt
            if isinstance(fInt, Face) and fInt.Intersect(seg_curve) in [SetComparisonResult.Overlap, SetComparisonResult.Subset]:
                return rInt

            # If the element is a curtain wall
            if se.WallType.FamilyName == 'Curtain Wall':
                # Get the interior finish face
                wall_elements= se.get_Geometry(op)
                for wall_element in wall_elements:
                    if isinstance(wall_element,Solid):
                        for f in wall_element.Faces:
                            ref = f.Reference
                return ref

        elif isinstance(se, RevitLinkInstance):
            link_doc = se.GetLinkDocument()
            seg_elem = link_doc.GetElement(seg.LinkElementId)
            #print(seg_elem.Id, type(seg_elem))
            # If the element is a model line (room separator)
            if isinstance(seg_elem, ModelLine):
                return seg_elem.GeometryCurve.Reference.CreateLinkReference(se)

            # If the element is a wall
            elif isinstance(seg_elem, Wall):
                ref = process_wall_faces(link_doc, seg_elem, seg_curve, se)
                # If the element is a curtain wall
                if seg_elem.WallType.FamilyName == 'Curtain Wall':
                    ref = process_curtain_wall(seg_elem, se)
                return ref
            else:
                print("Invalid reference for element: {}".format(se.Id))  

    elif doc_opted =='Documentation File':
        #print("Got selected link")
        link_doc = link_instance.GetLinkDocument()
        se = link_doc.GetElement(seg.ElementId)
        if se:
            #print("Processing element {} of {}".format(se.Id, type(se)))
            seg_curve = seg.GetCurve()

        #get_link_element_refernce(seg, link_doc, se, link_instance)
        # If the element is a model line (room separator)
        if isinstance(se, ModelLine):
            return se.GeometryCurve.Reference.CreateLinkReference(link_instance)

        # If the element is a wall
        elif isinstance(se, Wall):
            ref = process_wall_faces(link_doc, se, seg_curve, link_instance)
            # If the element is a curtain wall
            if se.WallType.FamilyName == 'Curtain Wall':
                ref = process_curtain_wall(se, link_instance)
            return ref
        
        elif isinstance(se, RevitLinkInstance):
            link2_doc = se.GetLinkDocument()
            seg_elem = link2_doc.GetElement(seg.LinkElementId)
            # If the element is a model line (room separator)
            if isinstance(seg_elem, ModelLine):
                return seg_elem.GeometryCurve.Reference.CreateLinkReference(se)

            # If the element is a wall
            elif isinstance(seg_elem, Wall):
                ref = process_wall_faces(link2_doc, seg_elem, seg_curve, se)
                # If the element is a curtain wall
                if seg_elem.WallType.FamilyName == 'Curtain Wall':
                    ref = process_curtain_wall(seg_elem, se)
                return ref
            
            else:
                print("Invalid reference for element: {}".format(se.Id)) 
        else:       
            #print("Invalid reference for element: {}".format(se.Id))  
            return None
    return ref

def tolist(obj1):
	if hasattr(obj1,"__iter__"): return obj1
	else: return [obj1]

# Set the wall orientation, depending on whether exterior or interior is selected
def wallNormal(wall, face_is_int):
    if face_is_int:
        wall_normal = wall.Orientation.Negate()
    else:
        wall_normal = wall.Orientation
    return wall_normal

# Move wall location to align wall face at cutplane level
def translate_loc_crv(view, wall, wallNormal, lineEndExtend):
    wallLn = wall.Location.Curve

    # Take the level height, subtract the height of the wall base
    zOf = view.GenLevel.Elevation - wallLn.GetEndPoint(0).Z
    #print(view.GenLevel.Elevation)
    # Get the cut plane offset of the view
    cPlaneH = view.GetViewRange().GetOffset(PlanViewPlane.CutPlane)
    #cPlaneHiMM = UnitUtils.ConvertFromInternalUnits(cPlaneH, UnitTypeId.Millimeters)

    # Translate the wall curve to the cut plane height
    translation_vector = XYZ(0, 0, zOf + cPlaneH)
    transform = Transform.CreateTranslation(translation_vector)
    wallLn = wallLn.CreateTransformed(transform)

    # Translate the wall curve to the external edge
    #wallwidthMM = UnitUtils.ConvertFromInternalUnits(wall.Width, UnitTypeId.Millimeters)
    translation_vector = wallNormal.Multiply(wall.Width / 2)
    transform = Transform.CreateTranslation(translation_vector)
    wallLn = wallLn.CreateTransformed(transform)
    if lineEndExtend>0:
        #mid_pt = (wallLn.GetEndPoint(0) + wallLn.GetEndPoint(1))/2

        wall_dir1 = wallLn.GetEndPoint(0) - wallLn.GetEndPoint(1)
        wall_dir2 = wallLn.GetEndPoint(1) - wallLn.GetEndPoint(0)

        ptMvSt = wallLn.GetEndPoint(0).Add(wall_dir1.Multiply(lineEndExtend))
        ptMvEnd = wallLn.GetEndPoint(1).Add(wall_dir2.Multiply(lineEndExtend))

        lineAtExternalEdgeAtCutPlaneHeight = Line.CreateBound(ptMvSt, ptMvEnd)
#        print('Line is extended')
        return lineAtExternalEdgeAtCutPlaneHeight

    else:
#        print('original line')
        return wallLn

def filter_linear_walls(walls):
    linear_walls = []
    for wall in walls:
        location = wall.Location
        if location and hasattr(location, "Curve"):  # Ensure the wall has a location curve
            curve = location.Curve
            if isinstance(curve, Line):  # Check if the curve is a line (linear)
                linear_walls.append(wall)
    return linear_walls

def intersecting_wall(view, wall, collected_walls, face_normal):
    intersected_walls = []
    intersected_walls.append(wall)
    for view_wall in collected_walls:
        if translate_loc_crv(view, wall, face_normal, 0).Intersect(translate_loc_crv(view, view_wall, face_normal, 0)) == SetComparisonResult.Overlap:
            intersected_walls.append(view_wall)
            #print(type(view_wall))
        elif translate_loc_crv(view, wall, face_normal, 0).Intersect(translate_loc_crv(view, view_wall, face_normal, 0)) == SetComparisonResult.Overlap:
            if view_wall not in intersected_walls:
                intersected_walls.append(view_wall)
                #print(type(view_wall))
        elif wall.Location.Curve.Intersect(view_wall.Location.Curve) == SetComparisonResult.Overlap:
            if view_wall not in intersected_walls:
                intersected_walls.append(view_wall)
                #print(type(view_wall))
    #print(len(intersected_walls))
    return intersected_walls

def wall_face_to_room(wall, room):

    options = DB.Options()
    geometry = wall.get_Geometry(options)

    room_center = room.Location.Point
    filtered_face = None

    for geom_obj in geometry:
        if isinstance(geom_obj, DB.Solid):
            for face in geom_obj.Faces:
                normal = face.FaceNormal

                face_point = face.Evaluate(DB.UV(0.5, 0.5))
                vector_to_room = room_center - face_point

                # Check if the face normal is pointing towards the room
                if normal.DotProduct(vector_to_room) > 0:
                    filtered_face = face
                    break

    return filtered_face

def dim_room_wall(view, room, seg, collected_walls, offset_distance):
    wall = doc.GetElement(seg.ElementId)
    #print("Def processing: {}".format(wall.Id))
    ref_line_extension = 0
    room_center = room.Location.Point

    filtered_face = wall_face_to_room(wall, room)
    face_normal = filtered_face.FaceNormal
    main_face_midpoint = filtered_face.Evaluate(DB.UV(0.5, 0.5))

    intersected_walls = intersecting_wall(view, wall, collected_walls, face_normal)
    count_intersecting_walls = len(intersected_walls)
    #print("Count of walls:{}".format(count_intersecting_walls))
    #print(wall.Id for wall in intersected_walls)

    line_on_face = translate_loc_crv(view, wall, face_normal, ref_line_extension)
    ref_line = translate_curve(line_on_face, face_normal, offset_distance)

    vert_edges = []
    intersecting_faces = []
    #print(len(intersected_walls))
    for wall_int in intersected_walls:
        for obj in wall_int.get_Geometry(opts):
            if isinstance(obj, Solid):
                for face in obj.Faces:
                    face_normal = face.ComputeNormal(UV(0.5, 0.5))
                    if isAlmostEqualTo(wall.Orientation, face_normal) or isAlmostEqualTo(wall.Orientation.Negate(), face_normal):
                        pass
                    else:
                        intersecting_faces.append(face)

                for edge in obj.Edges:
                    edge_curve = edge.AsCurve()
                    if isinstance (edge_curve, Line):
                        edge_direction = edge_curve.Direction.Normalize()
                        intersection_result = edge_curve.Intersect(line_on_face)
                        # Check if the edge intersects the line and is vertical
                        if edge_direction.IsAlmostEqualTo(XYZ(0, 0, 1)) or edge_direction.IsAlmostEqualTo(XYZ(0, 0, -1)):
                            #if isclose(abs(edge_direction.Z), 1, abs_tol= 1e-3):
                            if intersection_result != SetComparisonResult.Disjoint:
                                vert_edges.append(edge)
                                
    #print("Vert Edges no: {}".format(vert_edges))
    edge_locations = []# Get location as sorting parameters to remove duplicates

    wall_check = []
    for i, edge in enumerate(vert_edges):
        point = edge.AsCurve().GetEndPoint(0)
        point_loc = (point.X + point.Y, i)  # Include the index to ensure uniqueness
        edge_locations.append((round(point_loc[0], 7), point_loc[1]))
        if room.IsPointInRoom(point):
            wall_check.append(True)
        
    # Sort edges based on the unique edge_locations
    vert_edges_sorted = [x for _, x in sorted(zip(edge_locations, vert_edges))]
    vert_edges_loc_sorted = sorted(edge_locations, key=lambda x: x[0])
    #print(vert_edges_sorted )
    #print(vert_edges_loc_sorted)

    #     #For removing corner conditions of interior walls, finding the corner edges which have to be removed.
    #     if len(vert_edges_sorted)>2 and vert_edges_loc_sorted[0] == vert_edges_loc_sorted[1]:
    # #                    print('hi')
    #         vert_edges_loc_sorted.remove(vert_edges_loc_sorted[0])
    #         vert_edges_sorted.remove(vert_edges_sorted[0])
    #         vert_edges_loc_sorted.remove(vert_edges_loc_sorted[0])#Since first 0 removes the first value and second item will become the first item after the previous step
    #         vert_edges_sorted.remove(vert_edges_sorted[0])
                        
    #     edgeslen=len(vert_edges_loc_sorted)
    #     if len(vert_edges_sorted)>2 and vert_edges_loc_sorted[edgeslen-1] == vert_edges_loc_sorted[edgeslen-2]:
    # #                    print('bye')
    #         vert_edges_loc_sorted.remove(vert_edges_loc_sorted[edgeslen-1])
    #         vert_edges_sorted.remove(vert_edges_sorted[edgeslen-1])
    #         vert_edges_loc_sorted.remove(vert_edges_loc_sorted[edgeslen-2])#Since first [edgeslen-1] removes the last value and second last item will become the last item after the previous step 
    #         vert_edges_sorted.remove(vert_edges_sorted[edgeslen-2])              

    #Filetring out only the unique edges 
    vert_edge_uni=[]
    vert_edge_uni_loc_sorted=[]
    for i in vert_edges_loc_sorted:
        if i not in vert_edge_uni_loc_sorted:
            indexd=vert_edges_loc_sorted.index(i)
            vert_edge_uni_loc_sorted.append(i)
            vert_edge_uni.append(vert_edges_sorted[indexd])


    if vert_edge_uni: 
        dim = ReferenceArray()
        #print("vert_edge length:{}".format(len(vert_edge_uni)))
        if count_intersecting_walls > 3: 
            first_pt = vert_edge_uni[0].AsCurve().GetEndPoint(0)
            if room.IsPointInRoom(first_pt):
                dim.Append(vert_edge_uni[0].Reference)
                dim.Append(vert_edge_uni[1].Reference)
            else:
                dim.Append(vert_edge_uni[len(vert_edge_uni) - 2].Reference)
                dim.Append(vert_edge_uni[len(vert_edge_uni) - 1].Reference)
        else:
            dim.Append(vert_edge_uni[0].Reference)
            dim.Append(vert_edge_uni[len(vert_edge_uni) - 1].Reference)
        try:
            doc.Create.NewDimension(view, ref_line, dim)
        except Exception as e:
            print("Error creating dimension: {}".format(e))

