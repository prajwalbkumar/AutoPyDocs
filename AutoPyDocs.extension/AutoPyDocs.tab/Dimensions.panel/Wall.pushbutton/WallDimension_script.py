# -*- coding: utf-8 -*-
# '''Test'''
__title__ = " Walls"
__author__ = "roma ramnani, astle james"

import clr
import time

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB      import *
from pyrevit                import revit, forms, script
import Autodesk.Revit.DB    as DB
from doc_functions          import get_view_on_sheets
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

opts = Options()
#without compute references, none of this works
opts.ComputeReferences = True
opts.IncludeNonVisibleObjects = True
opts.View = doc.ActiveView


def isParallel(v1, v2):
    #it needs two vectors
    return v1.CrossProduct(v2).IsAlmostEqualTo(XYZ(0, 0, 0))

def isAlmostEqualTo(vec1, vec2, tolerance=1e-9):
    return vec1.IsAlmostEqualTo(vec2, tolerance)

def isPerpendicular(v1, v2):
    if v1.DotProduct(v2)== 0:
        return True
    else:
        return False

def tolist(obj1):
	if hasattr(obj1,"__iter__"): return obj1
	else: return [obj1]

def CurveToVector(crv):
    start_point = crv.GetEndPoint(0)
    end_point = crv.GetEndPoint(1)
    vec = end_point - start_point
    return vec.Normalize()

# Set the wall orientation, depending on whether exterior or interior is selected
def wallNormal(wall, extOrInt):
    if extOrInt:
        wall_normal = wall.Orientation
    else:
        wall_normal = wall.Orientation.Negate()
    return wall_normal


# Move wall location to align wall face at cutplane level
def locToCutCrv(wall, wallNormal, lineEndExtend):
    wallLn = wall.Location.Curve
    
    # Take the level height, subtract the height of the wall base
    zOf = view.GenLevel.Elevation - wallLn.GetEndPoint(0).Z
    #print(view.GenLevel.Elevation)
    # Get the cut plane offset of the view
    cPlaneH = view.GetViewRange().GetOffset(PlanViewPlane.CutPlane)
    #cPlaneHiMM = UnitUtils.ConvertFromInternalUnits(cPlaneH, UnitTypeId.Millimeters)
    
    # Translate the wall curve to the cut plane height
    translation_vector = XYZ(0, 0, zOf)
    transform = Transform.CreateTranslation(translation_vector)
    wallLn = wallLn.CreateTransformed(transform)
    
    # Translate the wall curve to the external edge
    #wallwidthMM = UnitUtils.ConvertFromInternalUnits(wall.Width, UnitTypeId.Millimeters)
    translation_vector = wallNormal.Multiply(wall.Width / 2)
    transform = Transform.CreateTranslation(translation_vector)
    wallLn = wallLn.CreateTransformed(transform)
    if lineEndExtend>0:
        mid_pt = (wallLn.GetEndPoint(0) + wallLn.GetEndPoint(1))/2

        wall_dir1 = mid_pt - wallLn.GetEndPoint(0)
        wall_dir2 = mid_pt - wallLn.GetEndPoint(1)

        ptMvSt = wallLn.GetEndPoint(0).Add(wall_dir1.Multiply(lineEndExtend))
        ptMvEnd = wallLn.GetEndPoint(1).Add(wall_dir2.Multiply(lineEndExtend))

        lineAtExternalEdgeAtCutPlaneHeight = Line.CreateBound(ptMvSt, ptMvEnd)
        return lineAtExternalEdgeAtCutPlaneHeight
  
    return wallLn

def translate_curve(curve, direction, distance):
    translation_vector = direction.Multiply(distance)
    transform = Transform.CreateTranslation(translation_vector)
    translated_curve = curve.CreateTransformed(transform)
    return translated_curve

def create_model_line(curve):
    work_plane = view.SketchPlane
    plane_origin = work_plane.Origin
    plane_normal = work_plane.Normal
    def project_point(point, plane_origin, plane_normal):
        # Vector from the plane origin to the point
        origin_to_point = point - plane_origin
        # Project the vector onto the plane's normal to find the offset
        distance_to_plane = origin_to_point.DotProduct(plane_normal)
        # Calculate the projected point
        projected_point = point - plane_normal.Multiply(distance_to_plane)
        return projected_point

    p1 = curve.GetEndPoint(0)
    p2 = p1 = curve.GetEndPoint(0)

    # Get projected start and end points
    projected_start = project_point(p1, plane_origin, plane_normal)
    projected_end = project_point(p2, plane_origin, plane_normal)

    # Create the new projected curve
    if isinstance(curve, Line):
        projected_curve = Line.CreateBound(projected_start, projected_end)
    elif isinstance(curve, Arc):
        # Additional handling required for arcs
        mid_point = curve.Evaluate(0.5, True)
        projected_mid = project_point(mid_point, plane_origin, plane_normal)
        projected_curve = Arc.Create(projected_start, projected_end, projected_mid)
    else:
        raise NotImplementedError("Curve type not supported for projection.")

    model_line = doc.Create.NewModelCurve(projected_curve, view.SketchPlane)  
    if model_line:
        print("Line created")

def filter_linear_walls(walls):
    linear_walls = []
    for wall in walls:
        location = wall.Location
        if location and hasattr(location, "Curve"):  # Ensure the wall has a location curve
            curve = location.Curve
            if isinstance(curve, Line):  # Check if the curve is a line (linear)
                linear_walls.append(wall)
    return linear_walls




t = Transaction(doc, "Dimension internal Wall")
#start transaction
t.Start()

view_types = {'Floor Plans': ViewType.FloorPlan,'Reflected Ceiling Plans': ViewType.CeilingPlan,'Area Plans': ViewType.AreaPlan}
selected_views = get_view_on_sheets(doc, view_types)
wall_types_list = []
wall_types = FilteredElementCollector(doc).OfClass(WallType).WhereElementIsElementType().ToElements()
for wall_type in wall_types:
    #family_name = wall_type.FamilyName
    type_name = wall_type.LookupParameter("Type Name").AsString()
    #wall_name = ": ".join([family_name, type_name])
    wall_types_list.append(type_name)
selected_wall_types = forms.SelectFromList.show(wall_types_list,title='Select Wall Types to dimension',multiselect=True,button_name='Select')
if not selected_wall_types:
    forms.alert ("Wall type not selected", exitscript = True)

extOrInt = False
offDist = 1.2

# Extend dimesion for exterior side dimensions
if extOrInt:
    intersectLineEndExtend = 500
else:
    intersectLineEndExtend = 0

for view in selected_views:
    collected_walls = FilteredElementCollector(doc, view.Id).OfClass(Wall).WhereElementIsNotElementType().ToElements()
    linear_walls_in_view = filter_linear_walls(collected_walls)
    for targetWall in linear_walls_in_view:
        if targetWall.WallType.LookupParameter("Type Name").AsString() in selected_wall_types:
            intersectedWalls = []
            intersectedWalls.append(targetWall)
            
            for collectedWall in linear_walls_in_view:
                if targetWall.Location.Curve.Intersect(collectedWall.Location.Curve) == SetComparisonResult.Overlap:
                    intersectedWalls.append(collectedWall)
                    
            line_on_face = locToCutCrv(targetWall, wallNormal(targetWall, extOrInt), intersectLineEndExtend)
            
            ref_line = translate_curve(line_on_face, wallNormal(targetWall, extOrInt), offDist)
            
            frontFaceIW = []
            vertEdges = []# Get vertical edges at intersection
            for wallInt in intersectedWalls:
                for obj in wallInt.get_Geometry(opts):
                    if isinstance(obj, Solid):
                        for face in obj.Faces:
                            face_normal = face.ComputeNormal(UV(0.5, 0.5))
                            if isAlmostEqualTo(wallInt.Orientation, face_normal):
                                frontFaceIW.append(face)
                                
                        for edge in obj.Edges:
                            edgeC = edge.AsCurve()
                            edgeCNorm = edgeC.Direction.Normalize()
                            intersection_result = edgeC.Intersect(line_on_face)
                            if edgeCNorm.IsAlmostEqualTo(XYZ(0, 0, 1)) or edgeCNorm.IsAlmostEqualTo(XYZ(0, 0, -1)):
                                if intersection_result != SetComparisonResult.Disjoint:
                                    vertEdges.append(edge)
                                    
            for obj in targetWall.get_Geometry(opts):
                if isinstance(obj, Solid): 
                    faceTW = obj.Faces
                                    
            strayEdges = []
            for v in vertEdges:
                strayEdges.append(v)
                
            for faIW in frontFaceIW:
                i = 0
                length = len(strayEdges)
                while (i < length):
                    if strayEdges[i].GetFace(0) in faceTW:
                        strayEdges.Remove(strayEdges[i])
                        length = length - 1
                    elif strayEdges[i].GetFace(1) in faceTW:
                        strayEdges.Remove(strayEdges[i])
                        length = length - 1            
                    #if our wall is external.... #if edge face is an intersecting wall's front face, we don't want it                
                    elif strayEdges[i].GetFace(0) == faIW and extOrInt == True:
                        strayEdges.Remove(strayEdges[i])            
                        length = length - 1 
                    elif strayEdges[i].GetFace(1) == faIW and extOrInt == True:
                        strayEdges.Remove(strayEdges[i])
                        length = length - 1 
                    # strayEdges.Remove(ed)
                    #or if the edge reference is to a non-wall
                        continue
                    i = i+1           
                                                    
            vertEdgesLoc = []# Get location as sorting parameters to remove duplicates
            for v in vertEdges:
                vLoc = v.AsCurve().GetEndPoint(0).X + v.AsCurve().GetEndPoint(0).Y
                vertEdgesLoc.append(round(vLoc,7))
            
            #if the wall is exterior, we want to remove references to internal wall edge
            if extOrInt == True:
                i=0
                length = len(vertEdgesLoc)
                strayCLoc2 = []
                while (i < length):
                    for stray in strayEdges:
                        stLoc = stray.AsCurve().GetEndPoint(0).X + stray.AsCurve().GetEndPoint(0).Y
                        #getting eroneous values, Revit accuracy not good enough? round is built in method
                        if round(vertEdgesLoc[i],7) == round(stLoc,7):
                            vertEdges.Remove(vertEdges[i])
                            vertEdgesLoc.Remove(vertEdgesLoc[i])
                            length = length - 1    
                            continue
                    i = i+1

            vertEdgesSorted = [x for _,x in sorted(zip(vertEdgesLoc,vertEdges))]
            vertEdgesLocSorted = sorted(vertEdgesLoc)
            
            vertEdgesSortedUnique=[]#Filetring out only the unique edges
            vertEdgesLocSortedUnique=[]
            
            for i in vertEdgesLocSorted:
                indexd=vertEdgesLocSorted.index(i)
                if i not in vertEdgesLocSortedUnique:
                    vertEdgesLocSortedUnique.append(i)
                    vertEdgesSortedUnique.append(vertEdgesSorted[indexd])
                    
            print(vertEdgesLocSorted)
            print(vertEdgesLocSortedUnique)
            
            try:
                index1=0
                while index1<=len(vertEdgesSortedUnique)-1:
                    print(index1, len(vertEdgesSortedUnique))
                    vertEdgeSub1 = ReferenceArray()
                    if index1 % 2==0:
                        vertEdgeSub1.Append(vertEdgesSortedUnique[index1].Reference)
                        vertEdgeSub1.Append(vertEdgesSortedUnique[index1+1].Reference)
                        dim = doc.Create.NewDimension(view, ref_line, vertEdgeSub1)
                        print (dim,'dd')
                        index1=index1+2
                    elif index1 % 2!=0:
                        index1=index1+1
                        print('odd no.')
                        
            except Exception as e:
                print(e)
                
        else:
            print("No wall found for selected wall types in view {}".format(view.Name))



t.Commit()

end_time = time.time()
runtime = end_time - start_time
print("Runtime : {} seconds".format(runtime))