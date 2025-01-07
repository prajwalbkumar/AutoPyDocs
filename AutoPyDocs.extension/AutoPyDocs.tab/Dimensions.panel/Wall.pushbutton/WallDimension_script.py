# -*- coding: utf-8 -*-
# '''Test'''
__title__ = " Walls"
__author__ = "roma ramnani, Astle James"

import clr
import time

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB      import *
from pyrevit                import revit, forms, script
import Autodesk.Revit.DB    as DB
from doc_functions          import get_view_on_sheets
from Extract.RunData        import get_run_data
output = script.get_output()
doc = revit.doc
view = doc.ActiveView
output = script.get_output()
app = __revit__.Application 
rvt_year = app.SubVersionNumber
model_name = doc.Title
tool_name = __title__ 
user_name = app.Username

failed_data_Collected=[]

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
    translation_vector = XYZ(0, 0, zOf + cPlaneH)
    transform = Transform.CreateTranslation(translation_vector)
    wallLn = wallLn.CreateTransformed(transform)
    
    # Translate the wall curve to the external edge
    #wallwidthMM = UnitUtils.ConvertFromInternalUnits(wall.Width, UnitTypeId.Millimeters)
    translation_vector = wallNormal.Multiply(wall.Width / 2)
    transform = Transform.CreateTranslation(translation_vector)
    wallLn = wallLn.CreateTransformed(transform)
    if lineEndExtend>0:
        mid_pt = (wallLn.GetEndPoint(0) + wallLn.GetEndPoint(1))/2

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

view_types = {'Floor Plans': ViewType.FloorPlan,'Reflected Ceiling Plans': ViewType.CeilingPlan,'Area Plans': ViewType.AreaPlan}
selected_views = get_view_on_sheets(doc, view_types)
wall_types_list = []
wall_types = FilteredElementCollector(doc).OfClass(WallType).WhereElementIsElementType().ToElements()
for wall_type in wall_types:
    type_name = wall_type.LookupParameter("Type Name").AsString()
    wall_types_list.append(type_name)
selected_wall_types = forms.SelectFromList.show(wall_types_list,title='Select Wall Types to dimension',multiselect=True,button_name='Select')
if not selected_wall_types:
    forms.alert ("Wall type not selected", exitscript = True)
collected_walls = FilteredElementCollector(doc, view.Id).OfClass(Wall).WhereElementIsNotElementType().ToElements()    
linear_walls_in_view = filter_linear_walls(collected_walls)

def DimensionWallsExterior(offDist, extOrInt):
    # Extend dimesion for exterior side dimensions
    if extOrInt:
        #UnitUtils.ConvertToInternalUnits(500, UnitTypeId.Millimeters)
        intersectLineEndExtend = 0.5# 1 foot = 304.8 millimeters/ 500mm/offset is defined here(what ever is given in python code is divided by308 in mm in revit)
    else:
        intersectLineEndExtend = 0
#    print(intersectLineEndExtend,'intersectLineEndExtend')
    
    for view in selected_views:
        for targetWall in linear_walls_in_view:
            #For exterior walls which has to be dimensioned in the external face:
            if targetWall.WallType.LookupParameter("Type Name").AsString() in selected_wall_types and 'Ext' in targetWall.WallType.LookupParameter("Type Name").AsString():
#                print ('yes-Ext/Ext')
                intersectedWalls = []
                intersectedWalls.append(targetWall)
                for collectedWall in linear_walls_in_view:
                    #For Ext walls line is offseted to the internal side to collect even the wall stoping at internal face.
                    if locToCutCrv(targetWall, wallNormal(targetWall, False), 0).Intersect(locToCutCrv(collectedWall, wallNormal(collectedWall, False), 0)) == SetComparisonResult.Overlap:
                        intersectedWalls.append(collectedWall)
                    if locToCutCrv(targetWall, wallNormal(targetWall, False), 0).Intersect(locToCutCrv(collectedWall, wallNormal(collectedWall, True), 0)) == SetComparisonResult.Overlap:
                        if collectedWall not in intersectedWalls:
                            intersectedWalls.append(collectedWall)    
                    if targetWall.Location.Curve.Intersect(collectedWall.Location.Curve) == SetComparisonResult.Overlap:
                        if collectedWall not in intersectedWalls:
                            intersectedWalls.append(collectedWall)                                          
                        
                line_on_face = locToCutCrv(targetWall, wallNormal(targetWall, extOrInt), intersectLineEndExtend)# for dimensioning the external side of exterior wall
                line_on_face_interior = locToCutCrv(targetWall, wallNormal(targetWall, False), 0.2)# for dimensioning the internal side of exterior wall   
                
                ref_line = translate_curve(line_on_face, wallNormal(targetWall, extOrInt), offDist)# for dimensioning the external side of exterior wall
                ref_line_interior = translate_curve(line_on_face_interior, wallNormal(targetWall, False), offDist)# for dimensioning the internal side of exterior wall   

                frontFaceIW = []
                vertEdges = []# Get vertical edges at intersection for dimensioning the internal side of exterior wall  
                vertEdgesExt = []# Get vertical edges at intersection for dimensioning the External side of exterior wall  
                for wallInt in intersectedWalls:
                    for obj in wallInt.get_Geometry(opts):
                        if isinstance(obj, Solid):
                            for face in obj.Faces:
                                face_normal = face.ComputeNormal(UV(0.5, 0.5))
                                if isAlmostEqualTo(wallInt.Orientation, face_normal):
                                    frontFaceIW.append(face)
                                    
                            # for dimensioning the external side of exterior wall       
                            for edge in obj.Edges:
                                edgeC = edge.AsCurve()
                                edgeCNorm = edgeC.Direction.Normalize()
                                intersection_result = edgeC.Intersect(line_on_face)
                                if edgeCNorm.IsAlmostEqualTo(XYZ(0, 0, 1)) or edgeCNorm.IsAlmostEqualTo(XYZ(0, 0, -1)):
                                    if intersection_result != SetComparisonResult.Disjoint:
                                        vertEdgesExt.append(edge)
                                        
                            # for dimensioning the internal side of exterior wall  
                            for edge in obj.Edges:
                                edgeC = edge.AsCurve()
                                edgeCNorm = edgeC.Direction.Normalize()
                                intersection_result = edgeC.Intersect(line_on_face_interior)
                                if edgeCNorm.IsAlmostEqualTo(XYZ(0, 0, 1)) or edgeCNorm.IsAlmostEqualTo(XYZ(0, 0, -1)):
                                    if intersection_result != SetComparisonResult.Disjoint:
                                        vertEdges.append(edge)
                
                # for i in intersectedWalls:
                #     print('intersected walls collected', output.linkify(i.Id))
                    
                # for i in frontFaceIW:
                #     print('faces collected',i)
                    
                # for i in vertEdgesExt:
                #     print('edges collected ext',i)
                    
                # for i in vertEdges:
                #     print('edges collected int',i)
                    
                # for dimensioning the ext side of exterior wall                                        
                vertEdgesLocExt = []# Get location as sorting parameters to remove duplicates
                for v in vertEdgesExt:
                    vLocExt = v.AsCurve().GetEndPoint(0).X + v.AsCurve().GetEndPoint(0).Y
                    vertEdgesLocExt.append(round(vLocExt,7))
                
                # for dimensioning the ext side of exterior wall   
                vertEdgesSortedExt = [x for _,x in sorted(zip(vertEdgesLocExt,vertEdgesExt))]
                vertEdgesLocSortedExt = sorted(vertEdgesLocExt)           
                         
                # for dimensioning the external side of exterior wall
                try:
                    if len(vertEdgesLocSortedExt)>=2:
                        ExternalSideDim=ReferenceArray()
                        ExternalSideDim.Append(vertEdgesSortedExt[0].Reference)#First
                        ExternalSideDim.Append(vertEdgesSortedExt[len(vertEdgesSortedExt)-1].Reference)#Last
                        if ExternalSideDim.Size >= 2:
                            dim = doc.Create.NewDimension(view, ref_line, ExternalSideDim)
                except Exception as e:
                    failed_data_Collected.append(['Error', e])
                    failed_data_Collected.append(['Element Name', targetWall.Name])
                    failed_data_Collected.append(['Element ID', output.linkify(targetWall.Id)])                
                        
                                               
                # for dimensioning the internal side of exterior wall                                        
                vertEdgesLoc = []# Get location as sorting parameters to remove duplicates
                for v in vertEdges:
                    vLoc = v.AsCurve().GetEndPoint(0).X + v.AsCurve().GetEndPoint(0).Y
                    vertEdgesLoc.append(round(vLoc,7))
                
                # for dimensioning the internal side of exterior wall   
                vertEdgesSorted = [x for _,x in sorted(zip(vertEdgesLoc,vertEdges))]
                vertEdgesLocSorted = sorted(vertEdgesLoc)
                
                                        
                #creating exception to remove extra edges at the ends for 2nd category of walls:
                
                
                if len(vertEdgesSorted)>2 and vertEdgesLocSorted[0] == vertEdgesLocSorted[1]:
#                    print('hi')
                    vertEdgesLocSorted.remove(vertEdgesLocSorted[0])
                    vertEdgesSorted.remove(vertEdgesSorted[0])
                    vertEdgesLocSorted.remove(vertEdgesLocSorted[0])#Since first 0 removes the first value and second item will become the first item after the previous step
                    vertEdgesSorted.remove(vertEdgesSorted[0])
                    
                edgeslen=len(vertEdgesLocSorted)
                if len(vertEdgesSorted)>2 and vertEdgesLocSorted[edgeslen-1] == vertEdgesLocSorted[edgeslen-2]:
#                    print('bye')
                    vertEdgesLocSorted.remove(vertEdgesLocSorted[edgeslen-1])
                    vertEdgesSorted.remove(vertEdgesSorted[edgeslen-1])
                    vertEdgesLocSorted.remove(vertEdgesLocSorted[edgeslen-2])#Since first [edgeslen-1] removes the last value and second last item will become the last item after the previous step 
                    vertEdgesSorted.remove(vertEdgesSorted[edgeslen-2])              

                # for dimensioning the internal side of exterior wall    
                vertEdgesSortedUnique=[]#Filetring out only the unique edges
                vertEdgesLocSortedUnique=[]
                for i in vertEdgesLocSorted:
                    if i not in vertEdgesLocSortedUnique:
                        indexd=vertEdgesLocSorted.index(i)
                        vertEdgesLocSortedUnique.append(i)
                        vertEdgesSortedUnique.append(vertEdgesSorted[indexd])
                        
#                print(vertEdgesLocSorted)
#                print(vertEdgesLocSortedUnique)
                

                # for dimensioning the internal side of exterior wall 
                # vertEdgeUniLocTemp = []
                # vertEdgeSub = ReferenceArray()
                # for eL, e in zip(vertEdgesLocSortedUnique, vertEdgesSortedUnique):
                #     if eL not in vertEdgeUniLocTemp:
                #         vertEdgeUniLocTemp.append(eL)
                #         vertEdgeSub.Append(e.Reference)
                        
                # if vertEdgeSub.Size >= 2:
                #    dim = doc.Create.NewDimension(view, ref_line_interior, vertEdgeSub)
#                if vertEdgeSub.Size <= 2:
#                   print ('no sufficient edges to dimension')
                try:
                    index1=0
                    while index1<=len(vertEdgesSortedUnique)-1:
#                        print(index1, len(vertEdgesSortedUnique))
                        vertEdgeSub1 = ReferenceArray()
                        if index1 % 2==0:
                            vertEdgeSub1.Append(vertEdgesSortedUnique[index1].Reference)
                            vertEdgeSub1.Append(vertEdgesSortedUnique[index1+1].Reference)
                            dim = doc.Create.NewDimension(view, ref_line_interior, vertEdgeSub1)
#                            print (dim,'dd')
                            index1=index1+2
                        elif index1 % 2!=0:
                            index1=index1+1
#                            print('odd no.')
                            
                except Exception as e:
                    failed_data_Collected.append(['Error', e])
                    failed_data_Collected.append(['Element Name', targetWall.Name])
                    failed_data_Collected.append(['Element ID', output.linkify(targetWall.Id)])
#                    print(e)
                    
            else:
                pass
#                print("No wall found for selected wall types in view {}".format(view.Name))

def DimensionWallsInterior(offDist, extOrInt):
    # Extend dimesion for exterior side dimensions
    if extOrInt:
        #UnitUtils.ConvertToInternalUnits(500, UnitTypeId.Millimeters)
        intersectLineEndExtend = 0.5# 1 foot = 304.8 millimeters/ 500mm/offset is defined here(what ever is given in python code is divided by308 in mm in revit)
    else:
        intersectLineEndExtend = 0
#    print(intersectLineEndExtend,'intersectLineEndExtend')
    
    for view in selected_views:
        for targetWall in linear_walls_in_view:
            #For int walls which has to be dimensioned in the external face:
            if targetWall.WallType.LookupParameter("Type Name").AsString() in selected_wall_types and 'Ext' not in targetWall.WallType.LookupParameter("Type Name").AsString():
#                print ('yes-Ext/Ext')
                intersectedWallsIntFace = []#For Interior walls internal face
                intersectedWallsExtFace = []#For Interior walls external face
                
                intersectedWallsIntFace.append(targetWall)#For Interior walls internal face
                intersectedWallsExtFace.append(targetWall)#For Interior walls external face
                
                for collectedWall in linear_walls_in_view:
                    #For Interior walls internal face
                    if locToCutCrv(targetWall, wallNormal(targetWall, False), intersectLineEndExtend).Intersect(locToCutCrv(collectedWall, wallNormal(collectedWall, False), intersectLineEndExtend)) == SetComparisonResult.Overlap:
                        intersectedWallsIntFace.append(collectedWall)
                        
                for collectedWall in linear_walls_in_view:        
                    #For Interior walls external face
                    if locToCutCrv(targetWall, wallNormal(targetWall, True), intersectLineEndExtend).Intersect(locToCutCrv(collectedWall, wallNormal(collectedWall, False), intersectLineEndExtend)) == SetComparisonResult.Overlap:
                        intersectedWallsExtFace.append(collectedWall)  
                        
                #For Interior walls internal face           
                line_on_face_Int = locToCutCrv(targetWall, wallNormal(targetWall, False), intersectLineEndExtend)
                
                #For Interior walls external face
                line_on_face_Ext = locToCutCrv(targetWall, wallNormal(targetWall, True), intersectLineEndExtend)
                
                #For Interior walls internal face  
                ref_line_Int = translate_curve(line_on_face_Int, wallNormal(targetWall, False), offDist)
                
                #For Interior walls external face
                ref_line_Ext = translate_curve(line_on_face_Ext, wallNormal(targetWall, True), offDist) 
                
                #For Interior walls internal face  
                frontFaceIW = []
                
                #For Interior walls internal face  
                vertEdgesInt = []
                
                #For Interior walls internal face
                for wallInt in intersectedWallsIntFace:
                    for obj in wallInt.get_Geometry(opts):
                        if isinstance(obj, Solid):
                            for face in obj.Faces:
                                face_normal = face.ComputeNormal(UV(0.5, 0.5))
                                if isAlmostEqualTo(wallInt.Orientation, face_normal):
                                    frontFaceIW.append(face)
                                    
                            # for dimensioning the internal side of interior wall  
                            for edge in obj.Edges:
                                edgeC = edge.AsCurve()
                                edgeCNorm = edgeC.Direction.Normalize()
                                intersection_result = edgeC.Intersect(line_on_face_Int)
                                if edgeCNorm.IsAlmostEqualTo(XYZ(0, 0, 1)) or edgeCNorm.IsAlmostEqualTo(XYZ(0, 0, -1)):
                                    if intersection_result != SetComparisonResult.Disjoint:
                                        vertEdgesInt.append(edge)
                                        
                #For Interior walls external face  
                frontFaceEW = []                        
                #For Interior walls external face  
                vertEdgesExt = []     
                #For Interior walls external face  
                for wallExt in intersectedWallsExtFace:
                    for obj in wallExt.get_Geometry(opts):
                        if isinstance(obj, Solid):
                            for face in obj.Faces:
                                face_normal = face.ComputeNormal(UV(0.5, 0.5))
                                if isAlmostEqualTo(wallExt.Orientation, face_normal):
                                    frontFaceEW.append(face)
                                    
                            #For Interior walls external face   
                            for edge in obj.Edges:
                                edgeC = edge.AsCurve()
                                edgeCNorm = edgeC.Direction.Normalize()
                                intersection_result = edgeC.Intersect(line_on_face_Ext)
                                if edgeCNorm.IsAlmostEqualTo(XYZ(0, 0, 1)) or edgeCNorm.IsAlmostEqualTo(XYZ(0, 0, -1)):
                                    if intersection_result != SetComparisonResult.Disjoint:
                                        vertEdgesExt.append(edge)
                                        

#***************************************************************************************************************************************************
                #For Interior walls internal  face                        
                vertEdgesLocInt = []# Get location as sorting parameters to remove duplicates

                for v in vertEdgesInt:
                    vLocInt = v.AsCurve().GetEndPoint(0).X + v.AsCurve().GetEndPoint(0).Y
                    vertEdgesLocInt.append(round(vLocInt,7))
    
                #For Interior walls internal  face     
                vertEdgesSortedInt = [x for _,x in sorted(zip(vertEdgesLocInt,vertEdgesInt))]
                vertEdgesLocSortedInt = sorted(vertEdgesLocInt)
                
                #For Interior walls internal  face.
                #For removing corner conditions of interior walls.
                #Step 1 - Collect faces of edges of all walls
                FacesIWs=[]
                for walls in intersectedWallsIntFace:
                    for obj in walls.get_Geometry(opts):
                        if isinstance(obj, Solid):
                            for face in obj.Faces:
                                face_normal = face.ComputeNormal(UV(0.5, 0.5))
                                if isAlmostEqualTo(wallExt.Orientation, face_normal) or isAlmostEqualTo(wallExt.Orientation.Negate(), face_normal):
                                    pass
                                else:
                                    FacesIWs.append(face)

                # finding the corner edges which have to be removed.
                if len(vertEdgesSortedInt)>=4:
                    if vertEdgesLocSortedInt[1]==vertEdgesLocSortedInt[2]:
                        if vertEdgesSortedInt[1].GetFace(0) in FacesIWs or vertEdgesSortedInt[1].GetFace(1) in FacesIWs:
                            if vertEdgesSortedInt[2].GetFace(0) in FacesIWs or vertEdgesSortedInt[2].GetFace(1) in FacesIWs:
                                vertEdgesSortedInt.remove(vertEdgesSortedInt[1])
                                vertEdgesSortedInt.remove(vertEdgesSortedInt[1])# when index 1 is removed 2nd index will become 1st index now
                                vertEdgesLocSortedInt.remove(vertEdgesLocSortedInt[1])
                                vertEdgesLocSortedInt.remove(vertEdgesLocSortedInt[1])# when index 1 is removed 2nd index will become 1st index now                        
                if len(vertEdgesSortedInt)>=4:#if the length is still greater than 4 after removing edges
                    if vertEdgesLocSortedInt[len(vertEdgesLocSortedInt)-2]==vertEdgesLocSortedInt[len(vertEdgesLocSortedInt)-3]:
                        if vertEdgesSortedInt[len(vertEdgesLocSortedInt)-2].GetFace(0) in FacesIWs or vertEdgesSortedInt[len(vertEdgesLocSortedInt)-2].GetFace(1) in FacesIWs:
                            if vertEdgesSortedInt[len(vertEdgesLocSortedInt)-3].GetFace(0) in FacesIWs or vertEdgesSortedInt[len(vertEdgesLocSortedInt)-3].GetFace(1) in FacesIWs:
                                vertEdgesSortedInt.remove(vertEdgesSortedInt[len(vertEdgesSortedInt)-2])
                                vertEdgesSortedInt.remove(vertEdgesSortedInt[len(vertEdgesSortedInt)-2])# when 2nd last index is removed 3rd last index will become 2nd last index now  
                                vertEdgesLocSortedInt.remove(vertEdgesLocSortedInt[len(vertEdgesLocSortedInt)-2])
                                vertEdgesLocSortedInt.remove(vertEdgesLocSortedInt[len(vertEdgesLocSortedInt)-2])# when 2nd last index is removed 3rd last index will become 2nd last index now                 
                # for i in vertEdgesSortedInt:
                #     print('final vertical edges interior',i)    
                if len(vertEdgesSortedInt)>=4:#if the length is still greater than 4 after removing edges
                    if vertEdgesLocSortedInt[0]==vertEdgesLocSortedInt[1]:
                        # print('yes')
                        vertEdgesSortedInt.remove(vertEdgesSortedInt[0])
                        vertEdgesSortedInt.remove(vertEdgesSortedInt[0])
                        vertEdgesLocSortedInt.remove(vertEdgesLocSortedInt[0])
                        vertEdgesLocSortedInt.remove(vertEdgesLocSortedInt[0])
                # for i in vertEdgesSortedInt:
                #     print('final vertical edges interiorrrrr',i)                 
                
                
                #For Interior walls internal  face   
                vertEdgesSortedUniqueInt=[]#Filetring out only the unique edges
                vertEdgesLocSortedUniqueInt=[]
                for i in vertEdgesLocSortedInt:
                    if i not in vertEdgesLocSortedUniqueInt:
                        indexd=vertEdgesLocSortedInt.index(i)
                        vertEdgesLocSortedUniqueInt.append(i)
                        vertEdgesSortedUniqueInt.append(vertEdgesSortedInt[indexd])
                
                #  #For Interior walls internal  face   
                # InternalSideDim = ReferenceArray()     
                # for i in vertEdgesSortedUniqueInt:
                #     InternalSideDim.Append(i.Reference)
                # dim = doc.Create.NewDimension(view, ref_line_Int, InternalSideDim)
                   
                try:
                    index1=0
                    while index1<=len(vertEdgesSortedUniqueInt)-1:
    #                   print(index1, len(vertEdgesSortedUniqueInt))
                        InternalSideDim = ReferenceArray()
                        if index1 % 2==0:
                            InternalSideDim.Append(vertEdgesSortedUniqueInt[index1].Reference)
                            InternalSideDim.Append(vertEdgesSortedUniqueInt[index1+1].Reference)
                            dim = doc.Create.NewDimension(view, ref_line_Int, InternalSideDim)
    #                       print (dim,'dd')
                            index1=index1+2
                        elif index1 % 2!=0:
                            index1=index1+1
    #                       print('odd no.')
                            
                except Exception as e:
                    failed_data_Collected.append(['Error', e])
                    failed_data_Collected.append(['Element Name', targetWall.Name])
                    failed_data_Collected.append(['Element ID', output.linkify(targetWall.Id)])                    
#                    print(e)




t = Transaction(doc, "Dimension internal Wall")
#start transaction
t.Start()

failed_data_Collected_1=[]

try:
    DimensionWallsExterior(1.2, True)
    DimensionWallsInterior(1.2, True)
except Exception as e:
    failed_data_Collected_1.append(['Error', e])
    
t.Commit()

if len(failed_data_Collected)!=0:
    output.print_table(table_data=failed_data_Collected, title='Exceptions', columns=["Error"])
    
if len(failed_data_Collected_1)!=0:
    output.print_table(table_data=failed_data_Collected, title='Exceptions_1', columns=["Error"])

end_time = time.time()
runtime = end_time - start_time
print("Runtime : {} seconds".format(runtime))

