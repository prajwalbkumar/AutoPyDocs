# -*- coding: utf-8 -*-
'''Tag Walls'''
__title__ = "Tag Walls"
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

# Step 1: Allow user to select desired wall types
wall_types_collector = [wt for wt in FilteredElementCollector(doc).OfClass(WallType) if isinstance(wt, WallType)]
wall_type_options = {
    wt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(): wt
    for wt in FilteredElementCollector(doc).OfClass(WallType)
    if wt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) is not None
}


selected_wall_types = forms.SelectFromList.show(
    sorted(wall_type_options.keys()),
    title="Select Wall Types to Tag",
    multiselect=True
)
if not selected_wall_types:
    forms.alert("No wall types selected. Exiting script.", exitscript=True)

selected_wall_type_ids = [wall_type_options[name].Id for name in selected_wall_types]



# Define all possible view types for selection
all_view_types = {
    'Floor Plans': ViewType.FloorPlan,
    'Reflected Ceiling Plans': ViewType.CeilingPlan,
    'Area Plans': ViewType.AreaPlan,
    'Structural Plans' : ViewType.EngineeringPlan
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



# Select tag family type
tag_families = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_WallTags)
tag_family_options = {fs.Family.Name: fs for fs in tag_families if fs.IsActive}
if not tag_family_options:
    forms.alert("No wall tag families found in the project.", exitscript=True)

selected_tag_family_name = forms.SelectFromList.show(
    sorted(tag_family_options.keys()),
    title="Select Wall Tag Family",
    multiselect=False
)
if not selected_tag_family_name:
    forms.alert("No tag family selected. Exiting script.", exitscript=True)
selected_tag_family = tag_family_options[selected_tag_family_name]




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

walls_tagged = []

try:
    t = Transaction(doc, "Tag Walls")
    t.Start()
    geom_view = None

    # Collect all linked instances
    linked_instance = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    if linked_instance:
        documentation_file = forms.alert("Is this a Documentation File or a Live File", warn_icon=False, options=["Documentation File", "Live File"])

        if not documentation_file:
            forms.alert("No file option selected. Exiting script.", exitscript=True)

        if documentation_file == "Documentation File":
            link_name = []
            for link in linked_instance:
                link_name.append(link.Name)

            ar_instance_name = forms.SelectFromList.show(link_name, title = "Select AR Linked File", width=600, height=600, button_name="Select File", multiselect=False)

            if not ar_instance_name:
                script.exit()

            for link in linked_instance:
                if ar_instance_name == link.Name:
                    ar_instance = link
                    break

            ar_doc = ar_instance.GetLinkDocument()
            transform = ar_instance.GetTransform()
            if not ar_doc:
                forms.alert("No instance found of the selected AR File.\n"
                            "Use Manage Links to Load the Link in the File!", title = "Link Missing", warn_icon = True)
                script.exit()

            walls_in_view = [wall for wall in FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType() if wall.GetTypeId() in selected_wall_type_ids]
            views_in_link = FilteredElementCollector(ar_doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()

            for link_view in views_in_link:
                if link_view.Name == "3D-Navisworks-Export":
                    geom_view = link_view
                    break
        
        else:
            linked_instance = None
            ar_doc = doc

    else:
        linked_instance = None
        ar_doc = doc
        
    for view in selected_views:
        if linked_instance:
            walls_in_view = [wall for wall in FilteredElementCollector(ar_doc, geom_view.Id).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType()if wall.GetTypeId() in selected_wall_type_ids]
        else:
             walls_in_view = [wall for wall in FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType()if wall.GetTypeId() in selected_wall_type_ids]


        if not walls_in_view:
            error_in_view = True

        # Collect all tags in the selected view
        existing_tags = FilteredElementCollector(doc, view.Id).OfClass(IndependentTag).ToElements()

        # Collect IDs of elements already tagged in the view
        already_tagged_wall_ids = set()
        for tag in existing_tags:
            tagged_references = tag.GetTaggedReferences()
            for tagged_reference in tagged_references:
                tagged_element = doc.GetElement(tagged_reference.ElementId)
                if isinstance(tagged_element, Wall):  # Check if the tagged element is a wall
                    already_tagged_wall_ids.add(tagged_element.Id)


        wall_dict = {wall: (wall.Location.Curve.GetEndPoint(1).Z + wall.LookupParameter("Base Offset").AsDouble()) for wall in walls_in_view}
        linked_walls = []
        
        for wall in walls_in_view:
            if wall.Id in already_tagged_wall_ids:
                continue

            if ar_doc.GetElement(wall.LevelId).Name != view.GenLevel.Name:
                continue

            linked_wall_id = wall.Id
            if linked_instance:
                    
                ref = Reference(wall).CreateLinkReference(ar_instance)
            else:
                ref = Reference(wall)
            s_factor = view.Scale
            offset_distance = 3 * s_factor / 100

            wall_location = wall.Location
            if isinstance(wall_location, LocationCurve):
                curve = wall_location.Curve
            
            if curve:
                # Get the start and end points of the curve
                start_point = curve.GetEndPoint(0)
                end_point = curve.GetEndPoint(1)
                wall_direction = (end_point - start_point).Normalize()
                offset_vector = XYZ(wall_direction.Y, wall_direction.X, 0) * offset_distance
            
                mid_point = curve.Evaluate(0.5, True)
                tag_point = mid_point + offset_vector

            # Create the tag at the midpoint of the boundary segment
            tag = IndependentTag.Create(
                doc,
                view.Id,
                ref,
                True,
                TagMode.TM_ADDBY_CATEGORY,
                TagOrientation.Horizontal,
                tag_point
            )

            if tag:
                walls_tagged.append(wall.Id)
        # Set the selected tag family
            tag.ChangeTypeId(selected_tag_family.Id)

            # Adjust the tag position explicitly
            tag.TagHeadPosition = tag_point

    t.Commit()
    total_walls_tagged = len(walls_tagged)
    # Record the end time
    end_time = time.time()
    runtime = end_time - start_time

    run_result = "Tool ran successfully"
    element_count = total_walls_tagged
    error_occured = "Nil"
    get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)


except Exception as e:
    # Print error message
    print("Error occurred: {}".format(str(e)))

    # Record the end time and runtime
    end_time = time.time()
    runtime = end_time - start_time

    # Log the error details
    error_occured = "Error occurred: {}".format(str(e))
    run_result = "Error"
    element_count = 10

    # Function to log run data in case of error
    get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)

    t.RollBack()

if walls_tagged:
    output.print_md(header_data) 
    output.print_md("### Total Walls Tagged: **{}**".format(total_walls_tagged))

    print("\n\n")
    output.print_md("---")

