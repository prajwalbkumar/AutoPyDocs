# -*- coding: utf-8 -*-     
'''Clear Annotations'''
__title__ = "Clear Annotations"
__author__ = "prajwalbkumar"


# Imports
import math
import os
import time
from datetime import datetime
from Extract.RunData import get_run_data
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

# Record the start time
start_time = time.time()
manual_time = 45

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

view = doc.ActiveView.Id
warning = forms.alert("Running this would delete all annotaions in the active view. \n 🚨🚨🚨🚨🚨🚨🚨🚨", title="Delete Annotations", options=["DELETE ALL ANNOTATIONS", "NO DO NOT DELETE"])

if not warning or warning == "NO DO NOT DELETE":
    script.exit()

else:
    warning = forms.alert("You really sure???? \n 🚨🚨🚨🚨🚨🚨🚨🚨", title="Delete Annotations", options=["DELETE ALL ANNOTATIONS NOW", "NOPE"])
    if not warning or warning == "NOPE":
        script.exit()

t = Transaction(doc, "Deleting Annotations")
t.Start()

elements = FilteredElementCollector(doc, view).OfCategory(BuiltInCategory.OST_Dimensions).ToElementIds()
for element in elements:
    doc.Delete(element)

elements = FilteredElementCollector(doc, view).OfCategory(BuiltInCategory.OST_SpotElevations).ToElementIds()
for element in elements:
    doc.Delete(element)

elements = FilteredElementCollector(doc, view).OfCategory(BuiltInCategory.OST_WallTags).ToElementIds()
for element in elements:
    doc.Delete(element)

t.Commit()
