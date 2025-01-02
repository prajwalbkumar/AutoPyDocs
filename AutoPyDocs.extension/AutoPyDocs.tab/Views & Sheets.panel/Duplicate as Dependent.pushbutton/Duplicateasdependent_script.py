# -*- coding: utf-8 -*-   
__title__ = "Duplicate as Dependent"
__author__ = "abhiramnair"

# Imports
import clr
import os
import time
import xlrd

clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from pyrevit import revit, forms, script
import Autodesk.Revit.DB as DB
from Extract.RunData import get_run_data
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from System.Windows.Forms import Form, Label, Button, DataGridView, DockStyle, DialogResult
import time
import os
from datetime import datetime



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



# Define all possible view types for selection
all_view_types = {
    'Floor Plans': ViewType.FloorPlan,
    'Reflected Ceiling Plans': ViewType.CeilingPlan,
    'Area Plans': ViewType.AreaPlan,
    'Structural Plans': ViewType.EngineeringPlan,
    'Sections': ViewType.Section,
    'Elevations': ViewType.Elevation
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
filtered_views = [
    view for view in views_collector
    if view.ViewType in selected_view_types and not view.IsTemplate
]

if not filtered_views:
    forms.alert("No views of the selected types were found in the document.", exitscript=True)

# Collect all Viewport elements in the document
viewports_collector = FilteredElementCollector(doc).OfClass(Viewport)
views_on_sheets_ids = {viewport.ViewId for viewport in viewports_collector}
filtered_views_on_sheets = [
    view for view in filtered_views if view.Id in views_on_sheets_ids
]

if not filtered_views_on_sheets:
    forms.alert("No views of the selected types were found on sheets in the document.", exitscript=True)

view_dict = {view.Name: view for view in filtered_views_on_sheets}

# Show a dialog to select specific views
selected_view_names = forms.SelectFromList.show(
    sorted(view_dict.keys()),
    title='Select Views',
    multiselect=True
)

if not selected_view_names:
    forms.alert("No views were selected.", exitscript=True)

selected_views = [view_dict[name] for name in selected_view_names]

# Define the form class for selecting number of duplicates
class InteractiveViewDuplicatorForm(Form):
    def __init__(self, views):
        super(InteractiveViewDuplicatorForm, self).__init__()
        self.Text = "Duplicate Views as Dependents"
        self.Width = 500
        self.Height = 400

        label = Label()
        label.Text = "Specify the number of dependent views for each view:"
        label.Dock = DockStyle.Top
        self.Controls.Add(label)

        self.data_grid = DataGridView()
        self.data_grid.Dock = DockStyle.Fill
        self.data_grid.ColumnCount = 2
        self.data_grid.Columns[0].Name = "View Name"
        self.data_grid.Columns[1].Name = "Dependent Views (Enter Integer)"
        for view in views:
            self.data_grid.Rows.Add(view.Name, "1")
        self.Controls.Add(self.data_grid)

        ok_button = Button()
        ok_button.Text = "OK"
        ok_button.Dock = DockStyle.Bottom
        ok_button.Click += self.ok_clicked
        self.Controls.Add(ok_button)

        cancel_button = Button()
        cancel_button.Text = "Cancel"
        cancel_button.Dock = DockStyle.Bottom
        cancel_button.Click += self.cancel_clicked
        self.Controls.Add(cancel_button)

        self.result = None

    def ok_clicked(self, sender, args):
        self.result = []
        for row in self.data_grid.Rows:
            view_name = row.Cells[0].Value
            dep_count = row.Cells[1].Value
            if view_name and dep_count.isdigit():
                self.result.append((view_name, int(dep_count)))
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()



new_names = []
# Show the form
form = InteractiveViewDuplicatorForm(selected_views)
if form.ShowDialog() == DialogResult.OK:
    selected_views_and_counts = form.result

    # Start Transaction
    t = Transaction(doc, "Duplicate Views as Dependents")
    t.Start()

    view_dict = {view.Name: view for view in filtered_views}

    # Function to generate a unique name
    def get_unique_name(base_name, existing_names):
        suffix = 1
        new_name = "{}_Dependent_{}".format(base_name, suffix)
        while new_name in existing_names:
            suffix += 1
            new_name = "{}_Dependent_{}".format(base_name, suffix)
        return new_name

    # Collect existing view names to check for conflicts
    existing_view_names = {view.Name for view in filtered_views}

    for view_name, num_dependents in selected_views_and_counts:
        if view_name in view_dict:
            view = view_dict[view_name]

            # Create the dependent views
            for _ in range(num_dependents):
                unique_name = get_unique_name(view.Name, existing_view_names)

                # Duplicate as dependent
                dependent_view = view.Duplicate(ViewDuplicateOption.AsDependent)
                dependent_element = doc.GetElement(dependent_view)
                dependent_element.Name = unique_name
                if dependent_element:
                    new_names.append((dependent_element.Name, dependent_element.Id))

                # Add the new name to the existing names set
                existing_view_names.add(unique_name)

    t.Commit()
    # Record the end time
    counter = len(new_names)
    end_time = time.time()
    runtime = end_time - start_time

    run_result = "Tool ran successfully"
    element_count = counter if counter else 0
    error_occured = "Nil"
    get_run_data(__title__, runtime, element_count, manual_time, run_result, error_occured)
            

    if new_names:
        output.print_md(header_data)
        #Print the header for skipped views
        output.print_md("## VIEW DUPLICATED AS DEPENDENT  😊")  # Markdown Heading 2 
        output.print_md("---")  # Markdown Line Break
        output.print_md(" The following views have successfully duplicated as dependent")  # Print a Line
        # Create a table to display the skipped views
        output.print_table(table_data=new_names, columns=["VIEW NAME", "VIEW ID"])  # Print a Table

    forms.show_balloon("Dependent views created successfully.", "Success")
else:
    forms.show_balloon("Operation canceled by user.", "Canceled")
    # Ensure to print or log anything else if necessary

runtime = time.time() - start_time
print("Script completed successfully in {:.2f} seconds.".format(runtime))