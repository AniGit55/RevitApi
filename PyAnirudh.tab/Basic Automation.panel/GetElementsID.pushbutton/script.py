# -*- coding: utf-8 -*-

__title__ = "Get Elements ID"
__doc__ = """Version = 1.0
Description:
Retrieve Element ID, Name and Category
_____________________________________________________________________
How-to:
-> Select the Element 
-> Click on the button """

#⬇️ IMPORTS
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog

#📦 VARIABLES
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

#🎯 MAIN
selected_element_ids = uidoc.Selection.GetElementIds()

if not selected_element_ids:
    TaskDialog.Show("Selected Elements", "No elements selected.")
else:
    selected_elements = [doc.GetElement(eid) for eid in selected_element_ids]

    display_info = []
    for e in selected_elements:
        try:
            name = ""
            if hasattr(e, "Name"):
                name = e.Name
            elif e.GetType():
                name = e.GetType().Name

            category = e.Category.Name if e.Category else "No Category"
            element_id = e.Id.IntegerValue

            display_info.append("ID: {}, Name: {}, Category: {}".format(element_id, name, category))
        except Exception as ex:
            display_info.append("ID: {}, Error: {}".format(e.Id, str(ex)))

    # Final output
    if display_info:
        TaskDialog.Show("Selected Elements", "\n".join(display_info))
    else:
        TaskDialog.Show("Selected Elements", "No valid information to display.")
