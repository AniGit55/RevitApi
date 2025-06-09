# -*- coding: utf-8 -*-
__title__ = "Reset Graphic Overrides"
__author__ = "Anirudh"
__doc__ = "Removes graphic overrides applied to non-structural categories in the active view."

from Autodesk.Revit.DB import *
from pyrevit import revit, DB
from pyrevit import script

doc = revit.doc
uidoc = revit.uidoc
view = uidoc.ActiveView

if not view.CanBePrinted:
    script.exit("Please open a printable view (not a schedule, legend, or sheet).")

print("Active view: {0} (Type: {1})".format(view.Name, view.ViewType))

t = Transaction(doc, "Reset Graphic Overrides")
t.Start()

# Categories to reset (match those in AutoGraphicOverrides.py)
target_categories = [
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_LightingFixtures
]

# Empty override settings to reset to default
reset_settings = OverrideGraphicSettings()

# Reset overrides for each category
for bic in target_categories:
    try:
        print("Processing category: {0}".format(bic.ToString()))
        elems = FilteredElementCollector(doc, view.Id).OfCategory(bic).WhereElementIsNotElementType().ToElements()
        print("Found {0} elements in category {1}".format(len(elems), bic.ToString()))

        for elem in elems:
            # Apply reset (remove overrides)
            view.SetElementOverrides(elem.Id, reset_settings)
            elem_type = doc.GetElement(elem.GetTypeId())
            elem_type_name = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if elem_type and elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "Unknown"
            print("Reset override for: ID={0}, Category={1}, Type={2}".format(elem.Id.IntegerValue, bic.ToString(), elem_type_name))

    except Exception as e:
        print("Error in category {0}: {1}".format(bic, e))

t.Commit()
script.get_output().print_md("✅ **Graphic overrides reset for all elements in the active view.**")