# -*- coding: utf-8 -*-
__title__ = "Auto Graphic Overrides"
__author__ = "Anirudh"
__doc__ = "Applies graphic overrides to non-structural walls and all windows in the active view."

from Autodesk.Revit.DB import *
from pyrevit import revit, DB
from pyrevit import script
from System.Drawing import Color

doc = revit.doc
uidoc = revit.uidoc
view = uidoc.ActiveView

if not view.CanBePrinted:
    script.exit("Please open a printable view (not a schedule, legend, or sheet).")

print("Active view: {0} (Type: {1})".format(view.Name, view.ViewType))

t = Transaction(doc, "Auto Graphic Overrides")
t.Start()

# Categories to process
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

# Graphic override settings
gray = Color.FromArgb(180, 180, 180)
revit_color = DB.Color(gray.R, gray.G, gray.B)

override_settings = OverrideGraphicSettings()
override_settings.SetProjectionLineColor(revit_color)
override_settings.SetProjectionLineWeight(1)
override_settings.SetSurfaceTransparency(60)

# Helper function to check if a wall is structural
def is_structural_wall(wall):
    try:
        wall_type = doc.GetElement(wall.GetTypeId())
        if not wall_type:
            print("No wall type found for wall ID={0}".format(wall.Id.IntegerValue))
            return True  # Conservatively assume structural if type is missing

        # Get wall type name for debugging
        type_name = wall_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if wall_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "Unknown"

        # Check Structural Checkbox (type-level)
        struct_param = wall_type.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT)
        is_structural_type = struct_param.AsInteger() == 1 if struct_param else False
        struct_param_found = struct_param is not None

        # Check Structural Checkbox (instance-level)
        is_structural_instance = False
        struct_instance_param = wall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT)
        if struct_instance_param:
            is_structural_instance = struct_instance_param.AsInteger() == 1
        else:
            struct_instance_param = None

        # Check Structural Usage
        struct_usage_param = wall_type.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_USAGE_PARAM)
        struct_usage_value = struct_usage_param.AsInteger() if struct_usage_param else -1
        struct_usage_name = {0: "Non-Bearing", 1: "Bearing", 2: "Shear", 3: "Structural Combined"}.get(struct_usage_value, "Unknown")

        # Check Wall Function
        function_param = wall_type.get_Parameter(BuiltInParameter.FUNCTION_PARAM)
        function_value = function_param.AsInteger() if function_param else -1
        function_name = {0: "Interior", 1: "Exterior", 2: "Foundation", 3: "Retaining", 4: "Soffit", 5: "Core-Shaft"}.get(function_value, "Unknown")

        # Check Structural Material
        has_struct_material = False
        struct_material_param = wall_type.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if struct_material_param and struct_material_param.HasValue:
            material = doc.GetElement(struct_material_param.AsElementId())
            has_struct_material = material is not None and any(mat in material.Name for mat in ["Concrete", "Steel"])  # Adjust material names as needed

        # Debug: Print parameter values
        print("Checking wall ID={0}, Type={1}, Structural (Type)={2}, Structural Param Found={3}, Structural (Instance)={4}, Structural Usage={5} ({6}), Wall Function={7} ({8}), Has Structural Material={9}".format(
            wall.Id.IntegerValue, type_name, is_structural_type, struct_param_found, is_structural_instance, struct_usage_value, struct_usage_name, function_value, function_name, has_struct_material))

        # Wall is structural if:
        # - Structural checkbox is checked (instance-level, since type-level is not found)
        # - Structural Usage is Bearing, Shear, or Structural Combined
        # - Wall Function is Foundation or Retaining
        # - Has a structural material
        if is_structural_instance:
            return True
        if struct_usage_value in [1, 2, 3]:  # Bearing, Shear, or Structural Combined
            return True
        if function_value in [2, 3]:  # Foundation or Retaining
            return True
        if has_struct_material:
            return True

        return False
    except Exception as e:
        print("Error checking structural status for wall ID={0}: {1}".format(wall.Id.IntegerValue, e))
        return True  # Conservatively assume structural on error

# Collect structural walls to exclude
structural_wall_ids = set()
wall_elements = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
for wall in wall_elements:
    if is_structural_wall(wall):
        structural_wall_ids.add(wall.Id)
        # Debug: Print structural wall info
        wall_type = doc.GetElement(wall.GetTypeId())
        type_name = wall_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if wall_type and wall_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "Unknown"
        print("Identified structural wall: ID={0}, Type={1}".format(wall.Id.IntegerValue, type_name))

# Apply overrides
for bic in target_categories:
    try:
        print("Processing category: {0}".format(bic.ToString()))
        elems = FilteredElementCollector(doc, view.Id).OfCategory(bic).WhereElementIsNotElementType().ToElements()
        print("Found {0} elements in category {1}".format(len(elems), bic.ToString()))

        for elem in elems:
            # Check visibility
            if not elem.CanBeHidden(view):
                elem_type = doc.GetElement(elem.GetTypeId())
                elem_type_name = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if elem_type and elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "Unknown"
                print("Skipping element ID={0}, Type={1}: Cannot be hidden in view".format(elem.Id.IntegerValue, elem_type_name))
                continue
            if not elem.get_BoundingBox(view):
                elem_type = doc.GetElement(elem.GetTypeId())
                elem_type_name = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if elem_type and elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "Unknown"
                print("Skipping element ID={0}, Type={1}: No bounding box in view".format(elem.Id.IntegerValue, elem_type_name))
                continue

            # Handle walls
            if isinstance(elem, Wall):
                if elem.Id in structural_wall_ids:
                    wall_type = doc.GetElement(elem.GetTypeId())
                    type_name = wall_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if wall_type and wall_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "Unknown"
                    print("Skipping structural wall: ID={0}, Type={1}".format(elem.Id.IntegerValue, type_name))
                    continue
                # Apply override to non-structural walls
                view.SetElementOverrides(elem.Id, override_settings)
                elem_type = doc.GetElement(elem.GetTypeId())
                elem_type_name = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if elem_type and elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "Unknown"
                print("Applied override to: ID={0}, Category={1}, Type={2}".format(elem.Id.IntegerValue, bic.ToString(), elem_type_name))
                continue

            # Handle windows (apply override to all windows)
            if bic == BuiltInCategory.OST_Windows:
                host = getattr(elem, "Host", None)
                host_type = "Unknown"
                host_id = "No Host"
                if host and isinstance(host, Wall):
                    host_wall_type = doc.GetElement(host.GetTypeId())
                    host_type = host_wall_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if host_wall_type and host_wall_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "Unknown"
                    host_id = str(host.Id.IntegerValue)
                # Apply override to all windows
                view.SetElementOverrides(elem.Id, override_settings)
                elem_type = doc.GetElement(elem.GetTypeId())
                elem_type_name = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if elem_type and elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "Unknown"
                print("Applied override to window: ID={0}, Type={1}, Host Wall Type={2}, Host Wall ID={3}".format(
                    elem.Id.IntegerValue, elem_type_name, host_type, host_id))
                continue

            # Apply override to other categories
            view.SetElementOverrides(elem.Id, override_settings)
            elem_type = doc.GetElement(elem.GetTypeId())
            elem_type_name = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if elem_type and elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "Unknown"
            print("Applied override to: ID={0}, Category={1}, Type={2}".format(elem.Id.IntegerValue, bic.ToString(), elem_type_name))

    except Exception as e:
        print("Error in category {0}: {1}".format(bic, e))

t.Commit()
script.get_output().print_md("✅ **Overrides applied. Non-structural walls and all windows affected.**")