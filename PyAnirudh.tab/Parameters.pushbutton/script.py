# -*- coding: utf-8 -*-
__title__   = "Para Manager"
__doc__     = """Version = 1.0
________________________________________________________________
"""
import os
import clr
import re
import System

clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xml")
clr.AddReference('Microsoft.Office.Interop.Excel')

from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from System.Windows.Markup import XamlReader
from System.Collections.Generic import List
from System.Xml import XmlReader
from System.IO import StringReader
from System.Windows import MessageBox, MessageBoxButton
from System.Runtime.InteropServices import Marshal
from pyrevit import script
from pyrevit import forms
from System.Threading import Thread

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
selection_ids = uidoc.Selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selection_ids if doc.GetElement(eid) is not None]

def safe_set_value(param, value):
    st = param.StorageType
    try:
        if st == StorageType.String:
            param.Set(str(value))
        elif st == StorageType.Integer:
            valstr = str(value).strip().lower()
            if valstr in ["yes", "true", "1"]:
                param.Set(1)
            elif valstr in ["no", "false", "0"]:
                param.Set(0)
            elif valstr.isdigit():
                param.Set(int(valstr))
            else:
                pass # ignore non-numeric or enum labels (like "Vertical", "By Type", etc.)
        elif st == StorageType.Double:
            try:
                valstr = str(value)
                val_num = float(re.findall(r"[-+]?\d*\.\d+|\d+", valstr.replace(",", ""))[0])
                param.Set(val_num)
            except Exception:
                pass
        elif st == StorageType.ElementId:
            try:
                eid = int(value)
                param.Set(ElementId(eid))
            except Exception:
                pass
    except Exception:
        pass

# Mapping from Revit API enum to user-friendly labels
BUILTINPARAMGROUP_TO_LABEL = {
    BuiltInParameterGroup.INVALID: "Other",
    BuiltInParameterGroup.PG_DATA: "Data",
    BuiltInParameterGroup.PG_CONSTRUCTION: "Construction",
    BuiltInParameterGroup.PG_IDENTITY_DATA: "Identity Data",
}

group_pairs = [
    ("Analysis Results", "PG_ANALYSIS_RESULTS"),
    ("Analytical Alignment", "PG_ANALYTICAL_ALIGNMENT"),
    ("Analytical Model", "PG_ANALYTICAL_MODEL"),
    ("Constraints", "PG_CONSTRAINTS"),
    ("Construction", "PG_CONSTRUCTION"),
    ("Data", "PG_DATA"),
    ("Dimensions", "PG_GEOMETRY"),
    ("Division Geometry", "PG_DIVISION_GEOMETRY"),
    ("Electrical", "PG_ELECTRICAL"),
    ("Electrical - Circuiting", "PG_ELECTRICAL_CIRCUITING"),
    ("Electrical - Lighting", "PG_ELECTRICAL_LIGHTING"),
    ("Electrical - Loads", "PG_ELECTRICAL_LOADS"),
    ("Electrical Engineering", "PG_ELECTRICAL_ENGINEERING"),
    ("Energy Analysis", "PG_ENERGY_ANALYSIS"),
    ("Fire Protection", "PG_FIRE_PROTECTION"),
    ("Forces", "PG_FORCES"),
    ("General", "PG_GENERAL"),
    ("Graphics", "PG_GRAPHICS"),
    ("Green Building Properties", "PG_GREEN_BUILDING"),
    ("Identity Data", "PG_IDENTITY_DATA"),
    ("IFC Parameters", "PG_IFC"),
    ("Layers", "PG_LAYER"),
    ("Materials and Finishes", "PG_MATERIALS"),
    ("Mechanical", "PG_MECHANICAL"),
    ("Mechanical - Flow", "PG_MECHANICAL_FLOW"),
    ("Mechanical - Loads", "PG_MECHANICAL_LOADS"),
    ("Model Properties", "PG_MODEL_PROPERTIES"),
    ("Moments", "PG_MOMENTS"),
    ("Other", "INVALID"),
    ("Overall Legend", "PG_OVERALL_LEGEND"),
    ("Phasing", "PG_PHASING"),
    ("Photometrics", "PG_PHOTOMETRICS"),
    ("Plumbing", "PG_PLUMBING"),
    ("Primary End", "PG_PRIMARY_END"),
    ("Rebar Set", "PG_REBAR_ARRAY"),
    ("Releases / Member Forces", "PG_RELEASES_MEMBER_FORCES"),
    ("Secondary End", "PG_SECONDARY_END"),
    ("Segments and Fittings", "PG_SEGMENTS_FITTINGS"),
    ("Set", "PG_SET"),
    ("Slab Shape Edit", "PG_SLAB_SHAPE_EDIT"),
    ("Structural", "PG_STRUCTURAL"),
    ("Structural Analysis", "PG_STRUCTURAL_ANALYSIS"),
    ("Text", "PG_TEXT"),
    ("Title Text", "PG_TITLE"),
    ("Visibility", "PG_VISIBILITY"),
]

GROUP_LABEL_TO_ENUM = {}
for label, enum_attr in group_pairs:
    if hasattr(BuiltInParameterGroup, enum_attr):
        GROUP_LABEL_TO_ENUM[label] = getattr(BuiltInParameterGroup, enum_attr)
# For display:
BUILTINPARAMGROUP_TO_LABEL = {v: k for k, v in GROUP_LABEL_TO_ENUM.items()}

class ParameterVM(object):
    def __init__(self, param, group_under="Other", discipline="Common", insttype="Instance"):
        self.param = param
        self.IsSelected = True
        self.Name = param.Definition.Name
        self.Type = "Shared Parameter" if param.IsShared else "Project Parameter"
        self.Discipline = discipline
        self.PType = self.get_param_type(param)
        self.GroupUnder = self.get_group_under(param)
        self.InstType = insttype
        self.Editable = not param.IsReadOnly
        self.Value = self.get_val_for_edit()

    def get_val_for_edit(self):
        try:
            st = self.param.StorageType
            if st == StorageType.String:
                return self.param.AsString() or ""
            elif st == StorageType.Integer:
                val_str = self.param.AsValueString()
                if val_str:
                    return val_str
                return str(self.param.AsInteger())
            elif st == StorageType.Double:
                val_str = self.param.AsValueString()
                if val_str:
                    return val_str
                return str(self.param.AsDouble())
            elif st == StorageType.ElementId:
                eid = self.param.AsElementId()
                if eid == ElementId.InvalidElementId:
                    return ""
                element = doc.GetElement(eid)
                if element and hasattr(element, "Name"):
                    return element.Name
                elif element and hasattr(element, "LookupParameter"):
                    name_param = element.LookupParameter("Name")
                    if name_param:
                        return name_param.AsString()
                val_str = self.param.AsValueString()
                if val_str:
                    return val_str
                return str(eid.IntegerValue)
            else:
                return ""
        except Exception as ex:
            return "Error: {}".format(ex)

    def get_param_type(self, param):
        s = param.StorageType
        definition = param.Definition
        gd_str = ""
        try:
            if hasattr(definition, "ParameterType"):
                gd = definition.ParameterType
                gd_str = str(gd)
            else:
                gd = None
                gd_str = ""
        except:
            gd_str = ""
        if s == StorageType.String:
            return "Text"
        elif s == StorageType.Integer:
            if "yesno" in gd_str.lower():
                return "Yes/No"
            if definition.Name.lower() in ["is visible", "enabled", "is checked", "show"]:
                return "Yes/No"
            return "Integer"
        elif s == StorageType.Double:
            return "Number"
        elif s == StorageType.ElementId:
            return "ElementId"
        else:
            return gd_str or "Other"

    def get_group_under(self, param):
        if hasattr(param.Definition, "ParameterGroup"):
            group_enum = param.Definition.ParameterGroup
            return BUILTINPARAMGROUP_TO_LABEL.get(group_enum, str(group_enum))
        return self.GroupUnder

def build_vm_for_element(elem):
    vms = []
    type_elem = None
    if elem.GetTypeId() != ElementId.InvalidElementId:
        type_elem = elem.Document.GetElement(elem.GetTypeId())
    seen_param_names = set()
    def add_vm(p, insttype):
        if p.Definition.Name in seen_param_names:
            return
        seen_param_names.add(p.Definition.Name)
        definition = p.Definition
        group_under = definition.ParameterGroup.ToString().replace('PG_', '') if hasattr(definition, 'ParameterGroup') else "Other"
        discipline = "Common"
        vms.append(ParameterVM(p, group_under=group_under, discipline=discipline, insttype=insttype))
    for p in elem.Parameters:
        add_vm(p, "Instance")
    if type_elem:
        for tp in type_elem.Parameters:
            add_vm(tp, "Type")
    return vms

if not selected_elements:
    from System.Windows import MessageBox
    MessageBox.Show("Please select an element.", "No Selection")
    script.exit()

all_vm = build_vm_for_element(selected_elements[0])

# -- Load parameter manager XAML
main_xaml_path = os.path.join(os.path.dirname(__file__), 'ParameterManager.xaml')
with open(main_xaml_path, 'r') as f:
    xaml_str = f.read()
window = XamlReader.Load(XmlReader.Create(StringReader(xaml_str)))

dataGrid = window.FindName('dataGrid')
chkShowExisting = window.FindName('chkShowExisting')
chkAll = window.FindName('chkAll')
txtSearch = window.FindName('txtSearch')
chkApplyAllSimilar = window.FindName('chkApplyAllSimilar')
btnApply = window.FindName('btnApply')
btnCancel = window.FindName('btnCancel')
btnExport = window.FindName('btnExport')
btnAddParameter = window.FindName('btnAddParameter')

def refresh_grid():
    txt = txtSearch.Text if hasattr(txtSearch, "Text") else ""
    show_existing = chkShowExisting.IsChecked
    filtered = []
    for item in all_vm:
        if txt.lower() not in item.Name.lower(): continue
        if show_existing and not item.Editable: continue
        filtered.append(item)
    dataGrid.ItemsSource = List[object](filtered)
    chkAll.IsChecked = True
    for item in filtered: item.IsSelected = True
    try: dataGrid.Items.Refresh()
    except: pass

def select_all_changed(sender, args):
    state = chkAll.IsChecked
    items = dataGrid.ItemsSource
    if items is None: return
    for item in items: item.IsSelected = state
    try: dataGrid.Items.Refresh()
    except: pass

def search_changed(sender, args): refresh_grid()
def show_existing_changed(sender, args): refresh_grid()

def apply_clicked(sender, args):
    dataGrid.CommitEdit()
    dataGrid.CommitEdit()
    apply_to_all = chkApplyAllSimilar.IsChecked if hasattr(chkApplyAllSimilar, "IsChecked") else False
    from Autodesk.Revit.DB import Transaction, FilteredElementCollector
    t = Transaction(doc, "Update Parameters")
    t.Start()
    try:
        similar_elems = []
        if apply_to_all:
            base_elem = selected_elements[0]
            cat_id = base_elem.Category.Id if base_elem.Category else None
            base_type_id = base_elem.GetTypeId()
            collector = FilteredElementCollector(doc).OfCategoryId(cat_id).WhereElementIsNotElementType()
            for e in collector:
                if e.GetTypeId() == base_type_id:
                    similar_elems.append(e)
        for vm in dataGrid.ItemsSource:
            if not vm.Editable or not vm.IsSelected:
                continue
            param_name = vm.Name
            new_value = vm.Value
            if apply_to_all and similar_elems:
                for elem in similar_elems:
                    for p in elem.Parameters:
                        if p.Definition.Name == param_name and not p.IsReadOnly:
                            safe_set_value(p, new_value)
            else:
                p = vm.param
                safe_set_value(p, new_value)
        t.Commit()
    except Exception as e:
        t.RollBack()
        from System.Windows import MessageBox
        MessageBox.Show("Transaction rolled back:\n{}".format(e), "Error")
    window.Close()

def cancel_clicked(sender, args): window.Close()

def show_add_parameter_dialog():
    dlg_xaml_path = os.path.join(os.path.dirname(__file__), 'AddParameterDialog.xaml')
    with open(dlg_xaml_path, 'r') as f:
        xaml_str = f.read()
    dlg = XamlReader.Load(XmlReader.Create(StringReader(xaml_str)))
    txtParamName = dlg.FindName('txtParamName')
    cmbDiscipline = dlg.FindName('cmbDiscipline')
    cmbDataType = dlg.FindName('cmbDataType')
    cmbGroup = dlg.FindName('cmbGroup')
    cmbBindAs = dlg.FindName('cmbBindAs')
    btnOK = dlg.FindName('btnOK')
    btnCancel = dlg.FindName('btnCancel')
    from Autodesk.Revit.DB import DisciplineTypeId, SpecTypeId, BuiltInParameterGroup
    def safe_get_attr(obj, name):
        try: return getattr(obj, name)
        except AttributeError: return None
    disc_opts_long = [
        ("Common", "Common"),
        ("Electrical", "Electrical"),
        ("Energy", "Energy"),
        ("HVAC", "HVAC"),
        ("Infrastructure", "Infrastructure"),
        ("Piping", "Piping"),
        ("Structural", "Structural")
    ]
    disc_opts = []
    for label, attr in disc_opts_long:
        val = safe_get_attr(DisciplineTypeId, attr)
        if val is not None:
            disc_opts.append((label, val))
    cmbDiscipline.Items.Clear()
    for label, _ in disc_opts:
        cmbDiscipline.Items.Add(label)
    cmbDiscipline.SelectedIndex = 0
    def get_first_valid(*paths):
        for path in paths:
            obj = SpecTypeId
            try:
                parts = path.split('.')
                for part in parts:
                    obj = getattr(obj, part)
                return obj
            except AttributeError:
                continue
        return None
    type_options_master = [
        ("Angle", ["Angle"]),
        ("Area", ["Area"]),
        ("Cost", ["Currency"]),
        ("Count", ["Int.Integer", "Integer", "Number"]),
        ("Distance", ["Length"]),
        ("Force", ["Force"]),
        ("Image", ["String.Image", "Image"]),
        ("Length", ["Length"]),
        ("Material", ["Material"]),
        ("Mass Density", ["MassDensity"]),
        ("Number", ["Number"]),
        ("Percentage", ["Number.Percent", "Percent"]),
        ("Slope", ["Slope"]),
        ("Text", ["String.Text", "Text"]),
        ("Time", ["Time"]),
        ("URL", ["String.Url", "Url"]),
        ("Volume", ["Volume"]),
        ("Yes/No", ["Boolean.YesNo", "YesNo"]),
        ("Currency", ["Currency"]),
        ("Speed", ["Speed"]),
        ("Rotation Angle", ["Angle"]),
        ("Fill Pattern", ["FillPattern"]),
        ("Multiline Text", ["String.MultilineText", "MultilineText"]),
        ("Family Type", ["FamilyType"]),
    ]
    type_opts = []
    for label, paths in type_options_master:
        val = get_first_valid(*paths)
        if val is not None:
            type_opts.append((label, val))
    cmbDataType.Items.Clear()
    for label, _ in type_opts:
        cmbDataType.Items.Add(label)
    cmbDataType.SelectedIndex = 0
    group_names = [
        "Analysis Results", "Analytical Alignment", "Analytical Model", "Constraints", "Construction",
        "Data", "Dimensions", "Division Geometry", "Electrical", "Electrical - Circuiting",
        "Electrical - Lighting", "Electrical - Loads", "Electrical Engineering", "Energy Analysis",
        "Fire Protection", "Forces", "General", "Graphics", "Green Building Properties",
        "Identity Data", "IFC Parameters", "Layers", "Materials and Finishes", "Mechanical",
        "Mechanical - Flow", "Mechanical - Loads", "Model Properties", "Moments", "Other",
        "Overall Legend", "Phasing", "Photometrics", "Plumbing", "Primary End", "Rebar Set",
        "Releases / Member Forces", "Secondary End", "Segments and Fittings", "Set",
        "Slab Shape Edit", "Structural", "Structural Analysis", "Text", "Title Text", "Visibility"
    ]
    cmbGroup.Items.Clear()
    for label in group_names:
        cmbGroup.Items.Add(label)
    cmbGroup.SelectedIndex = 0
    bindas_opts = ["Instance", "Type"]
    cmbBindAs.Items.Clear()
    for label in bindas_opts:
        cmbBindAs.Items.Add(label)
    cmbBindAs.SelectedIndex = 0
    def ok_clicked(sender, args): dlg.DialogResult = True
    def cancel_clicked(sender, args): dlg.DialogResult = False
    btnOK.Click += ok_clicked
    btnCancel.Click += cancel_clicked
    result = dlg.ShowDialog()
    if result:
        pname = txtParamName.Text.strip()
        if not pname:
            MessageBox.Show("Parameter name cannot be empty!", "Error")
            return None
        disc = disc_opts[cmbDiscipline.SelectedIndex][1]
        dtype = type_opts[cmbDataType.SelectedIndex][1]
        group_label = cmbGroup.SelectedItem
        group_enum = GROUP_LABEL_TO_ENUM.get(group_label, BuiltInParameterGroup.INVALID)
        bindas = bindas_opts[cmbBindAs.SelectedIndex]
        return dict(name=pname, discipline=disc, datatype=dtype, group=group_enum, bindas=bindas)
    return None

def add_parameter_clicked(sender, args):
    paramdata = show_add_parameter_dialog()
    if paramdata is None: return
    if not paramdata['name']:
        MessageBox.Show("Parameter name cannot be empty!", "Error")
        return
    shared_param_file = r"C:\Temp\revit_shared_params.txt"
    folder = os.path.dirname(shared_param_file)
    if not os.path.exists(folder):
        try:
            os.makedirs(folder)
        except Exception as e:
            MessageBox.Show("Could not create folder:\n{}\n\nError: {}".format(folder, e), "Error")
            raise
    if not os.path.exists(shared_param_file):
        ask = MessageBox.Show("Shared parameter file not found:\n{}\n\nDo you want to create it now?",
                              "Shared Parameter File", MessageBoxButton.YesNo)
        if str(ask).lower() == "yes":
            try:
                with open(shared_param_file, 'w') as f:
                    f.write('')
            except Exception as e:
                MessageBox.Show("Failed to create shared parameter file:\n{}\n\nError: {}".format(shared_param_file, e), "Error")
                raise
        else:
            raise Exception("A shared parameter file is required. Operation canceled.")
    doc.Application.SharedParametersFilename = shared_param_file
    sp_file = doc.Application.OpenSharedParameterFile()
    if sp_file is None:
        MessageBox.Show("Shared parameter file not found. Please check file path.", "Error")
        return
    group_name = "Scripted"
    sp_group = None
    for g in sp_file.Groups:
        if g.Name == group_name:
            sp_group = g
            break
    if not sp_group:
        sp_group = sp_file.Groups.Create(group_name)
    defn = None
    for d in sp_group.Definitions:
        if d.Name == paramdata['name']:
            defn = d
            break
    if defn is None:
        opt = ExternalDefinitionCreationOptions(paramdata['name'], paramdata['datatype'])
        opt.Visible = True
        defn = sp_group.Definitions.Create(opt)
    cat_set = CategorySet()
    if selected_elements:
        cat = selected_elements[0].Category
        if cat: cat_set.Insert(cat)
    else:
        MessageBox.Show("No element selected.", "Error")
        return
    t = Transaction(doc, "Bind Shared Parameter")
    t.Start()
    try:
        bindings = doc.ParameterBindings
        if paramdata['bindas'] == "Instance":
            bindings.Insert(defn, InstanceBinding(cat_set), paramdata['group'])
        else:
            bindings.Insert(defn, TypeBinding(cat_set), paramdata['group'])
        t.Commit()
    except Exception as ex:
        t.RollBack()
        MessageBox.Show("Failed to bind parameter:\n" + str(ex), "Error")
        return
    global all_vm
    all_vm = build_vm_for_element(selected_elements[0])
    refresh_grid()

def export_clicked(sender, args):
    try:
        import System
        excel = System.Type.GetTypeFromProgID('Excel.Application')
        excel_app = System.Activator.CreateInstance(excel)
        wb = excel_app.Workbooks.Add()
        ws = wb.Worksheets[1]
        items = dataGrid.ItemsSource
        if not items:
            MessageBox.Show("No data to export.", "Export")
            return
        export_items = [item for item in items if hasattr(item, 'IsSelected') and item.IsSelected]
        if not export_items:
            MessageBox.Show("No parameters selected for export.", "Export")
            return
        element_name = None
        if selected_elements and hasattr(selected_elements[0], 'Name') and selected_elements[0].Name:
            element_name = selected_elements[0].Name
        else:
            try:
                element_name = "{}_{}".format(selected_elements[0].Category.Name, selected_elements[0].Id)
            except:
                element_name = "Element"
        ws.Name = element_name[:31]
        headers = [
            "Parameter Name",
            "Type of Parameter",
            "Discipline",
            "Type/CD",
            "Group Under",
            "Instance/Type",
            "Value"
        ]
        for col, h in enumerate(headers, 1):
            ws.Cells[1, col].Value2 = h
            header_range = ws.Range[ws.Cells(1, 1), ws.Cells(1, len(headers))]
            header_range.Font.Bold = True
        row = 2
        for item in export_items:
            ws.Cells[row, 1].Value2 = getattr(item, "Name", "")
            ws.Cells[row, 2].Value2 = getattr(item, "Type", "")
            ws.Cells[row, 3].Value2 = getattr(item, "Discipline", "")
            ws.Cells[row, 4].Value2 = getattr(item, "PType", "")
            ws.Cells[row, 5].Value2 = getattr(item, "GroupUnder", "")
            ws.Cells[row, 6].Value2 = getattr(item, "InstType", "")
            ws.Cells[row, 7].Value2 = getattr(item, "Value", "")
            row += 1
        ws.Columns.Autofit()
        excel_app.Visible = True
        MessageBox.Show('Exported to Excel (window is now open).', 'Export Complete')
    except Exception as e:
        MessageBox.Show('Export failed: {}'.format(e), 'Error')

def log_message(message):
    try:
        with open(r"C:\Temp\revit_script_log.txt", "a") as f:
            f.write(message + "\n")
    except:
        pass

def import_clicked(sender, args):
    log_message("Starting import_clicked function...")
    # Import Excel interop
    import clr
    clr.AddReference('Microsoft.Office.Interop.Excel')

    # Pick Excel file
    excel_path = forms.pick_file(file_ext='xlsx')
    if not excel_path:
        log_message("No Excel file selected. Exiting.")
        return
    log_message("Excel file selected: {}".format(excel_path))

    excel = System.Type.GetTypeFromProgID('Excel.Application')
    excel_app = System.Activator.CreateInstance(excel)
    excel_app.Visible = False
    wb = excel_app.Workbooks.Open(excel_path)

    # Determine sheet name
    if selected_elements and hasattr(selected_elements[0], 'Name') and selected_elements[0].Name:
        sheet_name = selected_elements[0].Name[:31]
    else:
        try:
            sheet_name = "{}_{}".format(selected_elements[0].Category.Name, selected_elements[0].Id)[:31]
        except:
            sheet_name = "Element"
    log_message("Using sheet name: {}".format(sheet_name))

    # Access worksheet
    try:
        ws = wb.Worksheets[sheet_name]
    except:
        MessageBox.Show("Sheet not found in Excel: " + sheet_name, "Import Error")
        wb.Close(False)
        excel_app.Quit()
        return

    # Read header columns to map indices
    headers = []
    for col in range(1, 8):
        h = ws.Cells[1, col].Value2
        if h:
            headers.append(h)
    log_message("Headers found: {}".format(", ".join(headers)))

    def get_col_i(header):
        for i, h in enumerate(headers):
            if header.lower() in h.lower():
                return i
        return None
    name_i     = get_col_i("Parameter Name")
    type_i     = get_col_i("Type of Parameter")
    disc_i     = get_col_i("Discipline")
    ptype_i    = get_col_i("Type/CD")
    group_i    = get_col_i("Group Under")
    insttype_i = get_col_i("Instance/Type")
    val_i      = get_col_i("Value")

    from Autodesk.Revit.DB import DisciplineTypeId, SpecTypeId, BuiltInParameterGroup, ExternalDefinitionCreationOptions, InstanceBinding, TypeBinding, CategorySet, StorageType, ElementId

    def safe_get_attr(obj, name, fallback=None):
        try:
            val = getattr(obj, name)
            return val
        except AttributeError:
            return fallback

    DISC_MAP = {
        "Common":       safe_get_attr(DisciplineTypeId, "Common"),
        "Electrical":   safe_get_attr(DisciplineTypeId, "Electrical"),
        "Energy":       safe_get_attr(DisciplineTypeId, "Energy"),
        "HVAC":         safe_get_attr(DisciplineTypeId, "HVAC"),
        "Infrastructure":safe_get_attr(DisciplineTypeId, "Infrastructure"),
        "Piping":       safe_get_attr(DisciplineTypeId, "Piping"),
        "Structural":   safe_get_attr(DisciplineTypeId, "Structural")
    }

    def get_spec_type(type_str):
        SPEC_PATHS = {
            "Text": ["String.Text", "Text"],
            "Yes/No": ["Boolean.YesNo", "YesNo"],
            "Number": ["Number"],
            "Angle": ["Angle"],
            "Area": ["Area"],
            "Cost": ["Currency"],
            "Count": ["Int.Integer", "Integer", "Number"],
            "Distance": ["Length"],
            "Force": ["Force"],
            "Image": ["String.Image", "Image"],
            "Length": ["Length"],
            "Material": ["Material"],
            "Mass Density": ["MassDensity"],
            "Percentage": ["Number.Percent", "Percent"],
            "Slope": ["Slope"],
            "Time": ["Time"],
            "URL": ["String.Url", "Url"],
            "Volume": ["Volume"],
            "Currency": ["Currency"],
            "Speed": ["Speed"],
            "Rotation Angle": ["Angle"],
            "Fill Pattern": ["FillPattern"],
            "Multiline Text": ["String.MultilineText", "MultilineText"],
            "Family Type": ["FamilyType"]
        }
        from Autodesk.Revit.DB import SpecTypeId
        if type_str in SPEC_PATHS:
            for path in SPEC_PATHS[type_str]:
                obj = SpecTypeId
                try:
                    for p in path.split('.'):
                        obj = getattr(obj, p)
                    return obj
                except:
                    continue
        return SpecTypeId.String.Text

    GROUP_MAP = {
        "Other": BuiltInParameterGroup.INVALID
    }

    el = selected_elements[0]
    update_count = 0
    add_count = 0
    t = Transaction(doc, "Import Parameters from Excel")
    t.Start()
    try:
        row = 2
        while ws.Cells[row, 1].Value2:
            param_name = ws.Cells[row, name_i + 1].Value2
            type_str   = ws.Cells[row, type_i + 1].Value2 if type_i is not None else ""
            disc_str   = ws.Cells[row, disc_i + 1].Value2 if disc_i is not None else ""
            group_str  = ws.Cells[row, group_i + 1].Value2 if group_i is not None else ""
            bindas_str = ws.Cells[row, insttype_i + 1].Value2 if insttype_i is not None else ""
            param_value= ws.Cells[row, val_i + 1].Value2

            p = el.LookupParameter(param_name)
            if p:
                # Do NOT overwrite already-set values
                already_set = False
                try:
                    if p.StorageType == StorageType.String:
                        v = p.AsString()
                        if v is not None and str(v).strip() != "":
                            already_set = True
                    elif p.StorageType == StorageType.Integer:
                        v = p.AsInteger()
                        if v is not None:
                            already_set = True
                    elif p.StorageType == StorageType.Double:
                        v = p.AsDouble()
                        if v is not None:
                            already_set = True
                    elif p.StorageType == StorageType.ElementId:
                        v = p.AsElementId()
                        if v is not None and v.IntegerValue != -1:
                            already_set = True
                except Exception:
                    already_set = False
                if already_set:
                    row += 1
                    continue
                # Only set if value is not already there
                try:
                    safe_set_value(p, param_value)
                    update_count += 1
                except Exception:
                    pass
            else:
                # Add/bind as needed
                discipline = DISC_MAP.get(disc_str, DisciplineTypeId.Common)
                spec_type = get_spec_type(type_str)
                group_under = GROUP_MAP.get(group_str, BuiltInParameterGroup.INVALID)
                bindas_instance = str(bindas_str).lower().startswith("inst")
                param_name = param_name.strip()
                shared_param_file = r"C:\Temp\revit_shared_params.txt"
                doc.Application.SharedParametersFilename = shared_param_file
                sp_file = doc.Application.OpenSharedParameterFile()
                if sp_file is not None:
                    group_name = "Scripted"
                    sp_group = None
                    for g in sp_file.Groups:
                        if g.Name == group_name:
                            sp_group = g
                            break
                    if not sp_group:
                        sp_group = sp_file.Groups.Create(group_name)
                    defn = None
                    for d in sp_group.Definitions:
                        if d.Name == param_name:
                            defn = d
                            break
                    if defn is None:
                        opt = ExternalDefinitionCreationOptions(param_name, spec_type)
                        opt.Visible = True
                        defn = sp_group.Definitions.Create(opt)
                    cat_set = CategorySet()
                    elem_cat = el.Category
                    cat_set.Insert(elem_cat)
                    bindings = doc.ParameterBindings
                    if bindas_instance:
                        bindings.Insert(defn, InstanceBinding(cat_set), group_under)
                    else:
                        bindings.Insert(defn, TypeBinding(cat_set), group_under)
                    p_new = el.LookupParameter(param_name)
                    if p_new:
                        try:
                            safe_set_value(p_new, param_value)
                        except Exception:
                            pass
                    add_count += 1
            row += 1
        t.Commit()
        log_message("Transaction committed successfully.")
    except Exception as ex:
        t.RollBack()
        log_message("Import failed: {}".format(str(ex)))
        MessageBox.Show("Import failed: {}".format(ex), "Import Error")
        wb.Close(False)
        excel_app.Quit()
        return

    # Close Excel
    wb.Close(False)
    excel_app.Quit()
    log_message("Excel application closed.")

    # Update the view model and refresh the grid
    log_message("Updating view model...")
    try:
        global all_vm
        all_vm = build_vm_for_element(selected_elements[0])
        log_message("View model updated successfully.")
    except Exception as ex:
        log_message("Error updating view model: {}".format(str(ex)))
        MessageBox.Show("Error updating view model: {}".format(str(ex)), "Error")
        return

    log_message("Refreshing grid...")
    try:
        refresh_grid()
        log_message("Grid refreshed successfully.")
    except Exception as ex:
        log_message("Error refreshing grid: {}".format(str(ex)))
        MessageBox.Show("Error refreshing grid: {}".format(str(ex)), "Error")
        return

    # Show success message
    MessageBox.Show(
        "Imported parameters.\nUpdated: {}\nAdded (new/bound): {}".format(update_count, add_count),
        "Import Success"
    )
    log_message("Import completed. Updated: {}, Added: {}".format(update_count, add_count))

def log_message(message):
    try:
        with open(r"C:\Temp\revit_script_log.txt", "a") as f:
            f.write(message + "\n")
    except:
        pass


def remove_parameter_clicked(sender, args):
    log_message("Starting remove_parameter_clicked function...")
    # Step 1: Gather selected VMs from the data grid
    vm_list = list(dataGrid.ItemsSource) if dataGrid.ItemsSource is not None else []
    selected_vms = [vm for vm in vm_list if
                    getattr(vm, "IsSelected", False) and getattr(vm, "Type", None) in ["Shared Parameter",
                                                                                       "Project Parameter"]]

    if not selected_vms:
        MessageBox.Show("No shared or project parameters selected to delete.", "No Parameters Selected")
        return

    # Step 2: Collect parameter names to delete
    names_to_delete = set(vm.Name for vm in selected_vms)
    log_message("Selected Param Names: " + ", ".join(names_to_delete))

    # Step 3: Check if parameters are used in schedules
    log_message("Checking for schedule usage...")
    try:
        schedules = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
        log_message("Found {} schedules to check.".format(len(schedules)))
        for i, sched in enumerate(schedules):
            try:
                if not hasattr(sched, "Name"):
                    log_message("Schedule at index {} has no Name attribute. Skipping.".format(i))
                    continue
                sched_name = sched.Name
                log_message("Checking schedule: {}".format(sched_name))
                sched_def = sched.Definition
                if sched_def is None:
                    log_message("Schedule {} has no definition. Skipping.".format(sched_name))
                    continue
                field_order = sched_def.GetFieldOrder()
                if not field_order:
                    log_message("Schedule {} has no fields. Skipping.".format(sched_name))
                    continue
                log_message("Schedule {} has {} fields.".format(sched_name, len(field_order)))
                for j, field in enumerate(field_order):
                    try:
                        field_param = sched_def.GetField(field)
                        if field_param is None:
                            log_message("Field {} in schedule {} is None. Skipping.".format(j, sched_name))
                            continue
                        if not hasattr(field_param, "ParameterName"):
                            log_message("Field {} in schedule {} has no ParameterName. Skipping.".format(j, sched_name))
                            continue
                        param_name = field_param.ParameterName
                        if not param_name:
                            log_message(
                                "Field {} in schedule {} has empty ParameterName. Skipping.".format(j, sched_name))
                            continue
                        if param_name in names_to_delete:
                            MessageBox.Show(
                                "Parameter {} is used in schedule {}. Cannot delete.".format(param_name, sched_name),
                                "Cannot Delete")
                            return
                    except Exception as ex:
                        log_message("Error checking field {} in schedule {}: {}".format(j, sched_name, str(ex)))
                        continue
            except Exception as ex:
                log_message("Error checking schedule at index {}: {}".format(i, str(ex)))
                continue
        log_message("Finished checking schedules. No dependencies found.")
    except Exception as ex:
        log_message("Error during schedule usage check: {}".format(str(ex)))
        MessageBox.Show("Schedule check failed. Proceeding with deletion, but there may be dependencies.", "Warning")

    # Step 4: Confirm deletion with user
    log_message("Prompting for user confirmation...")
    result = MessageBox.Show("Are you sure you want to delete these parameters?\n{}".format("\n".join(names_to_delete)),
                             "Confirm Deletion", MessageBoxButton.YesNo)
    log_message("Confirmation result: {}".format(str(result)))
    if str(result) != "Yes":
        MessageBox.Show("Deletion canceled by user.", "Info")
        return

    # Step 5: Verify shared parameter file accessibility
    log_message("Verifying shared parameter file...")
    shared_param_file = r"C:\Temp\revit_shared_params.txt"
    try:
        if not os.path.exists(shared_param_file):
            MessageBox.Show("Shared parameter file not found: {}".format(shared_param_file), "Error")
            return
        doc.Application.SharedParametersFilename = shared_param_file
        sp_file = doc.Application.OpenSharedParameterFile()
        if sp_file is None:
            MessageBox.Show("Failed to open shared parameter file: {}".format(shared_param_file), "Error")
            return
        log_message("Shared parameter file opened successfully.")
    except Exception as ex:
        log_message("Error accessing shared parameter file: {}".format(str(ex)))
        MessageBox.Show("Error accessing shared parameter file. Cannot proceed.", "Error")
        return

    # Step 6: Collect bindings to remove
    log_message("Collecting parameter bindings to remove...")
    binding_map = doc.ParameterBindings
    bindings_to_remove = []
    binding_count = 0
    it = binding_map.ForwardIterator()
    while it.MoveNext():
        binding_count += 1
        definition = it.Key
        if not hasattr(definition, "Name"):
            log_message("Binding {} has no Name attribute. Skipping.".format(binding_count))
            continue
        def_name = definition.Name
        log_message("Found binding: {}".format(def_name))
        if def_name in names_to_delete:
            bindings_to_remove.append((definition, def_name))
    log_message("Total bindings found: {}".format(binding_count))
    log_message("Bindings to remove: " + ", ".join(name for _, name in bindings_to_remove))

    # Step 7: Remove bindings one by one
    deleted = []
    for definition, def_name in bindings_to_remove:
        log_message("Starting transaction for removing binding: {}".format(def_name))
        t = Transaction(doc, "Delete Parameter Binding: {}".format(def_name))
        try:
            t.Start()
            log_message("Transaction started successfully for: {}".format(def_name))
            log_message("Removing binding for parameter: {}".format(def_name))
            if binding_map.Remove(definition):
                deleted.append(def_name)
                log_message("Successfully removed binding for: {}".format(def_name))
            else:
                log_message("Failed to remove binding for parameter: {}".format(def_name))
                MessageBox.Show("Failed to remove binding for parameter: {}".format(def_name), "Warning")
            log_message("Committing transaction for: {}".format(def_name))
            t.Commit()
            log_message("Transaction committed successfully for: {}".format(def_name))
        except Exception as ex:
            t.RollBack()
            log_message("Delete operation failed for {}: {}".format(def_name, str(ex)))
            MessageBox.Show("Delete operation failed for {}: {}".format(def_name, str(ex)), "Error")
            return
        # Small delay to allow Revit to stabilize
        try:
            Thread.Sleep(100)  # 100ms delay
            log_message("Delay completed after removing binding for: {}".format(def_name))
        except:
            log_message("Delay failed after removing binding for: {}".format(def_name))

    # Step 8: Finalize
    if deleted:
        log_message("Parameters deleted: " + ", ".join(deleted))
    else:
        log_message("No parameters were deleted.")
        MessageBox.Show("No parameters were deleted.", "Info")
        return

    # Step 9: Refresh the data grid
    log_message("Refreshing data grid...")
    try:
        global all_vm
        all_vm = build_vm_for_element(selected_elements[0])
        refresh_grid()
        log_message("Data grid refreshed successfully.")
        MessageBox.Show("Parameter deletion completed. Data grid refreshed.", "Success")
    except Exception as ex:
        log_message("Error refreshing data grid: {}".format(str(ex)))
        MessageBox.Show("Error refreshing data grid: {}".format(str(ex)), "Error")


chkAll.Click += select_all_changed
txtSearch.TextChanged += search_changed
chkShowExisting.Click += show_existing_changed
btnApply.Click += apply_clicked
btnCancel.Click += cancel_clicked
btnExport.Click += export_clicked
btnAddParameter.Click += add_parameter_clicked
btnImport = window.FindName('btnImport')
btnImport.Click += import_clicked
btnRemoveParameter = window.FindName("btnRemoveParameter")
btnRemoveParameter.Click += remove_parameter_clicked

refresh_grid()
window.ShowDialog()