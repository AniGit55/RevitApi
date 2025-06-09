# -*- coding: utf-8 -*-
__title__ = "Title Block"
__doc__ = '''Extracts parameters, general notes, revision table, and drawing details from selected sheets, or imports changes to General Notes from Excel back into Revit.'''
__author__ = 'Anirudh Pachore'

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, forms
from pyrevit.forms import alert
import os
import clr
import re
import System
from System.Runtime.InteropServices import Marshal

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()
output_folder = os.path.expanduser("~\\Documents")
excel_path = os.path.join(output_folder, "sheet_data_export.xlsx")


def get_param_value(element, param_name, built_in_param=None):
    if not element:
        return ""
    if built_in_param:
        param = element.get_Parameter(built_in_param)
        if param and param.HasValue:
            return param.AsString() or param.AsValueString() or ""
    param = element.LookupParameter(param_name)
    if param and not param.IsReadOnly:
        if param.StorageType == StorageType.String:
            return param.AsString() or ""
        elif param.StorageType == StorageType.Integer:
            return str(param.AsInteger())
        elif param.StorageType == StorageType.Double:
            return str(param.AsDouble())
        elif param.StorageType == StorageType.ElementId:
            return str(param.AsElementId().IntegerValue)
    return ""


# Show dialog to choose between Export and Import
options = ["Export to Excel", "Import from Excel"]
selected_option = forms.SelectFromList.show(
    options,
    title="Choose Action",
    button_name="Proceed",
    multiselect=False
)

if not selected_option:
    forms.alert("No action selected. Operation cancelled.", exitscript=True)

if selected_option == "Export to Excel":
    # Collect all sheets in the document
    all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType().ToElements()
    all_sheets = sorted(all_sheets, key=lambda x: x.SheetNumber)

    if not all_sheets:
        forms.alert("No sheets found in the document.", exitscript=True)

    # Create a dictionary for faster sheet lookup, using ElementId to ensure uniqueness
    sheet_dict = {"{} - {} (ID: {})".format(sheet.SheetNumber, sheet.Name, sheet.Id.IntegerValue): sheet for sheet in
                  all_sheets}

    # Create a list of display strings in the same order as all_sheets
    sheet_display = ["{} - {} (ID: {})".format(sheet.SheetNumber, sheet.Name, sheet.Id.IntegerValue) for sheet in
                     all_sheets]

    # Show the dialog to select sheets
    selected_sheet_names = forms.SelectFromList.show(
        sheet_display,
        title="Select Sheets to Extract",
        button_name="Select",
        multiselect=True,
        check_all=True
    )

    if not selected_sheet_names:
        forms.alert("No sheets selected. Operation cancelled.", exitscript=True)

    # Map selected sheet names to sheet objects
    sheets = [sheet_dict[sheet_name] for sheet_name in selected_sheet_names]

    # Get the TextNotes for Project ID, Pre., Check, Appro., and Date from the first Titleblock family
    titleblock = None
    for sheet in sheets:
        titleblocks = [doc.GetElement(id) for id in sheet.GetDependentElements(ElementClassFilter(FamilyInstance)) if
                       isinstance(doc.GetElement(id), FamilyInstance)]
        titleblock = titleblocks[0] if titleblocks else None
        if titleblock and titleblock.Symbol and titleblock.Symbol.Family:
            break

    project_id_value = ""
    pre_value = ""
    check_value = ""
    appro_value = ""
    date_value = ""

    if titleblock and titleblock.Symbol and titleblock.Symbol.Family:
        family = titleblock.Symbol.Family
        family_doc = doc.EditFamily(family)
        if family_doc:
            # Extract Project ID
            project_id_element = family_doc.GetElement(ElementId(8193766))
            if project_id_element and isinstance(project_id_element, TextNote):
                project_id_value = project_id_element.Text or ""

            # Extract Pre.
            pre_element = family_doc.GetElement(ElementId(9999991))  # Placeholder ID
            if pre_element and isinstance(pre_element, TextNote):
                pre_value = pre_element.Text or ""

            # Extract Check
            check_element = family_doc.GetElement(ElementId(9999992))  # Placeholder ID
            if check_element and isinstance(check_element, TextNote):
                check_value = check_element.Text or ""

            # Extract Appro.
            appro_element = family_doc.GetElement(ElementId(9999993))  # Placeholder ID
            if appro_element and isinstance(appro_element, TextNote):
                appro_value = appro_element.Text or ""

            # Extract Date
            date_element = family_doc.GetElement(ElementId(9999994))  # Placeholder ID
            if date_element and isinstance(date_element, TextNote):
                date_value = date_element.Text or ""

            family_doc.Close(False)

    # Start Excel export process
    excel_app = None
    workbook = None
    worksheet = None
    try:
        excel_app = Excel.ApplicationClass()
        excel_app.Visible = False
        workbook = excel_app.Workbooks.Add()
        worksheet = workbook.Worksheets[1]

        # Headers
        headers = [
            "Sheet Number", "Sheet Name", "Sheet Issue Date", "Drawn By", "Checked By", "Designed By", "Approved By",
            "Project ID", "Orig.", "Phas.", "Stag.", "Area", "Zone", "Doc Type", "Disc", "Ser No"
        ]
        headers.extend(
            ["General Notes", "Internal Revision", "Sheet Revision", "Revision Descriptions", "Pre.", "Check", "Appro.",
             "Date"])
        for col, header in enumerate(headers, 1):
            cell = worksheet.Cells[1, col]
            cell.Value2 = header
            cell.Font.Bold = True
            cell.Interior.Color = 0xB7DEE8  # Light blue background
            cell.WrapText = True

        # Autofit columns based on header text
        worksheet.Columns.AutoFit()

        # Freeze the header row
        worksheet.Rows(2).Select()
        excel_app.ActiveWindow.FreezePanes = True

        # Set specific height for the header row
        worksheet.Rows[1].RowHeight = 30

        # Increase width of Sheet Name column (column 2) after initial autofit
        worksheet.Columns[2].ColumnWidth = 30

        # Process each selected sheet
        for row_idx, sheet in enumerate(sheets, start=2):
            titleblocks = [doc.GetElement(id) for id in sheet.GetDependentElements(ElementClassFilter(FamilyInstance))
                           if isinstance(doc.GetElement(id), FamilyInstance)]
            titleblock = titleblocks[0] if titleblocks else None


            def safe_param(element, name, built_in_param=None):
                return get_param_value(element, name, built_in_param)


            # General Notes
            general_notes_text = ""
            viewports = FilteredElementCollector(doc, sheet.Id).OfClass(Viewport).ToElements()
            for vp in viewports:
                view = doc.GetElement(vp.ViewId)
                if isinstance(view, View) and view.ViewType == ViewType.Legend and "GENERAL NOTES" in view.Name.upper():
                    text_notes = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().OfClass(
                        TextNote).ToElements()
                    text_notes_sorted = sorted(text_notes, key=lambda x: (-x.Coord.Y, x.Coord.X))
                    raw_notes = " ".join([tn.Text for tn in text_notes_sorted])
                    raw_notes = re.sub(r'\s+', ' ', raw_notes).strip()
                    numbered_items = re.findall(r'\d+\.\s?.*?(?=(?:\d+\.\s?)|$)', raw_notes)
                    general_notes_text = "\n".join(item.strip() for item in numbered_items)
                    break

            # Revisions (only for Revision Descriptions)
            rev_descs = []
            for rev_id in sheet.GetAllRevisionIds():
                rev = doc.GetElement(rev_id)
                rev_descs.append(rev.Description)

            # Extract parameters
            data_row = [
                sheet.SheetNumber,
                sheet.Name,
                get_param_value(sheet, "Sheet Issue Date"),
                get_param_value(sheet, "Drawn By"),
                get_param_value(sheet, "Checked By"),
                get_param_value(sheet, "Designed By"),
                get_param_value(sheet, "Approved By"),
                project_id_value,
                safe_param(sheet, "Sheet No._Origin"),
                safe_param(sheet, "Sheet No._Project phase"),
                safe_param(sheet, "Sheet No._Stage"),
                safe_param(sheet, "Sheet No._Facility-Area"),
                safe_param(sheet, "Sheet No._Floor-Zone-Street"),
                safe_param(sheet, "Sheet No._Doc type"),
                safe_param(sheet, "Sheet No._Discipline"),
                sheet.SheetNumber,  # Ser No (same as Sheet Number)
                general_notes_text,
                safe_param(sheet, "Internal Revision"),
                safe_param(sheet, "Sheet No._Revision"),
                "\n".join(rev_descs),
                pre_value,
                check_value,
                appro_value,
                date_value
            ]

            # Write data to Excel, ensuring Sheet Number is treated as text
            for col_idx, value in enumerate(data_row, 1):
                cell = worksheet.Cells[row_idx, col_idx]
                if col_idx == 1:  # Sheet Number column
                    cell.Value2 = str(value)
                    cell.NumberFormat = "@"  # Set format to text
                else:
                    cell.Value2 = value

        # Apply formatting to the entire used range
        used_range = worksheet.UsedRange
        used_range.WrapText = True
        used_range.VerticalAlignment = -4160  # xlVAlignTop
        used_range.HorizontalAlignment = -4131  # xlLeft

        # Autofit columns and rows (except header row)
        worksheet.Columns.AutoFit()
        worksheet.Columns[2].ColumnWidth = 30  # Ensure Sheet Name column width
        worksheet.Rows.AutoFit()

        # Save and close
        workbook.SaveAs(excel_path)
        output.print_md(
            "✅ **Sheet data exported to Excel (with proper wrapping & auto height):** `{}`".format(excel_path))

    except Exception as e:
        TaskDialog.Show("Error", "An error occurred during export:\n" + str(e))
    finally:
        if worksheet:
            Marshal.ReleaseComObject(worksheet)
        if workbook:
            workbook.Close(False)
            Marshal.ReleaseComObject(workbook)
        if excel_app:
            excel_app.Quit()
            Marshal.ReleaseComObject(excel_app)

else:
    # Import from Excel
    if not os.path.exists(excel_path):
        forms.alert("Excel file not found at: {}".format(excel_path), exitscript=True)

    # Collect all sheets in Revit
    all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType().ToElements()
    sheet_dict = {sheet.SheetNumber: sheet for sheet in all_sheets}

    # Start Excel
    excel_app = Excel.ApplicationClass()
    excel_app.Visible = False
    workbook = excel_app.Workbooks.Open(excel_path)
    worksheet = workbook.Worksheets[1]

    # Get headers
    headers = [worksheet.Cells[1, col].Value2 for col in range(1, worksheet.UsedRange.Columns.Count + 1)]

    # Process each row (skip header row)
    for row in range(2, worksheet.UsedRange.Rows.Count + 1):
        try:
            # Get the Sheet Number and normalize it
            sheet_number = str(worksheet.Cells[row, headers.index("Sheet Number") + 1].Value2)
            if sheet_number.endswith(".0"):
                sheet_number = sheet_number[:-2]  # Remove ".0"

            # Skip if Sheet Number is empty
            if not sheet_number:
                output.print_md("⚠️ Row {}: Sheet Number is empty. Skipping.".format(row))
                continue

            # Find the corresponding sheet in Revit
            if sheet_number not in sheet_dict:
                output.print_md("⚠️ Sheet with Sheet Number '{}' not found in Revit⚫️. Skipping.".format(sheet_number))
                continue
            sheet = sheet_dict[sheet_number]

            # Get the General Notes from Excel
            general_notes_text = worksheet.Cells[row, headers.index("General Notes") + 1].Value2 or ""

            # Reformat the General Notes to ensure proper line breaks
            general_notes_text = re.sub(r'\s+', ' ', general_notes_text).strip()
            if not general_notes_text:
                output.print_md(
                    "ℹ️ Row {}: General Notes for sheet '{}' is empty or contains only whitespace. Skipping.".format(
                        row, sheet_number))
                continue

            numbered_items = re.findall(r'\d+\.\s?.*?(?=(?:\d+\.\s?)|$)', general_notes_text, re.DOTALL)
            formatted_text = "\n".join(item.strip() for item in numbered_items)
            if not formatted_text:
                output.print_md(
                    "ℹ️ Row {}: No numbered items found in General Notes for sheet '{}'. Skipping.".format(row,
                                                                                                           sheet_number))
                continue

            # Find the general notes legend view
            viewports = FilteredElementCollector(doc, sheet.Id).OfClass(Viewport).ToElements()
            general_notes_view = None
            for vp in viewports:
                view = doc.GetElement(vp.ViewId)
                if isinstance(view, View) and view.ViewType == ViewType.Legend and "GENERAL NOTES" in view.Name.upper():
                    general_notes_view = view
                    break

            if not general_notes_view:
                output.print_md("⚡ General Notes legend view not found on sheet '{}'. Skipping.".format(sheet_number))
                continue

            # Collect existing TextNotes elements
            text_notes = FilteredElementCollector(doc, general_notes_view.Id).WhereElementIsNotElementType().OfClass(
                TextNote).ToElements()
            text_notes_sorted = sorted(text_notes, key=lambda x: (-x.Coord.Y, x.Coord.X))

            # Get existing General Notes content
            existing_notes = ""
            if text_notes_sorted:
                raw_existing = " ".join([tn.Text for tn in text_notes_sorted])
                raw_existing = re.sub(r'\s+', ' ', raw_existing).strip()
                existing_numbered = re.findall(r'\d+\.\s?.*?(?=(?:\d+\.\s?)|$)', raw_existing, re.DOTALL)
                existing_notes = "\n".join(item.strip() for item in existing_numbered)

            # Skip if the General Notes haven't changed
            if formatted_text == existing_notes:
                output.print_md(
                    "ℹ️ Row {}: General Notes for sheet '{}' have not changed. Skipping.".format(row, sheet_number))
                continue

            # Start a transaction to modify TextNotes
            t = Transaction(doc, "Update General Notes on Sheet {}".format(sheet_number))
            t.Start()
            try:
                # Store properties of existing TextNotes
                positions = [(tn.Coord.X, tn.Coord.Y) for tn in text_notes_sorted]
                widths = [tn.Width for tn in text_notes_sorted]
                alignments = [tn.HorizontalAlignment for tn in text_notes_sorted]
                # Get the TextNoteType from existing notes or default
                text_note_type_id = text_notes_sorted[
                    0].TextNoteType.Id if text_notes_sorted else FilteredElementCollector(doc).OfClass(
                    TextNoteType).FirstElement().Id

                # Delete existing TextNotes (we only need one)
                for tn in text_notes_sorted:
                    doc.Delete(tn.Id)

                # Create a single TextNote with the formatted General Notes
                x = positions[0][0] if positions else 0
                y = positions[0][1] if positions else 0
                text_note_options = TextNoteOptions(text_note_type_id)
                text_note_options.HorizontalAlignment = alignments[0] if alignments else HorizontalTextAlignment.Left
                new_note = TextNote.Create(doc, general_notes_view.Id, XYZ(x, y, 0), formatted_text, text_note_type_id)
                new_note.Width = widths[0] if widths else 1.0  # Preserve original width or default to 1.0 feet

                t.Commit()
                output.print_md("✅ Updated General Notes on sheet '{}'".format(sheet_number))

            except Exception as e:
                t.RollBack()
                output.print_md("⚫️ Error updating General Notes on sheet '{}': {}".format(sheet_number, str(e)))
        except Exception as e:
            output.print_md("⚠️ Error processing row {}: {}".format(row, str(e)))
            continue

    # Clean up Excel
    try:
        workbook.Close(False)
        excel_app.Quit()
    finally:
        if 'worksheet' in locals():
            Marshal.ReleaseComObject(worksheet)
        if 'workbook' in locals():
            Marshal.ReleaseComObject(workbook)
        if 'excel_app' in locals():
            Marshal.ReleaseComObject(excel_app)

    output.print_md("✅ **Finished importing General Notes from Excel: `{}`**".format(excel_path))