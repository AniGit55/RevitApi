# -*- coding: utf-8 -*-
__title__ = "Title Block"
__doc__ = '''Extracts data from sheets based on a mapping file (Excel) and exports to Excel, or imports Legend sections from Excel back into Revit.

Mapping file format (Excel):
- Row 1: Header row (e.g., "Column Name", "Parameter Name") - SKIPPED
- Column A: Column Names (e.g., "Project ID", "Drawn By")
- Column B: Parameter Names (e.g., "Project ID", "Drawn By")

You will be prompted to select the mapping file via a dialog.
'''
__author__ = 'Anirudh Pachore'
import System
# Debug: Confirm script starts
print("Script started at: %s" % System.DateTime.Now)

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script, forms
from pyrevit.forms import alert
import os
import clr
import re

from System.Runtime.InteropServices import Marshal
from System.Windows.Forms import OpenFileDialog, DialogResult

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

# Debug: Confirm imports are successful
print("Imports completed successfully")

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()
output_folder = os.path.expanduser("~\\Documents")
excel_path = os.path.join(output_folder, "sheet_data_export.xlsx")

# Debug: Confirm output object is initialized
output.print_md("ℹ️ Output initialized. Script is running...")

# Default mapping if no file is selected
DEFAULT_MAPPING = {
    "Project ID": "Project ID",
    "Drawn By": "Drawn By",
    "Checked By": "Checked By",
    "Designed By": "Designed By",
    "Approved By": "Approved By",
    "Sheet Issue Date": "Sheet Issue Date",
    "Orig.": "Sheet No._Origin",
    "Phas.": "Sheet No._Project phase",
    "Stag.": "Sheet No._Stage",
    "Area": "Sheet No._Facility-Area",
    "Zone": "Sheet No._Floor-Zone-Street",
    "Doc Type": "Sheet No._Doc type",
    "Disc": "Sheet No._Discipline",
    "Internal Revision": "Internal Revision",
    "Sheet Revision": "Sheet No._Revision"
}

def select_excel_file():
    """
    Opens a file dialog for the user to select an Excel mapping file.
    Returns the file path or None if cancelled.
    """
    # Debug: Confirm we're calling the function
    output.print_md("ℹ️ Calling select_excel_file() function...")

    # Create and configure the file dialog
    dialog = OpenFileDialog()
    dialog.Title = "Select Mapping Excel File"
    dialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
    dialog.Multiselect = False
    dialog.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments)

    # Debug: Confirm dialog is configured
    output.print_md("ℹ️ File dialog configured. Showing dialog...")

    # Show dialog
    result = dialog.ShowDialog()

    # Debug: Confirm dialog result
    output.print_md("ℹ️ Dialog result: %s" % result)

    # If user selected a file, return the path
    if result == DialogResult.OK:
        file_path = dialog.FileName
        output.print_md("✅ Selected Excel file: `%s`" % file_path)
        return file_path
    else:
        output.print_md("⚠️ No file selected.")
        return None

def load_mapping_file():
    """
    Prompt user to select an Excel mapping file and read column names and parameter names.
    Returns a dictionary of {column_name: parameter_name}.
    """
    mapping = {}

    # Debug: Confirm we're entering the function
    output.print_md("ℹ️ Entering load_mapping_file() function...")

    # Use the select_excel_file function to get the file path
    mapping_file = select_excel_file()

    # If no file is selected, use default mapping
    if not mapping_file:
        output.print_md("⚠️ No mapping file selected. Using default mapping.")
        return DEFAULT_MAPPING.copy()

    # Read the Excel file
    excel_app = None
    workbook = None
    worksheet = None
    try:
        excel_app = Excel.ApplicationClass()
        excel_app.Visible = False
        workbook = excel_app.Workbooks.Open(mapping_file)
        worksheet = workbook.Worksheets[1]

        # Read column names (Column A) and parameter names (Column B), skipping the first row (header)
        row_count = worksheet.UsedRange.Rows.Count
        for row in range(2, row_count + 1):  # Start from row 2 to skip header
            col_name = worksheet.Cells[row, 1].Value2
            param_name = worksheet.Cells[row, 2].Value2
            if col_name and param_name:  # Skip rows where either is empty
                if not isinstance(col_name, str) or not isinstance(param_name, str):
                    output.print_md("⚠️ Row %d: Invalid data. Column Name and Parameter Name must be strings. Skipping." % row)
                    continue
                mapping[col_name] = param_name

        if not mapping:
            output.print_md("⚠️ No valid mappings found in the file. Using default mapping.")
            return DEFAULT_MAPPING.copy()

        output.print_md("✅ Loaded mapping file: `%s`" % mapping_file)
        output.print_md("Mapping: %s" % mapping)

    except Exception as e:
        output.print_md("⚠️ Error reading mapping file: `%s`. Error: %s. Using default mapping." % (mapping_file, str(e)))
        mapping = DEFAULT_MAPPING.copy()
    finally:
        if worksheet:
            Marshal.ReleaseComObject(worksheet)
        if workbook:
            workbook.Close(False)
            Marshal.ReleaseComObject(workbook)
        if excel_app:
            excel_app.Quit()
            Marshal.ReleaseComObject(excel_app)

    return mapping

def get_param_value(element, param_name, built_in_param=None):
    if not element:
        return ""
    if built_in_param:
        param = element.get_Parameter(built_in_param)
        if param and param.HasValue:
            return param.AsString() or param.AsValueString() or ""
    param = element.LookupParameter(param_name)
    if param and param.HasValue:
        if param.StorageType == StorageType.String:
            return param.AsString() or ""
        elif param.StorageType == StorageType.Integer:
            return str(param.AsInteger())
        elif param.StorageType == StorageType.Double:
            return str(param.AsDouble())
        elif param.StorageType == StorageType.ElementId:
            return str(param.AsElementId().IntegerValue)
    return ""

# Debug: Confirm we're about to load the mapping file
output.print_md("ℹ️ About to load mapping file...")

# Load the mapping file
dynamic_mapping = load_mapping_file()

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
    sheet_dict = {}
    for sheet in all_sheets:
        sheet_key = "%s - %s (ID: %s)" % (sheet.SheetNumber, sheet.Name, sheet.Id.IntegerValue)
        sheet_dict[sheet_key] = sheet

    # Create a list of display strings in the same order as all_sheets
    sheet_display = ["%s - %s (ID: %s)" % (sheet.SheetNumber, sheet.Name, sheet.Id.IntegerValue) for sheet in all_sheets]

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

    # Start Excel export process
    excel_app = None
    workbook = None
    worksheet = None
    try:
        excel_app = Excel.ApplicationClass()
        excel_app.Visible = False
        workbook = excel_app.Workbooks.Add()
        worksheet = workbook.Worksheets[1]

        # Fixed headers
        fixed_headers = [
            "Sheet Number", "Sheet Name", "ANVISNINGAR", "NOTES", "FÖRKLARINGAR", "Revision Descriptions"
        ]
        # Add dynamic headers from the mapping
        dynamic_headers = list(dynamic_mapping.keys())
        headers = fixed_headers + dynamic_headers

        # Write headers to Excel
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
            # Debug: Start Legend extraction for this sheet
            output.print_md("ℹ️ Extracting Legend sections for sheet '%s - %s'..." % (sheet.SheetNumber, sheet.Name))

            # Legend Sections
            anvisningar_text = ""
            notes_text = ""
            forklaringar_text = ""
            viewports = FilteredElementCollector(doc, sheet.Id).OfClass(Viewport).ToElements()
            if not viewports:
                output.print_md("⚠️ No viewports found on sheet '%s'. Skipping Legend sections." % sheet.SheetNumber)
            else:
                output.print_md("ℹ️ Found %d viewport(s) on sheet '%s'. Checking for Legend views..." % (len(viewports), sheet.SheetNumber))
                for vp in viewports:
                    view = doc.GetElement(vp.ViewId)
                    if not view:
                        output.print_md("⚠️ View not found for viewport ID %d on sheet '%s'. Skipping." % (vp.Id.IntegerValue, sheet.SheetNumber))
                        continue
                    output.print_md("ℹ️ Viewport ID %d: View Type = '%s', View Name = '%s'" % (vp.Id.IntegerValue, view.ViewType, view.Name))
                    # Look for any Legend view, regardless of name
                    if isinstance(view, View) and view.ViewType == ViewType.Legend:
                        output.print_md("✅ Found Legend view: '%s'. Extracting TextNotes..." % view.Name)
                        text_notes = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().OfClass(TextNote).ToElements()
                        if not text_notes:
                            output.print_md("⚠️ No TextNotes found in view '%s'. Legend sections will be empty." % view.Name)
                        else:
                            output.print_md("ℹ️ Found %d TextNote(s) in view '%s'. Processing..." % (len(text_notes), view.Name))
                            text_notes_sorted = sorted(text_notes, key=lambda x: (-x.Coord.Y, x.Coord.X))
                            # Debug: Print raw TextNote content
                            raw_texts = [tn.Text for tn in text_notes_sorted]
                            output.print_md("ℹ️ Raw TextNote content: %s" % raw_texts)

                            # Step 1: Split into sections based on \r\r\r
                            current_section = []
                            sections = []
                            for text in raw_texts:
                                if text.endswith("\r\r\r"):
                                    # End of a section
                                    current_section.append(text[:-3])  # Remove \r\r\r
                                    sections.append("\r".join(current_section))
                                    current_section = []
                                else:
                                    current_section.append(text)
                            # Add the last section if it exists
                            if current_section:
                                sections.append("\r".join(current_section))

                            output.print_md("ℹ️ Identified %d sections in the legend: %s" % (len(sections), sections))

                            # Step 2: Identify and assign sections
                            for section in sections:
                                lines = section.split("\r")
                                if not lines:
                                    continue
                                first_line = lines[0].strip()
                                # Assign section based on header
                                if first_line == "ANVISNINGAR":
                                    anvisningar_text = "\n".join(line.strip() for line in lines)
                                    output.print_md("✅ Found 'ANVISNINGAR' section: '%s'. Writing to Excel." % anvisningar_text)
                                elif "NOTES" in first_line:
                                    # Check if the section ends with "NOTES\r"
                                    if section.strip().endswith("NOTES") or section.strip().endswith("NOTES\r"):
                                        # The content before "NOTES\r" belongs to the NOTES section
                                        notes_lines = lines[:-1] if (lines[-1].strip() == "NOTES") else lines
                                        notes_text = "\n".join(line.strip() for line in notes_lines if line.strip())
                                        output.print_md("✅ Found 'NOTES' section: '%s'. Writing to Excel." % notes_text)
                                elif first_line == "FÖRKLARINGAR / LEGEND":
                                    forklaringar_text = "\n".join(line.strip() for line in lines)
                                    output.print_md("✅ Found 'FÖRKLARINGAR / LEGEND' section: '%s'. Writing to Excel." % forklaringar_text)

                        break
                else:
                    output.print_md("⚠️ No Legend view found on sheet '%s'. Legend sections will be empty." % sheet.SheetNumber)

            # Revisions (only for Revision Descriptions)
            rev_descs = []
            for rev_id in sheet.GetAllRevisionIds():
                rev = doc.GetElement(rev_id)
                rev_descs.append(rev.Description)

            # Fixed data
            fixed_data = [
                sheet.SheetNumber,
                sheet.Name,
                anvisningar_text,
                notes_text,
                forklaringar_text,
                "\n".join(rev_descs)
            ]

            # Dynamic data from mapping
            dynamic_data = []
            for col_name in dynamic_headers:
                param_name = dynamic_mapping[col_name]
                value = get_param_value(sheet, param_name)
                dynamic_data.append(value)

            # Combine fixed and dynamic data
            data_row = fixed_data + dynamic_data

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
        output.print_md("✅ **Sheet data exported to Excel (with proper wrapping & auto height):** `%s`" % excel_path)

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
        forms.alert("Excel file not found at: %s" % excel_path, exitscript=True)

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
                output.print_md("⚠️ Row %d: Sheet Number is empty. Skipping." % row)
                continue

            # Find the corresponding sheet in Revit
            if sheet_number not in sheet_dict:
                output.print_md("⚠️ Sheet with Sheet Number '%s' not found in Revit⚫️. Skipping." % sheet_number)
                continue
            sheet = sheet_dict[sheet_number]

            # Get the Legend sections from Excel
            anvisningar_text = worksheet.Cells[row, headers.index("ANVISNINGAR") + 1].Value2 or ""
            notes_text = worksheet.Cells[row, headers.index("NOTES") + 1].Value2 or ""
            forklaringar_text = worksheet.Cells[row, headers.index("FÖRKLARINGAR") + 1].Value2 or ""

            # Reconstruct the Legend content with original formatting
            sections = []
            if anvisningar_text:
                # Convert Excel line breaks (\n) back to \r
                anvisningar_text = anvisningar_text.replace("\n", "\r")
                sections.append(anvisningar_text + "\r\r\r")
            if notes_text:
                notes_text = notes_text.replace("\n", "\r")
                sections.append(notes_text + "\rNOTES\r")
            if forklaringar_text:
                forklaringar_text = forklaringar_text.replace("\n", "\r")
                sections.append(forklaringar_text)

            # Skip if no sections have content
            if not sections:
                output.print_md("ℹ️ Row %d: No Legend sections content for sheet '%s'. Skipping." % (row, sheet_number))
                continue

            # Find the first Legend view on the sheet
            viewports = FilteredElementCollector(doc, sheet.Id).OfClass(Viewport).ToElements()
            legend_view = None
            for vp in viewports:
                view = doc.GetElement(vp.ViewId)
                if isinstance(view, View) and view.ViewType == ViewType.Legend:
                    legend_view = view
                    break

            if not legend_view:
                output.print_md("⚡ No Legend view found on sheet '%s'. Skipping." % sheet_number)
                continue

            # Collect existing TextNotes elements
            text_notes = FilteredElementCollector(doc, legend_view.Id).WhereElementIsNotElementType().OfClass(TextNote).ToElements()
            text_notes_sorted = sorted(text_notes, key=lambda x: (-x.Coord.Y, x.Coord.X))

            # Get existing Legend content for comparison
            existing_content = "".join([tn.Text for tn in text_notes_sorted])
            new_content = "".join(sections)

            # Skip if the content hasn't changed
            if existing_content == new_content:
                output.print_md("ℹ️ Row %d: Legend content for sheet '%s' has not changed. Skipping." % (row, sheet_number))
                continue

            # Start a transaction to modify TextNotes
            t = Transaction(doc, "Update Legend on Sheet %s" % sheet_number)
            t.Start()
            try:
                # Store properties of existing TextNotes
                positions = [(tn.Coord.X, tn.Coord.Y) for tn in text_notes_sorted]
                widths = [tn.Width for tn in text_notes_sorted]
                alignments = [tn.HorizontalAlignment for tn in text_notes_sorted]
                # Get the TextNoteType from existing notes or default
                text_note_type_id = text_notes_sorted[0].TextNoteType.Id if text_notes_sorted else FilteredElementCollector(doc).OfClass(TextNoteType).FirstElement().Id

                # Delete existing TextNotes
                for tn in text_notes_sorted:
                    doc.Delete(tn.Id)

                # Recreate TextNotes for each section
                current_y = positions[0][1] if positions else 0
                current_x = positions[0][0] if positions else 0
                text_note_options = TextNoteOptions(text_note_type_id)
                text_note_options.HorizontalAlignment = alignments[0] if alignments else HorizontalTextAlignment.Left

                for i, section in enumerate(sections):
                    new_note = TextNote.Create(doc, legend_view.Id, XYZ(current_x, current_y - i * 0.5, 0), section, text_note_type_id)
                    new_note.Width = widths[0] if widths else 1.0  # Preserve original width or default to 1.0 feet

                t.Commit()
                output.print_md("✅ Updated Legend on sheet '%s'" % sheet_number)

            except Exception as e:
                t.RollBack()
                output.print_md("⚫️ Error updating Legend on sheet '%s': %s" % (sheet_number, str(e)))
        except Exception as e:
            output.print_md("⚠️ Error processing row %d: %s" % (row, str(e)))
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

    output.print_md("✅ **Finished importing Legend sections from Excel: `%s`**" % excel_path)