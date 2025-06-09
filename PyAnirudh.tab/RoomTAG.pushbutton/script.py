# pyRevit script to set room numbers using a predefined list (IronPython-compatible)

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB.Architecture import Room
from pyrevit import revit, script

# Sample list of room numbers to assign (must match the number of rooms selected)
room_numbers = ["101", "102", "103", "104", "105"]  # <-- Customize this list as needed

# Get document and selection
doc = revit.doc
uidoc = revit.uidoc

# Allow user to select rooms manually
selection_ids = uidoc.Selection.GetElementIds()
if not selection_ids:
    TaskDialog.Show("Room Number Setter", "Please select rooms before running the script.")
    script.exit()

# Convert selected elements to Room objects
rooms = []
for eid in selection_ids:
    elem = doc.GetElement(eid)
    if isinstance(elem, Room):
        rooms.append(elem)

if not rooms:
    TaskDialog.Show("Room Number Setter", "No valid room elements were selected.")
    script.exit()

# Ensure room count matches number list
if len(rooms) != len(room_numbers):
    msg = "Mismatch between selected rooms ({0}) and numbers in list ({1}).".format(len(rooms), len(room_numbers))
    TaskDialog.Show("Room Number Setter", msg)
    script.exit()

# Sort rooms by element Id to ensure stable ordering (optional)
rooms.sort(key=lambda r: r.Id.IntegerValue)

# Start transaction to update room numbers
t = Transaction(doc, "Set Room Numbers")
t.Start()

for i, room in enumerate(rooms):
    try:
        room.Number = room_numbers[i]
    except Exception as e:
        print("Failed to set number for Room ID {0}: {1}".format(room.Id, str(e)))

t.Commit()

TaskDialog.Show("Room Number Setter", "Updated {0} room numbers successfully.".format(len(rooms)))
