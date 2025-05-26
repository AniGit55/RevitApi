# -*- coding: utf-8 -*-
__title__ = "Rename Views"
__doc__ = """Version = 1.0
Date    = 31.07.2024
_____________________________________________________________________
Description:
Rename Views in Revit by using Find/Replace Logic.
_____________________________________________________________________
How-to:
-> Click on the button
-> Select Views
-> Define Renaming Rules
-> Rename Views
_____________________________________________________________________ """

#⬇️ IMPORTS
#------------------------------
# Regular + Autodesk
from Autodesk.Revit.DB import *

# pyRevit
from pyrevit import revit, forms

# .NET Imports (You often need List import)
import clr
from pyrevit.forms import alert

clr.AddReference("System")
from System.Collections.Generic import List


#📦 VARIABLES
#------------------------------
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application



#🎯 MAIN

#1️⃣ Select Views

#Get Views - Selected in Project Browser
selected_ids = uidoc.Selection.GetElementIds()
selected_elements = [doc.GetElement(eid) for eid in selected_ids]
selected_views = [elem for elem in selected_elements if issubclass(type(elem), View)]

# If None Selected - Prompt SelectViews from pyrevit.forms
if not selected_views:
    selected_views = forms.select_views()

# Ensure Views Selected
if not selected_views:
    forms.alert('No Views Selected. Try Again', exitscript=True)


# #2️⃣🅰️ Define Renaming Rules
# prefix  = 'pre-'
# find    = 'Level'
# replace = 'A'
# suffix  = '-suf'

#2️⃣🅱️ Define Renaming Rules in Dynamic Way (UI Form)
from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)
components = [Label('Prefix:'),  TextBox('prefix'),
              Label('Find:'),    TextBox('find'),
              Label('Replace:'), TextBox('replace'),
              Label('Suffix:'),  TextBox('suffix'),
              Separator(),       Button('Rename Views')]

# Show the form
form = FlexForm('Title', components)
form.show()

# Check if form was cancelled or closed
if not form.values:
    alert('Form was cancelled or closed. Please try again.', title='No Input', exitscript=True)

# Extract inputs
user_inputs = form.values #type: dict
prefix  = user_inputs.get('prefix', '')   # Default to empty string if not filled
find    = user_inputs.get('find', '')
replace = user_inputs.get('replace', '')
suffix  = user_inputs.get('suffix', '')

# If all fields are empty, exit the script
if not any([prefix, find, replace, suffix]):
    alert('All input fields are empty. Please enter at least one value.', title='No Rename Rules', exitscript=True)

# Start Transaction to make changes in project

t = Transaction(doc, 'A-Rename Views')
t.Start()

for view in selected_views:

    #3️⃣ Create new view line
    old_name= view.Name
    new_name= prefix + old_name.replace(find,replace) + suffix

    #4️⃣ Rename Views (Unique Name)
    for i in range(20):
        try:
            view.Name = new_name
            print('{} -> {}'.format(old_name, new_name))
            break
        except:
            new_name += '*'

t.Commit()

print ('Done')

