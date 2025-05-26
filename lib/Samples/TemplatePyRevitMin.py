# -*- coding: utf-8 -*-

__title__ = "EF Template.min"

__context__ = 'selection'
# Make your button available only when certain categories are selected. Or Revit/View Types.

__doc__ = """Version = 1.0
Description:
This is a template file for pyRevit Scripts.
_____________________________________________________________________
How-to:
-> Click on the button """


#⬇️ IMPORTS
#------------------------------
# Regular + Autodesk
from Autodesk.Revit.DB import *

# pyRevit
from pyrevit import revit, forms

# .NET Imports (You often need List import)
import clr
clr.AddReference("System")
from System.Collections.Generic import List


#📦 VARIABLES
#------------------------------
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application


#🧬 FUNCTIONS
#------------------------------


#🎯 MAIN
#------------------------------
# START CODE HERE
