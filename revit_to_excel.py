"""
MIT License

Copyright (c) 2018 C. Claus 

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""


import clr
import sys
from random import random
from random import randint
from math import *
import collections
from collections import OrderedDict
from collections import defaultdict
#import xlsxwriter
#import xlwt




clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("RevitAPI") 
clr.AddReference("RevitAPIUI")
clr.AddReference("Microsoft.Office.Interop.Excel")


import Autodesk
from Autodesk.Revit.DB import * 
from Autodesk.Revit.UI import *


import System
from System.Collections.Generic import *
from System.Collections import *
from System import *
from System.Drawing import Point, Icon, Color
from System.Drawing import Color, Font, FontStyle, Point
from System.Windows.Forms import (Application, BorderStyle, FormBorderStyle, Button, CheckBox, Form, Label, Panel, ToolTip, RadioButton, CheckedListBox, CheckState)
from System.Drawing import Icon

from Microsoft.Office.Interop import Excel
import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()

"""
import System
from System import *
from System.Collections.Generic import *
from System.Collections import *
from System.IO import Directory, Path

from System.Runtime.InteropServices import Marshal
from Microsoft.Office.Interop import Excel
import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()
"""

to_metric_surface = 10.76391042
to_metric_volume = 35.31466671
to_metric_length = 0.03937007874

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView


#__window__.Hide()
#__window__.Close()
######################################################################################################
###################################Get Built in Categories############################################
######################################################################################################
categories_list = [    
                                'OST_Walls',
                                'OST_Floors',
                                'OST_Ceilings',
                                'OST_StructuralFraming',
                                'OST_StructuralFoundation',
                                'OST_StructuralColumns',
                                'OST_Roofs', 
                                'OST_Ramps',
                                'OST_Stairs',
                                'OST_Site',
                                'OST_DuctTerminal', 
                                'OST_Casework',
                                'OST_CableTray',
                                'OST_Conduit', 
                                'OST_ElectricalFixtures',
                                'OST_Furniture',
                                'OST_GenericModel',
                                'OST_Gutter',
                                'OST_MechanicalEquipment',
                                'OST_PlumbingFixtures',
                                'OST_Doors',
                                'OST_Windows',
                                'OST_CurtainWallMullions',
                                'OST_CurtainWallPanels',
                                'OST_RailingBalusterRail',
                                'OST_RailingBalusterRailCut',
                                'OST_RailingHandRail',
                                'OST_RailingHandRailAboveCut',
                                'OST_RailingRailPathExtensionLines',
                                'OST_RalingRailPathLines',
                                'OST_Railings',
                                'OST_RailingSystem',
                                'OST_RailingsystemBaluser',
                                'OST_CurtainWallPanels',
                                'OST_CurtainWallMullions',
                                'OST_CurtaSystem',
                                ]  
                                
ost_topography = ['OST_Topography']  


                                
ost_materials = [ 'OST_Materials']    



        
######################################################################################################
#####################################Get Revit Categories#############################################
######################################################################################################         
builtin_categories = System.Enum.GetValues(BuiltInCategory)

t1, t2 = [],[]
categories_assembly_code_list = []
topography = []

for category, builtin_category in [(category,builtin_category) for category in categories_list for builtin_category in builtin_categories]:
    if category == builtin_category.ToString():
        t1.append(FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(builtin_category).ToElements())
        t2.append(builtin_category)
        
        
elements, categories = [], []
categories_list = []


for i in range(len(t1)):
    if t1[i]:
        elements.append(t1[i])
        categories.append(t2[i])
        
for i in elements:
    for j in i:
        categories_list.append(j)
        
        
   
        
        
######################################################################################################
##########################################Get Revit Materials#########################################
######################################################################################################       
id_materials_dict = collections.defaultdict(list)

materials_list = ['OST_Materials']      
        
        
        
material_categories = System.Enum.GetValues(BuiltInCategory)

t3, t4 = [],[]
categories_assembly_code_list = []

for material_category, builtin_material_category in [(material_category,builtin_material_category) for material_category in materials_list for builtin_material_category in material_categories]:
    if material_category == builtin_material_category.ToString():
        t3.append(FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(builtin_material_category).ToElements())
        t4.append(builtin_material_category)
        
material_elements, categories = [], []
material_elements_list = []

for i in range(len(t3)):
    if t3[i]:
        material_elements.append(t3[i])
        categories.append(t4[i])
        
for i in material_elements:
    for j in i:
        material_elements_list.append(j)       

######################################################################################################
######################################### Metric Conversions #########################################
###################################################################################################### 

to_metric_surface = 10.76391042
to_metric_volume = 35.31466671


					
######################################################################################################
######################################### Get Element Data ###########################################
###################################################################################################### 

for i in categories_list:

	if doc.GetElement(i.GetTypeId()) is not None:
		element_type = doc.GetElement(i.GetTypeId())

	
def get_id(i):
	return i.Id
 	
def get_category(i):
	return i.Category.Name
	
def get_name(i):
	return i.Name
	
def get_assembly_code(type_parameter):
	if type_parameter.LookupParameter('Assembly Code'):
		return type_parameter.LookupParameter('Assembly Code').AsString()
		
def get_level(i):
	level_list = [  'Level',
					'Base Level',
					'Reference Level',
					'Base Constraint',
					'Host',
					'Work Plane'
					]
	
	levels_list = []
	
	for level_names in level_list:
		if i.LookupParameter(str(level_names)):
			return i.LookupParameter(str(level_names)).AsValueString()
		

	
def get_width(type_parameter, i):
	if type_parameter.LookupParameter('Width'):
		return type_parameter.LookupParameter('Width').AsValueString()
		
	
	if i.LookupParameter('Width'):
		return i.LookupParameter('Width').AsValueString()
	
def get_length(i):
	if i.LookupParameter('Length'):
		return i.LookupParameter('Length').AsValueString()
	
def get_unconnected_height(i):
	if i.LookupParameter('Unconnected Height'):
		return i.LookupParameter('Unconnected Height').AsValueString()
		
def get_thickness(i):
	if i.LookupParameter('Thickness'):
		return i.LookupParameter('Thickness').AsValueString()
	
def get_area(i):
	if i.LookupParameter('Area'):
		return round(i.LookupParameter('Area').AsDouble()/to_metric_surface,2)
		
def get_volume(i):
	if i.LookupParameter('Volume'):
		return round(i.LookupParameter('Volume').AsDouble()/to_metric_volume,2)
	
def get_perimeter(i):
	if i.LookupParameter('Perimeter'):
		return i.LookupParameter('Perimeter').AsValueString()
	
def get_thickness(i):
	if i.LookupParameter('Thickness'):
		return i.LookupParameter('Thickness').AsValueString()
	
def get_cut_length(i):
	if i.LookupParameter('Cut Length'):
		return i.LookupParameter('Cut Length').AsValueString()
		
def get_phase_created(i):
	if i.LookupParameter('Phase Created'):
		if i.LookupParameter('Phase Created').AsValueString() is not None:
			#print i.LookupParameter('Phase Created').AsValueString()
			return i.LookupParameter('Phase Created').AsValueString()
		
def get_phase_demolished(i):
	if i.LookupParameter('Phase Demolished'):
		if i.LookupParameter('Phase Created').AsValueString() is not None:
			#print i.LookupParameter('Phase Created').AsValueString()
			return i.LookupParameter('Phase Demolished').AsValueString()
		
		
	
data_list=[]
data_dict={}

width_dict={}

for items in categories_list:

	if doc.GetElement(items.GetTypeId()) is not None:

		element_type = doc.GetElement(items.GetTypeId())
	
	
	data_dict[get_id(i=items)]=	[	get_category(i=items),
									get_name(i=items),
									get_assembly_code(type_parameter=element_type),
									get_level(i=items),
									get_width(type_parameter=element_type, i=items),
									get_length(i=items),
									get_unconnected_height(i=items),
									get_thickness(i=items),
									get_area(i=items),
									get_volume(i=items),
									get_perimeter(i=items),
									get_thickness(i=items),
									#get_cut_length(i=items),
									get_phase_created(i=items),
									get_phase_demolished(i=items)]

	
    	

						
	
header = ["BatId","Revit Category","Type","Assembly Code","Level","Width","Length","Unconnected Height","Area","Volume","Perimeter","Thickness","Phase Created","Phase Demolished"]

bat_id_list=[]
revit_category_list=[]
revit_type_list=[]
assembly_code_list=[]
level_list=[]

width_list=[]
length_list=[]
unconnected_height_list=[]
area_list=[]
volume_list=[]

perimeter_list=[]
thickness_list=[]
cut_length_list=[]

phase_created_list=[]
phase_demolished_list=[]

for k, v in data_dict.iteritems():
	
	#A
	bat_id_list.append(k)
	
	#B
	revit_category_list.append(v[0])
	
	#C
	revit_type_list.append(v[1])
	
	#D
	assembly_code_list.append(v[2])
	
	#E
	level_list.append(v[3])
	
	#F
	width_list.append(v[4])
	
	#G
	length_list.append(v[5])
	
	#H
	unconnected_height_list.append(v[6])
	
	#I
	area_list.append(v[7])
	
	#J
	volume_list.append(v[8])
	
	#K
	perimeter_list.append(v[9])
	
	#L
	thickness_list.append(v[10])
	
	#M
	#cut_length_list.append(v[11])

	#N
	phase_created_list.append(v[11])
	
	#O
	phase_demolished_list.append(v[12])
	
	

	

excel = Excel.ApplicationClass()
excel.Visible = True 

#workbook = excel.Workbooks.Open('U:\\09. Python scripts\\00_meetstaat_generator\\meetstaat.xlsx')
workbook = excel.Workbooks.Add()

worksheet_columns = workbook.Worksheets.Add()
worksheet_columns.Name = 'kolommen'
worksheet_beams = workbook.Worksheets.Add()
worksheet_beams.Name = 'liggers'
worksheet_ceilings = workbook.Worksheets.Add()
worksheet_ceilings.Name = 'plafonds'
worksheet_floors = workbook.Worksheets.Add()
worksheet_floors.Name = 'vloeren'
worksheet_walls = workbook.Worksheets.Add()
worksheet_walls.Name = 'wanden'
worksheet_data = workbook.Worksheets.Add()
worksheet_data.Name = 'data'


#####################################################
###################### Header #######################
#####################################################
i=0

xlrange_header = worksheet_data.Range["A1:O1"]
a = Array.CreateInstance(object, 2, len(header))

while i < len(header):
	a[0,i] = header[i]
	i += 1
	
xlrange_header.Value2 = a





#####################################################
####################### BatID #######################
#####################################################

i = 0 

xlrange = worksheet_data.Range["A2:A" + str(len(bat_id_list)+1)]
a = Array.CreateInstance(object,len(bat_id_list), 3) 

while i < len(bat_id_list):
    a[i,0] = bat_id_list[i]
    i += 1

xlrange.Value2 = a


#####################################################
################## Revit Category ###################
#####################################################

i = 0 

xlrange_revit_category = worksheet_data.Range["B2:B" + str(len(revit_category_list)+1)]
a = Array.CreateInstance(object,len(revit_category_list), 3) 

while i < len(revit_category_list):
    a[i,0] = revit_category_list[i]
    i += 1

xlrange_revit_category.Value2 = a


#####################################################
################## Revit Type #######################
#####################################################

i = 0 

xlrange_revit_type = worksheet_data.Range["C2:C" + str(len(revit_type_list)+1)]
a = Array.CreateInstance(object,len(revit_type_list), 3) 

while i < len(revit_type_list):
    a[i,0] = revit_type_list[i]
    i += 1

xlrange_revit_type.Value2 = a



#####################################################
################## Assembly Code ####################
#####################################################

i = 0 

xlrange_assembly_code = worksheet_data.Range["D2:D" + str(len(assembly_code_list)+1)]
a = Array.CreateInstance(object,len(assembly_code_list), 3) 

while i < len(assembly_code_list):
    a[i,0] = assembly_code_list[i]
    i += 1

xlrange_assembly_code.Value2 = a



#####################################################
##################### Level #########################
#####################################################

i = 0 

xlrange_level = worksheet_data.Range["E2:E" + str(len(level_list)+1)]
a = Array.CreateInstance(object,len(level_list), 3) 

while i < len(level_list):
    a[i,0] = level_list[i]
    i += 1

xlrange_level.Value2 = a


#####################################################
##################### Width #########################
#####################################################

i = 0 

xlrange_width = worksheet_data.Range["F2:F" + str(len(width_list)+1)]
a = Array.CreateInstance(object,len(width_list), 3) 

while i < len(width_list):
    a[i,0] = width_list[i]
    i += 1

xlrange_width.Value2 = a


#####################################################
##################### Length ########################
#####################################################

i = 0 

xlrange_length = worksheet_data.Range["G2:G" + str(len(length_list)+1)]
a = Array.CreateInstance(object,len(length_list), 3) 

while i < len(length_list):
    a[i,0] = length_list[i]
    i += 1

xlrange_length.Value2 = a


#####################################################
############## Unconnected Height ###################
#####################################################

i = 0 

xlrange_unconnected_height = worksheet_data.Range["H2:H" + str(len(unconnected_height_list)+1)]
a = Array.CreateInstance(object,len(unconnected_height_list), 3) 

while i < len(unconnected_height_list):
    a[i,0] = unconnected_height_list[i]
    i += 1

xlrange_unconnected_height.Value2 = a



#####################################################
###################### Area #########################
#####################################################

i = 0 

xlrange_area = worksheet_data.Range["I2:I" + str(len(area_list)+1)]
a = Array.CreateInstance(object,len(area_list), 3) 

while i < len(area_list):
    a[i,0] = area_list[i]
    i += 1

xlrange_area.Value2 = a


#####################################################
#################### Volume #########################
#####################################################

i = 0 

xlrange_volume = worksheet_data.Range["J2:J" + str(len(volume_list)+1)]
a = Array.CreateInstance(object,len(volume_list), 3) 

while i < len(volume_list):
    a[i,0] = volume_list[i]
    i += 1

xlrange_volume.Value2 = a


#####################################################
################### Thickness #######################
#####################################################

i = 0 

xlrange_thickness = worksheet_data.Range["K2:K" + str(len(thickness_list)+1)]
a = Array.CreateInstance(object,len(thickness_list), 3) 

while i < len(thickness_list):
    a[i,0] = thickness_list[i]
    i += 1

xlrange_thickness.Value2 = a




#####################################################
################### Perimeter #######################
#####################################################

i = 0 

xlrange_perimeter = worksheet_data.Range["L2:L" + str(len(perimeter_list)+1)]
a = Array.CreateInstance(object,len(perimeter_list), 3) 

while i < len(perimeter_list):
    a[i,0] = perimeter_list[i]
    i += 1

xlrange_perimeter.Value2 = a

"""

#####################################################
################### Cut Length ######################
#####################################################

i = 0 

xlrange_cut_length = worksheet_data.Range["M2:M" + str(len(cut_length_list)+1)]
a = Array.CreateInstance(object,len(cut_length_list), 3) 

while i < len(cut_length_list):
    a[i,0] = cut_length_list[i]
    i += 1

xlrange_cut_length.Value2 = a

"""


#####################################################
################ Phase Created ######################
#####################################################

i = 0 

xlrange_phase_created = worksheet_data.Range["M2:M" + str(len(phase_created_list)+1)]
a = Array.CreateInstance(object,len(phase_created_list), 3) 

while i < len(phase_created_list):
    a[i,0] = phase_created_list[i]
    i += 1

xlrange_phase_created.Value2 = a




#####################################################
################ Phase Demolished ###################
#####################################################


i = 0 

xlrange_phase_demolished = worksheet_data.Range["M2:M" + str(len(phase_demolished_list)+1)]
a = Array.CreateInstance(object,len(phase_demolished_list), 3) 

while i < len(phase_demolished_list):
    a[i,0] = phase_demolished_list[i]
    i += 1
	
xlrange_phase_demolished.Value2 = a
























