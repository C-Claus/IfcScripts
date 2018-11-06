#!/usr/bin/python
# -*- coding: <encoding name> -*-

import ifcopenshell
import time
import xlsxwriter
import os
import collections
import sys
import shutil
import datetime


start_time = time.time()


filename = None


chb_red = "#E83426"
chb_grey = "#EDEDED"
white = "#FFFFFF"
black = "#000000"


def ifcfile(filename):
    
    print 'IFC is being read'
    
    ifc_file = ifcopenshell.open(filename)
    
    project = ifc_file.by_type('IfcProject')
    
    products = ifc_file.by_type('IfcProduct')
    
    project_dict = []
    
    data_dict = collections.defaultdict(list) 
    wall_dict = collections.defaultdict(list)
    floor_dict = collections.defaultdict(list)
    covering_dict = collections.defaultdict(list)
    beam_dict = collections.defaultdict(list)
    column_dict = collections.defaultdict(list)
    
    
    for j in project:
        project_dict.append(filename)
        project_dict.append(j.OwnerHistory.OwningApplication.ApplicationFullName)
        project_dict.append(j.OwnerHistory.OwningApplication.Version)
        project_dict.append(j.OwnerHistory.OwningUser.ThePerson.FamilyName)
        project_dict.append(j.OwnerHistory.OwningUser.TheOrganization.Name)
        project_dict.append(j.OwnerHistory.OwningUser.Roles)
        project_dict.append(j.Name)
        project_dict.append(j.Description)
        
        
    for i in products:

        data_dict[get_globalid(i)] = [  get_ifcproduct(i),
                                        get_ifcproduct_name(i),
                                        get_type_name(i),
                                        get_building_storey(i),
                                        get_classification(i),
                                        
                                        get_height(i),
                                        get_length(i),
                                        get_width(i),
                                        get_area(i),
                                        get_volume(i),
                                        get_perimeter(i),
                                        
                                        get_phase(i)
                                      
                                        ]
        
        
        if i.is_a().startswith('IfcWall'):
            wall_dict[get_globalid(i)] = [  get_ifcproduct(i),
                                            get_ifcproduct_name(i),
                                            get_type_name(i),
                                            get_building_storey(i),
                                            get_classification(i),
                                            
                                            get_height(i),
                                            get_length(i),
                                            get_width(i),
                                            get_area(i),
                                            get_volume(i),
                                            
                                            get_phase(i)
                                            
                                            ]
            
        if i.is_a().startswith('IfcSlab'):
            floor_dict[get_globalid(i)] = [ get_ifcproduct(i),
                                            get_ifcproduct_name(i),
                                            get_type_name(i),
                                            get_building_storey(i),
                                            get_classification(i),
                                            
                                            get_thickness(i),
                                            get_perimeter(i),
                                            get_area(i),
                                            get_volume(i),
                                            
                                            get_phase(i)
                                            
                                            ]
            
        if i.is_a().startswith('IfcCovering'):
            covering_dict[get_globalid(i)] = [ get_ifcproduct(i),
                                              get_ifcproduct_name(i),
                                              get_type_name(i),
                                              
                                              get_building_storey(i),
                                              get_classification(i),
                                            
                                              get_thickness(i),
                                              get_perimeter(i),
                                              get_area(i),
                                              get_volume(i),
                                              
                                              get_phase(i)
                                            
                                            ]
            
        if i.is_a().startswith('IfcBeam'):
            beam_dict[get_globalid(i)] = [  get_ifcproduct(i),
                                            get_ifcproduct_name(i),
                                            get_type_name(i),
                                            get_building_storey(i),
                                            get_classification(i),
                                            
                                            get_length(i),
                                            
                                            get_phase(i)
                                            
                                            ]
            
        if i.is_a().startswith('IfcColumn'):
            column_dict[get_globalid(i)] = [    get_ifcproduct(i),
                                                get_ifcproduct_name(i),
                                                get_type_name(i),
                                                get_building_storey(i),
                                                get_classification(i),
                                            
                                                get_height(i),
                                                
                                                get_phase(i)
                                            
                                            ]
            
                
      
                                
    write_to_excel( projectdata=project_dict,
                    ifcdata=data_dict, 
                    walldata=wall_dict,
                    floordata=floor_dict, 
                    coveringdata=covering_dict,
                    beamdata=beam_dict,
                    columndata=column_dict)   
    

    return products
    

def get_globalid(i):
    
    return i.GlobalId
    
def get_ifcproduct(i):

    return i.is_a()
        
def get_ifcproduct_name(i):
    
    name_list=[]

    #If Revit Export, remove BATid after colon
    if i.Name:
        if len(i.Name.split(':', 2 )) == 3:
            name_list.append(i.Name.split(':',2)[0] + ':' + i.Name.split(':',2)[1])
        else:
            name_list.append(i.Name)
    
    for i in name_list:
        return i
    
def get_type_name(i):

    if i.IsDefinedBy:
        for j in i.IsDefinedBy:
            if j.is_a('IfcRelDefinesByType'):
                if j.RelatingType.Name:
                    return j.RelatingType.Name
    
def get_building_storey(i):
    
    ifcproducts_without_building_storey = [     'IfcSite',
                                                'IfcBuilding',
                                                'IfcBuildingStorey',
                                                'IfcOpeningElement', 
                                                'IfcSpace'
                                                ]
    
    if i.is_a() not in ifcproducts_without_building_storey:
        if i.ContainedInStructure:    
            for contained_in_spatial_structure in  i.ContainedInStructure:
                if contained_in_spatial_structure.RelatingStructure.is_a('IfcBuildingStorey'):
                    return contained_in_spatial_structure.RelatingStructure.Name 
                

    
def get_classification(i):
    
    if i.HasAssociations:
        for has_associations in i.HasAssociations:
            if has_associations.is_a('IfcRelAssociatesClassification'):
                return has_associations.RelatingClassification.ItemReference
    
    
    
def get_materials(i):
    
    #For future use to retrieve materials per object

    material_list = []
    if i.HasAssociations:
        for j in i.HasAssociations:
            if j.is_a('IfcRelAssociatesMaterial'):
                
                if j.RelatingMaterial.is_a('IfcMaterial'):
                    material_list.append(j.RelatingMaterial.Name)
                   
                if j.RelatingMaterial.is_a('IfcMaterialList'):
                    for materials in j.RelatingMaterial.Materials:
                        material_list.append(materials.Name)
               
                        
                if j.RelatingMaterial.is_a('IfcMaterialLayerSetUsage'):
                    for materials in j.RelatingMaterial.ForLayerSet.MaterialLayers:
                        material_list.append(materials.Material.Name)
                        
                else:
                    pass
                
    return material_list

def get_thickness(i):
    
    for properties in i.IsDefinedBy:
        if properties.is_a('IfcRelDefinesByProperties'):
            if properties.RelatingPropertyDefinition.is_a('IfcPropertySet'):
                    if properties.RelatingPropertyDefinition.Name == "Dimensions":
                        for dimensions in properties.RelatingPropertyDefinition.HasProperties:
                            if dimensions.Name == "Thickness":
                                return dimensions.NominalValue[0]
                                continue
            
                           
            if properties.RelatingPropertyDefinition.is_a('IfcElementQuantity'):
                for quantities in properties.RelatingPropertyDefinition.Quantities:
                    if quantities.Name == 'Width':
                        return round(quantities.LengthValue, 0)
                        continue
                    
            if properties.RelatingPropertyDefinition.is_a('IfcElementQuantity'):
                for quantities in properties.RelatingPropertyDefinition.Quantities:
                    if quantities.Name == 'Thickness':
                        return round(quantities.LengthValue, 0)
                        continue
                    
            if properties.RelatingPropertyDefinition.Name == "Construction":
                for properties in properties.RelatingPropertyDefinition.HasProperties:
                    if properties.Name == 'Thickness':
                        return round(properties.NominalValue[0], 0)
                


def get_height(i):

    for properties in i.IsDefinedBy:
        if properties.is_a('IfcRelDefinesByProperties'):
        
            if properties.RelatingPropertyDefinition.is_a('IfcElementQuantity'):
              
                for quantities in properties.RelatingPropertyDefinition.Quantities:
                    if quantities.Name == 'Height':
                        return round(quantities.LengthValue, 0)
                        continue 
                    
                    #For Columns
                    if quantities.Name == 'Length':
                        return round(quantities.LengthValue, 0)
                        continue  
                                       
                      
            if properties.RelatingPropertyDefinition.is_a('IfcElementQuantity'): 
                for quantities in properties.RelatingPropertyDefinition.Quantities:
                    if quantities.Name == 'Height (Z Size)':
                        return round(quantities.LengthValue, 0)
                    
                    
                   
def get_length(i):

    for properties in i.IsDefinedBy:
        if properties.is_a('IfcRelDefinesByProperties'):
            if properties.RelatingPropertyDefinition.is_a('IfcPropertySet'):
                    if properties.RelatingPropertyDefinition.Name == "Dimensions":
                        for dimensions in properties.RelatingPropertyDefinition.HasProperties:
                            if dimensions.Name == "Length":
                                return dimensions.NominalValue[0]
                                continue

            if properties.RelatingPropertyDefinition.is_a('IfcElementQuantity'):
                for quantities in properties.RelatingPropertyDefinition.Quantities:
                    if quantities.Name == 'Length':
                        return round(quantities.LengthValue, 0)
                        continue
                
                    
            if properties.RelatingPropertyDefinition.Name == "ArchiCADQuantities":
                for q in properties.RelatingPropertyDefinition.Quantities:
                    if q.Name == '3D Length':
                        return round(q.LengthValue, 0)
                        
                    
def get_width(i):
    
    for properties in i.IsDefinedBy:
        if properties.is_a('IfcRelDefinesByProperties'):
            if properties.RelatingPropertyDefinition.is_a('IfcElementQuantity'):
                for quantities in properties.RelatingPropertyDefinition.Quantities:
                    if quantities.Name == 'Width':
                        return round(quantities.LengthValue, 0)
    
def get_area(i): 
    
    for properties in i.IsDefinedBy:
        if properties.is_a('IfcRelDefinesByProperties'):
            if properties.RelatingPropertyDefinition:

                if properties.RelatingPropertyDefinition.is_a('IfcPropertySet'):
                    if properties.RelatingPropertyDefinition.Name == "Dimensions":
                        for dimensions in properties.RelatingPropertyDefinition.HasProperties:
                            if dimensions.Name == "Area":
                                return dimensions.NominalValue[0]
                                continue
                            
                    
            if properties.RelatingPropertyDefinition.is_a('IfcElementQuantity'):
                for quantities in properties.RelatingPropertyDefinition.Quantities:
                    
                    if quantities.Name == 'GrossSideArea':
                        return round(quantities.AreaValue, 3) #/1000000000
                        continue  
                    
                    if quantities.Name == 'GrossArea':
                        return quantities.AreaValue
                        continue 
                    
                    if quantities.Name == 'Area':
                        return quantities.AreaValue
                        continue
                    
                    if quantities.Name == 'NetSideArea':
                        return quantities.AreaValue
                        continue
                    
                    if quantities.Name == 'NetArea':
                        return quantities.AreaValue
                  
                    
def get_volume(i): 
      
    for properties in i.IsDefinedBy:
        if properties.is_a('IfcRelDefinesByProperties'):
            
            if properties.RelatingPropertyDefinition.is_a('IfcPropertySet'):
                    if properties.RelatingPropertyDefinition.Name == "Dimensions":
                        for dimensions in properties.RelatingPropertyDefinition.HasProperties:
                            if dimensions.Name == "Volume":
                                return dimensions.NominalValue[0]
                                continue     
                            
            if properties.RelatingPropertyDefinition.is_a('IfcElementQuantity'):
                for quantities in properties.RelatingPropertyDefinition.Quantities:
                    if quantities.Name == 'GrossVolume':
                        return round(quantities.VolumeValue, 3)
                        continue
                    
                    if quantities.Name == 'NetVolume':
                        return round(quantities.VolumeValue, 3)
                        continue
                    
                    if quantities.Name == 'Volume':
                        return round(quantities.VolumeValue, 3)
                        continue
                    
                    
                    if quantities.Name == 'Net Volume':
                        return round(quantities.VolumeValue, 3)
                  
                    


def get_perimeter(i):
    
    for properties in i.IsDefinedBy:
        if properties.is_a('IfcRelDefinesByProperties'):
            if properties.RelatingPropertyDefinition.is_a('IfcElementQuantity'):
                for quantities in properties.RelatingPropertyDefinition.Quantities:
                    if quantities.Name == 'Perimeter':
                        return round(quantities.LengthValue, 0)    
                    
            
def get_phase(i):

    for properties in i.IsDefinedBy:
        if properties.is_a('IfcRelDefinesByProperties'):
            if properties.RelatingPropertyDefinition.is_a('IfcPropertySet'):
                
                if properties.RelatingPropertyDefinition.Name == "Phasing":
                    for phasing in properties.RelatingPropertyDefinition.HasProperties:
                        return phasing.NominalValue[0]
                        continue 
                    
                if properties.RelatingPropertyDefinition.Name == "AC_Pset_RenovationAndPhasing":
                    for phasing in properties.RelatingPropertyDefinition.HasProperties:
                        return phasing.NominalValue[0]
                        continue
                    
                if properties.RelatingPropertyDefinition.Name == "CPset_Phasing":
                    for phasing in properties.RelatingPropertyDefinition.HasProperties:
                        return phasing.NominalValue[0]
                    
            
            

header_list = [     'GlobalId',
                    'IfcProduct',
                    'Name',
                    'Type',
                    'Building Storey',
                    'Classification',
                    'Height (mm)',
                    'Length (mm)',
                    'Width (mm)',
                    'Area (m2)',
                    'Volume (m3)',
                    'Perimeter (mm)',
                    'Phase'
                    ]




################################################
############# Write Data to Excel ##############
################################################

def write_to_excel(projectdata, ifcdata, walldata, floordata, coveringdata, beamdata, columndata):
    print 'Workbook is being created'
 
        
    if len(projectdata) > 0:
        head, tail = os.path.split(str(projectdata[0]))
        
        print head 
    
        #Create a workbook
        workbook = xlsxwriter.Workbook(str(tail).replace('.ifc','').replace(' ','_') + '_IFC_meetstaat.xlsx')
        
      
        worksheet_project = workbook.add_worksheet('Ifc Project Info')
        worksheet_project.set_column('B:B', 10)
        
        worksheet = workbook.add_worksheet('Ifc Data Total')

      
        if len(walldata) > 0:
            worksheet_walls = workbook.add_worksheet('IfcWall')
            
        if len(floordata) > 0:
            worksheet_floors = workbook.add_worksheet('IfcSlab')
        
        if len(coveringdata) > 0:
            worksheet_coverings = workbook.add_worksheet('IfcCovering')
            
        if len(beamdata) > 0:
            worksheet_beams = workbook.add_worksheet('IfcBeam')
            
        if len(columndata) > 0:
            worksheet_columns = workbook.add_worksheet('IfcColumns')
            
            
            
        #####################################################################
        ################### Format initialization ###########################
        #####################################################################   
            
        cell_chb_red = workbook.add_format()
        cell_chb_red.set_pattern(1)
        cell_chb_red.set_bg_color(chb_red)
        cell_chb_red.set_font_size(12)
        cell_chb_red.set_font_color('white')
        
        cell_format_project = workbook.add_format()
        #cell_format.set_rotation(10)
        cell_format_project.set_pattern(1)
        cell_format_project.set_bg_color(chb_grey)
        cell_format_project.set_font_size(12)
        #cell_format_project.set_font_color('white')
        
        

        #####################################################################
        ################### Worksheet Project Information ###################
        #####################################################################
        
        
        worksheet_project.merge_range('A1:B1', 'Projectinformatie', cell_chb_red)
        project_parameters=['Bestandslocatie','Software','Versie','Persoon','Organisatie','Functie','Naam','Project','Tijd van meetstaat export']
    
        for i, v in enumerate(projectdata):
            worksheet_project.write("B"+str(i+2)+'"', str(v), cell_format_project) 
            #worksheet_project.write("B"+str(i+2)+'"', str(v), cell_format_project)
         
        for i, v in enumerate(project_parameters):
            worksheet_project.write("A"+str(i+2)+'"', str(v), cell_format_project)  
            
        worksheet_project.write("B10", datetime.datetime.now().strftime("%A, %d. %B %Y %I:%M%p"), cell_format_project)
        worksheet_project.freeze_panes(1, 0)

        
        
        #####################################################################
        #################### Worksheet Total Information ####################
        #####################################################################
        worksheet.add_table('A2:M'+str(len(ifcdata)+2)+"'", {'data': {},
                                                                'style': 'Table Style Medium 17',
                                                                'columns': [{'header': 'GlobalId'},
                                                                            {'header': 'IfcProduct'},
                                                                            {'header': 'Name'},
                                                                            {'header': 'Type'},
                                                                            {'header': 'Building Storey'},
                                                                            {'header': 'Classification'},
                                                                            {'header': 'Height (mm)'},
                                                                            {'header': 'Length (mm)'},
                                                                            {'header': 'Width (mm)'},
                                                                            {'header': 'Area (m2)'},
                                                                            {'header': 'Volume (m3)'},
                                                                            {'header': 'Perimeter (mm)'},
                                                                            {'header': 'Phase'}
                                                                            ]})
        
        
                                    
        if len(walldata) > 0:
            worksheet_walls.add_table('A2:L'+str(len(walldata)+2)+"'",{    'data': {},
                                                                    'style': 'Table Style Medium 17',
                                                                    'columns': [{'header': 'GlobalId'},
                                                                                {'header': 'IfcProduct'},
                                                                                {'header': 'Name'},
                                                                                {'header': 'Type'},
                                                                                {'header': 'Building Storey'},
                                                                                {'header': 'Classification'},
                                                                                {'header': 'Height (mm)'},
                                                                                {'header': 'Length (mm)'},
                                                                                {'header': 'Width (mm)'},
                                                                                {'header': 'Area (m2)'},
                                                                                {'header': 'Volume (m3)'},
                                                                                {'header': 'Phase'}
                                                                                ]})
          
    
        if len(floordata) > 0:
            worksheet_floors.add_table('A2:K'+str(len(floordata)+2)+"'",{    'data': {},
                                                                    'style': 'Table Style Medium 17',
                                                                    'columns': [{'header': 'GlobalId'},
                                                                                {'header': 'IfcProduct'},
                                                                                {'header': 'Name'},
                                                                                {'header': 'Type'},
                                                                                {'header': 'Building Storey'},
                                                                                {'header': 'Classification'},
                                                                                {'header': 'Thickness (mm)'},
                                                                                {'header': 'Perimeter (mm)'},
                                                                                {'header': 'Area (m2)'},
                                                                                {'header': 'Volume (m3)'},
                                                                                {'header': 'Phase'}
                                                                                ]})
        
        if len(coveringdata) > 0:
            worksheet_coverings.add_table('A2:K'+str(len(coveringdata)+2)+"'",{    'data': {},
                                                                    'style': 'Table Style Medium 17',
                                                                    'columns': [{'header': 'GlobalId'},
                                                                                {'header': 'IfcProduct'},
                                                                                {'header': 'Name'},
                                                                                {'header': 'Type'},
                                                                                {'header': 'Building Storey'},
                                                                                {'header': 'Classification'},
                                                                                {'header': 'Thickness (mm)'},
                                                                                {'header': 'Perimeter (mm)'},
                                                                                {'header': 'Area (m2)'},
                                                                                {'header': 'Volume (m3)'},
                                                                                {'header': 'Phase'}
                                                                                ]})
        
        
        if len(beamdata) > 0:
            worksheet_beams.add_table('A2:H'+str(len(beamdata)+2)+"'",{    'data': {},
                                                                    'style': 'Table Style Medium 17',
                                                                    'columns': [{'header': 'GlobalId'},
                                                                                {'header': 'IfcProduct'},
                                                                                {'header': 'Name'},
                                                                                {'header': 'Type'},
                                                                                {'header': 'Building Storey'},
                                                                                {'header': 'Classification'},
                                                                                {'header': 'Length (mm)'},
                                                                                {'header': 'Phase'}
                                                                                
                                                                                ]})
        if len(columndata) > 0:
            worksheet_columns.add_table('A2:H'+str(len(columndata)+2)+"'",{    'data': {},
                                                                    'style': 'Table Style Medium 17',
                                                                    'columns': [{'header': 'GlobalId'},
                                                                                {'header': 'IfcProduct'},
                                                                                {'header': 'Name'},
                                                                                {'header': 'Type'},
                                                                                {'header': 'Building Storey'},
                                                                                {'header': 'Classification'},
                                                                                {'header': 'Height (mm)'},
                                                                                {'header': 'Phase'}
                                                                                
                                                                                ]})
        
        
        row=2
        col=0
    
        for k, v in ifcdata.iteritems():
            for i in range(0, len(v)):
                worksheet.write(row, col, k)
                worksheet.write(row, i+1, v[i])
                 
            row += 1
            
 
           
        if len(walldata) > 0:
            row_wall = 2 
                
            for k, v in walldata.iteritems():
                for i in range(0, len(v)):  
                    worksheet_walls.write(row_wall, col, k)
                    worksheet_walls.write(row_wall, i+1, v[i]) 
                   
                row_wall += 1 
            
            
        if len(floordata) > 0:
            row_floor = 2
            
            for k, v in floordata.iteritems(): 
                for i in range(0, len(v)):  
                    worksheet_floors.write(row_floor, col, k)
                    worksheet_floors.write(row_floor, i+1, v[i]) 
                   
                row_floor += 1 
           
         
        if len(coveringdata) > 0: 
            row_coverings = 2
            
            for k, v in coveringdata.iteritems(): 
                for i in range(0, len(v)):  
                    worksheet_coverings.write(row_coverings, col, k)
                    worksheet_coverings.write(row_coverings, i+1, v[i]) 
                   
                row_coverings += 1 
            
            
        if len(beamdata) > 0:   
            row_beam = 2
            
            for k,v in beamdata.iteritems():
                for i in range(0, len(v)):
                    worksheet_beams.write(row_beam, col, k)
                    worksheet_beams.write(row_beam, i+1, v[i])
                    
                row_beam += 1
            
            
        if len(columndata) > 0:   
            row_column = 2
            
            for k,v in columndata.iteritems():
                for i in range(0, len(v)):
                    worksheet_columns.write(row_column,col, k)
                    worksheet_columns.write(row_column, i+1, v[i])
                    
                row_column += 1 
            
            
              
        worksheet.freeze_panes(2, 0)
        
        if len(walldata) > 0:
            worksheet_walls.freeze_panes(2, 0)
            
        if len(floordata) > 0:
            worksheet_floors.freeze_panes(2, 0)
        
        if len(coveringdata) > 0:
            worksheet_coverings.freeze_panes(2, 0)
          
        if len(beamdata) > 0:  
            worksheet_beams.freeze_panes(2,0)
        
        if len(columndata) > 0:    
            worksheet_columns.freeze_panes(2,0)
            
        cell_white = workbook.add_format()
        cell_white.set_pattern(1)
        cell_white.set_bg_color(white)
            
        cell_format = workbook.add_format()
        #cell_format.set_rotation(10)
        cell_format.set_pattern(1)
        cell_format.set_bg_color(chb_red)
        cell_format.set_font_size(12)
        cell_format.set_font_color('white')
        
        column_letters=['A','B','C','D','E','F','G','H','I','J','K','L','M']
        
        column_letters_walls=['A','B','C','D','E','F','G','H','I','J','K','L']
        column_letters_floors=['A','B','C','D','E','F','G','H','I','K']
        column_letters_coverings=['A','B','C','D','E','F','G','H','I','J','K']
        
        column_letters_beams=['A','B','C','D','E','F','G','H']
        column_letters_columns=['A','B','C','D','E','F','G','H']
        
        
        for column_letter in column_letters:
            worksheet.write(str(column_letter)+"1", '', cell_format)
 
        for i in column_letters_walls:
            if len(walldata) > 0:
                worksheet_walls.write(str(i)+"1", '', cell_format)

        for i in column_letters_floors:
            if len(floordata) > 0:
                worksheet_floors.write(str(i)+"1", '', cell_format)
          
        for i in column_letters_coverings:  
            if len(coveringdata) > 0:
                worksheet_coverings.write(str(i)+"1", '', cell_format)
                
        for i in column_letters_beams:    
            if len(beamdata) > 0:
                worksheet_beams.write(str(i)+"1", '', cell_format)

        for i in column_letters_columns:      
            if len(columndata) > 0:   
                worksheet_columns.write(str(i)+"1", '', cell_format)
                
                
                
        ###########################################################################
        ############################## Write Headers ##############################
        ###########################################################################       
                 
        
        header_total = {    'A2':'GlobalId',
                            'B2':'IfcProduct',
                            'C2':'Name',
                            'D2':'Type',
                            'E2':'Building Storey',
                            'F2':'Classification',
                            'G2':'Height (mm)',
                            'H2':'Length (mm)',
                            'I2':'Width (mm)',
                            'J2':'Area (m2)',
                            'K2':'Volume (m3)',
                            'L2':'Perimeter (mm)',
                            'M2':'Phase'}

        for k,v in header_total.iteritems():
            worksheet.write(k, v, cell_format)
         
        wall_header_dict={'A2':'GlobalId',
                          'B2':'IfcProduct',
                          'C2':'Name',
                          'D2':'Type',
                          'E2':'Building Storey',
                          'F2':'Classification',
                          'G2':'Height (mm)',
                          'H2':'Length (mm)',
                          'I2':'Width (mm)',
                          'J2':'Area (m2)',
                          'K2':'Volume (m3)',
                          'L2':'Phase'
                         }
        
        if len(walldata) > 0:
            for k, v in wall_header_dict.iteritems():
                worksheet_walls.write(str(k),str(v), cell_format)
                

            
        floor_header_dict={'A2': 'GlobalId',
                           'B2': 'IfcProduct',
                           'C2': 'Name',
                           'D2': 'Type',
                           'E2': 'Building Storey',
                           'F2': 'Classification',
                           'G2': 'Thickness (mm)',
                           'H2': 'Perimeter (mm)',
                           'I2': 'Area (m2)',
                           'J2': 'Volume (m3)',
                           'K2': 'Phase'
                            }
        
        if len(floordata) > 0:
            for k, v in floor_header_dict.iteritems():
                worksheet_floors.write(str(k), str(v), cell_format)
                
                

            
        covering_header_dict={  'A2':'GlobalId',
                                'B2':'IfcProduct',
                                'C2':'Name',
                                'D2':'Type',
                                'E2':'Building Storey',
                                'F2':'Classification',
                                'G2':'Thickness (mm)',
                                'H2':'Perimeter (mm)',
                                'I2':'Area (m2)',
                                'J2':'Volume (m3)',
                                'K2':'Phase'}

        if len(coveringdata) > 0:
            for k, v in covering_header_dict.iteritems():
                worksheet_coverings.write(str(k), str(v), cell_format)
                

            
        beam_header_dict={ 'A2':'GlobalId',
                           'B2':'IfcProduct',
                           'C2':'Name',
                           'D2':'Type',
                           'E2':'Building Storey',
                           'F2':'Classification',
                           'G2':'Length (mm)',
                           'H2':'Phase' }
     
        if len(beamdata) > 0:
            for k, v in beam_header_dict.iteritems():
                worksheet_beams.write(str(k), str(v), cell_format)
                
            
        column_header_dict={'A2':'GlobalId',
                            'B2':'IfcProduct',
                            'C2':'Name',
                            'D2':'Type',
                            'E2':'Building Storey',
                            'F2':'Classification',
                            'G2':'Height (mm)',
                            'H2':'Phase'}
        
        if len(columndata) > 0:
            for k, v in column_header_dict.iteritems():
                worksheet_columns.write(str(k), str(v), cell_format)
                
                

        
        totaal_objecten='=IF(SUBTOTAL(103,$A:$A)=2,SUBTOTAL(103,$A:$A)-1&" object",SUBTOTAL(103,$A:$A)-1&" objecten")'
        #totaal_lengte='=TEXT(SUBTOTAL(9,OFFSET(I$10,0,0,COUNTA($A:$A),1)),"#.###,000")&" mm"'
        
        totaal_oppervlakte='=TEXT(SUBTOTAL(9,OFFSET(J$3,0,0,COUNTA($A:$A),1)),"#.###,000")&" m2"'
        totaal_volume='=TEXT(SUBTOTAL(9,OFFSET(K$3,0,0,COUNTA($A:$A),1)),"#.###,000")&" m3"'
        
        totaal_oppervlakte_floor= '=TEXT(SUBTOTAL(9,OFFSET(I$3,0,0,COUNTA($A:$A),1)),"#.###,000")&" m2"'
        totaal_volume_floor='=TEXT(SUBTOTAL(9,OFFSET(J$3,0,0,COUNTA($A:$A),1)),"#.###,000")&" m3"'
        
        totaal_oppervlakte_coverings='=TEXT(SUBTOTAL(9,OFFSET(I$3,0,0,COUNTA($A:$A),1)),"#.###,000")&" m2"'
        totaal_volume_coverings='=TEXT(SUBTOTAL(9,OFFSET(J$3,0,0,COUNTA($A:$A),1)),"#.###,000")&" m3"'
        
        totaal_lengte_beams='=TEXT(SUBTOTAL(9,OFFSET(G$3,0,0,COUNTA($A:$A),1)),"#.###,000")&" mm"'
        totaal_lengte_columns='=TEXT(SUBTOTAL(9,OFFSET(G$3,0,0,COUNTA($A:$A),1)),"#.###,000")&" mm"'
        
        
        worksheet.write_formula('C1', totaal_objecten, cell_format)
        worksheet.write_formula('J1', totaal_oppervlakte, cell_format)
        worksheet.write_formula('K1', totaal_volume, cell_format)
        
        if len(walldata) > 0:
            worksheet_walls.write_formula('C1', totaal_objecten, cell_format)
            worksheet_walls.write_formula('J1', totaal_oppervlakte, cell_format)
            worksheet_walls.write_formula('K1', totaal_volume, cell_format)
            
        if len(floordata) > 0:
            worksheet_floors.write_formula('C1', totaal_objecten, cell_format)
            worksheet_floors.write_formula('I1', totaal_oppervlakte_floor, cell_format)
            worksheet_floors.write_formula('J1', totaal_volume_floor, cell_format)
        
        if len(coveringdata) >0:
            worksheet_coverings.write_formula('C1', totaal_objecten, cell_format)
            worksheet_coverings.write_formula('I1', totaal_oppervlakte_coverings, cell_format)
            worksheet_coverings.write_formula('J1', totaal_volume_coverings, cell_format)
            
        if len(beamdata) > 0:
            worksheet_beams.write_formula('C1', totaal_objecten, cell_format)
            worksheet_beams.write_formula('G1', totaal_lengte_beams, cell_format)
        
        if len(columndata) > 0:
            worksheet_columns.write_formula('C1', totaal_objecten, cell_format)
            worksheet_columns.write_formula('G1', totaal_lengte_columns, cell_format)
        
        
        
        worksheet_project.set_column('A:A', 30)
        worksheet_project.set_column('B:B', 80)
        worksheet_project.set_row(0, 60)
        
       
        worksheet.set_column('C:C', 60)
        worksheet.set_column('D:D', 60)
        
        worksheet.set_column('E:E', 20)
        
        worksheet.set_column('I:I', 20)
        worksheet.set_column('J:J', 20)
        worksheet.set_column('K:K', 20)
        worksheet.set_column('M:M', 20)
        
        worksheet.set_row(0,60)
        
        if len(walldata) > 0:
            worksheet_walls.set_column('C:C', 60)
            worksheet_walls.set_column('D:D', 60)
            worksheet_walls.set_column('E:E', 20)
            worksheet_walls.set_column('I:I', 20)
            worksheet_walls.set_column('J:J', 20)
            worksheet_walls.set_column('K:K', 20)
            worksheet_walls.set_column('L:L', 20)
            worksheet_walls.set_row(0,60)
        
        if len(floordata) > 0:
            worksheet_floors.set_column('C:C', 60)
            worksheet_floors.set_column('D:D', 60)
            worksheet_floors.set_column('E:E', 20)
            worksheet_floors.set_column('I:I', 20)
            worksheet_floors.set_column('J:J', 20)
            worksheet_floors.set_column('K:K', 20)
            worksheet_floors.set_row(0,60)
        
        if len(coveringdata) > 0:
            worksheet_coverings.set_column('C:C', 60)
            worksheet_coverings.set_column('D:D', 60)
            worksheet_coverings.set_column('E:E', 20)
            worksheet_coverings.set_column('I:I', 20)
            worksheet_coverings.set_column('J:J', 20)
            worksheet_coverings.set_column('K:K', 20)
            worksheet_coverings.set_row(0,60)
        
        if len(beamdata) > 0:
            worksheet_beams.set_column('C:C', 60)
            worksheet_beams.set_column('D:D', 60)
            worksheet_beams.set_column('E:E', 20)
            worksheet_beams.set_column('G:G', 20)
            worksheet_beams.set_column('H:H', 20)
            worksheet_beams.set_row(0,60)
    
        
        if len(columndata) > 0:
            worksheet_columns.set_column('C:C', 60)
            worksheet_columns.set_column('D:D', 60)
            worksheet_columns.set_column('E:E', 20)
            worksheet_columns.set_column('G:G', 20)
            worksheet_columns.set_column('H:H', 20)
            worksheet_columns.set_row(0,60)
            
    workbook.close()
    
    
    
def write_to_file(export_file):
    print 'Writing ', export_file, ' to excel'

    
    head, tail = os.path.split(str(export_file))
    
    meetstaat_file = tail.replace('.ifc','').replace(' ','_') + '_IFC_meetstaat.xlsx'
    
    excel_file = "excel//" + meetstaat_file
    

    if len(meetstaat_file) > 0:
  
        
        #move file to excel folder
        shutil.move(meetstaat_file, excel_file)
        
        #get absolute path
        os.path.abspath(excel_file)
        
        #open file from absolute path
        as_path = os.path.abspath(excel_file)
        
        #replace single backslash with double backslashes so files can be opened from the sever
        chb_ifcexcel_file = str(as_path).replace('\\','\\\\')

        #open file
        os.startfile(chb_ifcexcel_file)

  
    print 'Writing to Excel has been finished'

    return excel_file
        

        


    
if __name__ == '__main__':
    ifc_filename = None
        
    if ifc_filename is not None:
        
        write_to_file(export_file=ifcfile(filename=ifc_filename))


       

 
    
    
    

    
    
 




