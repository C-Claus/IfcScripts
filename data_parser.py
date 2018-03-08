import ifcopenshell
import csv
import time
import subprocess
from subprocess import call 
import sys
import os 
from os import listdir
import glob
from collections import OrderedDict


ifcfiles = raw_input('Enter Path: ')


for IFC_file in ifcfiles:
    ifcfile = ifcopenshell.open(IFC_file)


#initialisation of global variables
elements = ifcfile.by_type('IfcElement')
products = ifcfile.by_type('IfcProduct')
application = ifcfile.by_type('IfcApplication')
project = ifcfile.by_type('IfcProject')
site = ifcfile.by_type('IfcSite')
root = ifcfile.by_type('IfcRoot')
objects = ifcfile.by_type('IfcObject')



def check_software(ifcapplication): 
    software_list = []
    if ifcapplication:
        software_list.append(ifcapplication.ApplicationFullName)
        
    return software_list

def get_projectinfo(ifc_file):
    if project:
        print 'Project Name: ',project[0].Name   
    if application:
        print 'Application: ',application[0].ApplicationFullName    
    if site:
        print 'Latitude: ', site[0].RefLatitude
        print 'Longitude: ', site[0].RefLongitude
        print 'Elevation: ', site[0].RefElevation 
        

def get_global_id(ifcproduct):
    global_id_list = []    
    
    if ifcproduct.GlobalId:
        global_id_list.append(ifcproduct.GlobalId)
        
    return global_id_list 


def get_product(ifcproduct):
    product_list = []
    
    if ifcproduct:
        product_list.append(ifcproduct.is_a())

    return product_list   



def get_building_storey(ifcproduct):
    building_storey_list = []
     
    try:
        if ifcproduct:
            if ifcproduct.ContainedInStructure:            
                if ifcproduct.ContainedInStructure[0].RelatingStructure.is_a() == 'IfcBuildingStorey':
                    building_storey_list.append(ifcproduct.ContainedInStructure[0].RelatingStructure.Name)
    except:
        pass
    else:
        pass

    #IfcOpeningElement should not be linked directly to the spatial structure of the project,
    #i.e. the inverse relationship ContainedInStructure shall be NIL.
    #It is assigned to the spatial structure through the elements it penetrates.
    if len(building_storey_list) == 0:
        building_storey_list.append('N/A')
         
    return building_storey_list  
 

def get_product_name(ifcproduct):
    product_name_list = []
    
    if ifcproduct:
        product_name_list.append(ifcproduct.Name)
     
    if len(product_name_list) == 0:
        product_name_list.append('N/A')
        
    if product_name_list[0] == None:
        product_name_list.append('N/A')
        
    return product_name_list

def get_presentationlayerassignment(ifcproduct):
    presentationlayer_list = []
    
    if ifcproduct:
        if  ifcproduct.Representation != None: 
            for i in  ifcproduct.Representation.Representations:
                if i.RepresentationIdentifier == 'Body':
                    presentationlayer_list.append(i.LayerAssignments[0].Name)
        
    if len(presentationlayer_list) == 0:
        presentationlayer_list.append('N/A')
        
    if presentationlayer_list[0] == None:
        presentationlayer_list.append('N/A')
                     
    return presentationlayer_list 


def get_producttype_name(ifcproduct):
    type_name_list = []
    
    if ifcproduct:
        type_name_list.append(ifcproduct.ObjectType)
        
    if len(type_name_list) == 0:
        type_name_list.append('N/A')
        
    if type_name_list[0] == None:
        type_name_list.append('N/A')
        
    return type_name_list


                            
def get_assembly_code(ifcproduct):
    
    #Classifications of an object may be referenced from an external source rather than being
    #contained within the IFC model. This is done through the IfcClassificationReference class.
    
    assembly_code_list = []

    if ifcproduct.HasAssociations:
        if ifcproduct.HasAssociations[0].is_a() == 'IfcRelAssociatesClassification':
            assembly_code_list.append(ifcproduct.HasAssociations[0].RelatingClassification.ItemReference)
            
        if ifcproduct.HasAssociations[0].is_a() == 'IfcRelAssociatesMaterial':
            for i in ifcproduct.HasAssociations:
                if i.is_a() == 'IfcRelAssociatesClassification' :
                    assembly_code_list.append(i.RelatingClassification.ItemReference)
            
                                           
    if len(assembly_code_list) == 0:
        assembly_code_list.append('N/A')
        
    if assembly_code_list[0] == None:
        assembly_code_list.append('N/A')
        
    return assembly_code_list


def get_materials(ifcproduct):
    material_list = []
    
    if ifcproduct.HasAssociations:
        for i in ifcproduct.HasAssociations:
            if i.is_a('IfcRelAssociatesMaterial'):
                
                if i.RelatingMaterial.is_a('IfcMaterial'):
                    material_list.append(i.RelatingMaterial.Name)
                   
                if i.RelatingMaterial.is_a('IfcMaterialList'):
                    for materials in i.RelatingMaterial.Materials:
                        material_list.append(materials.Name)
               
                        
                if i.RelatingMaterial.is_a('IfcMaterialLayerSetUsage'):
                    for materials in i.RelatingMaterial.ForLayerSet.MaterialLayers:
                        material_list.append(materials.Material.Name)
                        
                else:
                    pass
                      
    if len(material_list) == 0:
        material_list.append('N/A')
       
    joined_material_list = ' | '.join(material_list)
                         
    return [joined_material_list]        
                    

def get_reference(ifcproduct):  
    reference_list = []
    
    if ifcproduct.IsDefinedBy:
        if ifcproduct.IsDefinedBy[0]:
            if ifcproduct.IsDefinedBy[0].is_a() == 'IfcRelDefinesByProperties':
                if ifcproduct.IsDefinedBy[0].RelatingPropertyDefinition.is_a() == 'IfcPropertySet':
                    for ifcproperty in ifcproduct.IsDefinedBy[0].RelatingPropertyDefinition.HasProperties:
                        if ifcproperty.Name == 'Reference':
                            reference_list.append(ifcproperty.Name)
                            reference_list.append(ifcproperty.NominalValue[0])
                        else:
                            pass
                                    
    if len(reference_list) == 0:
        reference_list.append('N/A')
                        
    return [reference_list]


def get_external_internal(ifcproduct):
    externality_list = []
 
    if ifcproduct.IsDefinedBy:
        if ifcproduct.IsDefinedBy[0]:
            if ifcproduct.IsDefinedBy[0].is_a() == 'IfcRelDefinesByProperties':
                if ifcproduct.IsDefinedBy[0].RelatingPropertyDefinition.is_a() == 'IfcPropertySet':
                    for ifcproperty in ifcproduct.IsDefinedBy[0].RelatingPropertyDefinition.HasProperties:
                        if ifcproperty.Name == 'IsExternal':
                            #externality_list.append(ifcproperty.Name)
                            externality_list.append(ifcproperty.NominalValue[0])
                        else:
                            pass
                                    
    if len(externality_list) == 0:
        externality_list.append('N/A')  

    return [externality_list]

def get_loadbearing(ifcproduct):
    load_bearing_list = []
    
    if ifcproduct.IsDefinedBy:
        if ifcproduct.IsDefinedBy[0]:
            if ifcproduct.IsDefinedBy[0].is_a() == 'IfcRelDefinesByProperties':
                if ifcproduct.IsDefinedBy[0].RelatingPropertyDefinition.is_a() == 'IfcPropertySet':
                    for ifcproperty in ifcproduct.IsDefinedBy[0].RelatingPropertyDefinition.HasProperties:
                        if ifcproperty.Name == 'LoadBearing':
                            load_bearing_list.append(ifcproperty.NominalValue[0])
                        else:
                            pass
                                    
    if len(load_bearing_list) == 0:
        load_bearing_list.append('N/A')
                         
    return [load_bearing_list]

def get_firerating(ifcproduct):
    fire_rating_list = []    
    
    if ifcproduct.IsDefinedBy:
        if ifcproduct.IsDefinedBy[0]:
            if ifcproduct.IsDefinedBy[0].is_a() == 'IfcRelDefinesByProperties':
                if ifcproduct.IsDefinedBy[0].RelatingPropertyDefinition.is_a() == 'IfcPropertySet':
                    for ifcproperty in ifcproduct.IsDefinedBy[0].RelatingPropertyDefinition.HasProperties:
                        if ifcproperty.Name == 'FireRating':
                            fire_rating_list.append(ifcproperty.Name)
                            fire_rating_list.append(ifcproperty.NominalValue[0])
                        else:
                            pass
                                    
    if len(fire_rating_list) == 0:
        fire_rating_list.append('N/A')
                         
    return [fire_rating_list]



def write_to_csv(ifc_file, ifc_applicaton, products):
    
    
    get_projectinfo(ifc_file=file)
    
    
    ifc_files = listdir(str(ifcfiles))
    
    for i in ifc_files:
        
        x = str(ifcfiles) + '/' + str(i)
        
        ifcfile = ifcopenshell.open(x)
        
        with open('csv_bestanden/csv_' + str(i.replace('.ifc', '.csv')), 'w+') as f:
            
        
            #writes header
            fieldnames = [ 
                            'GlobalId', 
                            'IfcProduct', 
                            'IfcBuildingStorey',
                            'Product.Name',
                            'ObjectType',
                            'Assembly Code',
                            'Materials' ,
                            'Reference',
                            'IsExternal',
                            'LoadBearing',
                            'FireRating',
                           
                            ]
            
            header = csv.DictWriter(f, delimiter= ';', fieldnames = fieldnames)
            header.writeheader()
            
            #writes data from ifc to csv
            writer=csv.writer(f,delimiter=';', dialect = 'excel', lineterminator='\n',)
            
            #print len(products)
            products = ifcfile.by_type('IfcProduct')
            #Writes data from the for loop to columns in a CSV file format      
            for product in products:
                    
                print product
                #Exclude the following IfcProduct
                if (product.is_a() != 'IfcSite' and 
                    product.is_a() != 'IfcBuilding' and 
                    product.is_a() != 'IfcBuildingStorey' and
                    product.is_a() != 'IfcOpeningElement' and
                    product.is_a() != 'IfcSpace' ):
    
                    column = []
                    
                   
                    for globalid in get_global_id(ifcproduct=product):
                        if globalid:
                            column.append(globalid)  
                            
                            
                    for ifc_product in get_product(ifcproduct=product):
                        if ifc_product:
                            column.append(ifc_product)
                        
                    for building_storey in get_building_storey(ifcproduct=product):          
                        if ifc_product:     
                            column.append(building_storey)
                        
                    for product_name in get_product_name(ifcproduct=product):
                        if product_name:
                            column.append(product_name)    
                            
                    for type_name in get_producttype_name(ifcproduct=product):
                        if type_name:
                            column.append(type_name)      
            
                    for assembly_code in get_assembly_code(ifcproduct=product):
                        if assembly_code:
                            column.append(assembly_code)
               
                    for material in get_materials(ifcproduct=product):
                        if material:
                            column.append(material)
                            
                    for reference in get_reference(ifcproduct=product):
                        if reference:
                            column.append(reference)
                            
                    for externality in get_external_internal(ifcproduct=product):
                        if externality:
                            column.append(externality)
                          
                    for loadbearing in get_loadbearing(ifcproduct=product):
                        if loadbearing:
                            column.append(loadbearing)
                            
                    for firerating in get_firerating(ifcproduct=product):
                        if firerating:
                            column.append(firerating)
                            
                    writer.writerow(column)
                    
                                  
if __name__ == "__main__":   
      
        
    write_to_csv(ifc_file=ifcfile, ifc_applicaton=application,  products=products)
 
 
   
end = raw_input('Press enter to exit')
    
print 'CSV completed'
    

 




            
