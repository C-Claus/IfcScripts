#coding=utf-8
#C.C.J. Claus 2020
#!/usr/bin/env python


import ifcopenshell 
import lxml 
from lxml import etree as ET
import os 
import uuid
from datetime import datetime
from collections import defaultdict

now = datetime.now()
year = now.strftime("%Y")
month = now.strftime("%m")
day = now.strftime("%d")

fname = 'path_to_file/file.ifc'

ifcfile = ifcopenshell.open(fname)
products = ifcfile.by_type('IfcProduct')
project = ifcfile.by_type('IfcProject')

head, tail = os.path.split(fname)



entity = 'IfcWall'


def get_software():
 
    for i in project:
        return (i.OwnerHistory.OwningApplication.ApplicationFullName)

def get_ifc_entities():
    
    get_quantities(products, entity)
        

def get_quantities(products, entity):
    
    print ("get  quantities") 
    
    
    
    type_list = []
    type_area_list = []
    
    for product in products:
        if product.is_a().startswith(entity):
            for properties in product.IsDefinedBy:
                if properties.is_a('IfcRelDefinesByProperties'):
                    if properties.RelatingPropertyDefinition.is_a('IfcElementQuantity'):
                        for quantities in properties.RelatingPropertyDefinition.Quantities:
                            
                            
                            if entity == 'IfcWall':
                                if "Revit" in get_software():
                                    #GrossSideArea
                                    if quantities.Name ==  'GrossArea' or quantities.Name == 'GrossSideArea':
                                        if product.IsDefinedBy:
                                            for j in product.IsDefinedBy:
                                                if j.is_a("IfcRelDefinesByType"):
                                                    if j.RelatingType.Name:
                                                        type_list.append(str(j.RelatingType.Name))
                                                        type_area_list.append([str(j.RelatingType.Name), quantities.AreaValue])
                                                        #type_area_list.append([str(j.RelatingType.Name).replace('Basic Wall:',''), quantities.AreaValue])
                                else:
                                    #NetSideArea
                                    if quantities.Name == 'NetArea' or quantities.Name == 'NetSideArea':
                                        if product.IsDefinedBy:
                                            for j in product.IsDefinedBy:
                                                if j.is_a("IfcRelDefinesByType"):
                                                    if j.RelatingType.Name:
                                                        type_list.append(str(j.RelatingType.Name))
                                                        type_area_list.append([j.RelatingType.Name, quantities.AreaValue])
                                                        
                                                        
                            if entity == 'IfcSlab':
                                if "Revit" in get_software():    
                                    if quantities.Name == 'GrossArea':
                                        if product.IsDefinedBy:
                                            for j in product.IsDefinedBy:
                                                if j.is_a("IfcRelDefinesByProperties"):   
                                                    if (j.RelatingPropertyDefinition.Name) == 'BaseQuantities':
                                                        if j.RelatingPropertyDefinition.is_a('IfcElementQuantity'):
                                                            for quantities in properties.RelatingPropertyDefinition.Quantities:
                                                                if quantities.Name ==  'GrossArea':
                                                                    type_list.append(str(product.Name.split(':')[1]))
                                                                    type_area_list.append([str(product.Name.split(':')[1]), quantities.AreaValue])
                            
                               
                                else:
                                    if quantities.Name ==  'GrossArea':
                                        if product.IsDefinedBy:
                                            for j in product.IsDefinedBy:
                                                if j.is_a("IfcRelDefinesByType"):
                                                    if j.RelatingType.Name:
                                                        type_list.append(str(j.RelatingType.Name))
                                                        type_area_list.append([str(j.RelatingType.Name), quantities.AreaValue])
                            
                                                              
                            if entity == 'IfcCovering':
                                if quantities.Name ==  'GrossArea':
                                    if product.IsDefinedBy:
                                        for j in product.IsDefinedBy:
                                            if j.is_a("IfcRelDefinesByType"):
                                                if j.RelatingType.Name:
                                                    type_list.append(str(j.RelatingType.Name))
                                                    type_area_list.append([str(j.RelatingType.Name), quantities.AreaValue])                    
     
    print ('type_list:', type_list)       
    result = defaultdict(int)
    result_aslist = defaultdict(list)
    
    for k, v in type_area_list:
        result[k] += v
    
    result_aslist[k].append((v))
                  
    return result



def create_quantities_bcsv(file_name, bouwdeel):
    
    #oppervlakte, lengte, breedte, volume
    
    root = ET.Element('SMARTVIEWSETS')
    doc = ET.SubElement(root, "SMARTVIEWSET")
 
    title = ET.SubElement(doc, "TITLE").text = "CHO smartview " + bouwdeel + tail
    description = ET.SubElement(doc, "DESCRIPTION").text = "Omschrijving volgt hier " 
    guid = ET.SubElement(doc, "GUID" ).text = str(uuid.uuid4())
    modification_date = ET.SubElement(doc, "MODIFICATIONDATE").text = str(datetime.timestamp(now))
    
    smartviews = ET.SubElement(doc, "SMARTVIEWS")
    
    for entity_type, value  in get_quantities(products, entity).items():
       
        
        #I think &#xb2; or &#178; should work.
        smartview = ET.SubElement(smartviews, "SMARTVIEW")
        
        smartview_title = ET.SubElement(smartview, "TITLE").text = str(entity_type) + ' (oppervlakte: ' + str(round(value, 2)) + 'm' + "\u00b2" + ')'
        smartview_description = ET.SubElement(smartview, "DESCRIPTION")
        smartview_creator = ET.SubElement(smartview, "CREATOR").text = "Coen Hagedoorn Bouwgroep"
        smartview_creation_date = ET.SubElement(smartview, "CREATIONDATE").text = day + " " + month + " " + year
        smartview_modifier = ET.SubElement(smartview, "MODIFIER").text = "Coen Hagedoorn Bouwgroep"
        smartview_modification_date = ET.SubElement(smartview, "MODIFICATIONDATE").text = day + " " + month + " " + year
        smartview_guid = ET.SubElement(smartview, "GUID").text = str(uuid.uuid4())
        
       
        rules = ET.SubElement(smartview, "RULES")
        rule = ET.SubElement(rules, "RULE")
        
        ifctype = ET.SubElement(rule, "IFCTYPE").text = str(entity.replace('Ifc',''))
        
        property = ET.SubElement(rule, "PROPERTY")
        property_name = ET.SubElement(property, "NAME").text= "Type"
        propertyset_name = ET.SubElement(property, "PROPERTYSETNAME").text = "Summary"
        property_type = ET.SubElement(property, "TYPE").text = "Summary"
        property_value_type = ET.SubElement(property, "VALUETYPE").text = "StringValue"
        property_unit = ET.SubElement(property, "UNIT").text = "None"
        
        
        condition = ET.SubElement(rule, "CONDITION")
        condition_type = ET.SubElement(condition, "TYPE").text = "Is"
        condition_value = ET.SubElement(condition, "VALUE").text = str(entity_type)
        
        action = ET.SubElement(rule, "ACTION")
        action_type = ET.SubElement(action, "TYPE").text = "Add"
        
        information_take_off = ET.SubElement(smartview, "INFORMATIONTAKEOFF")
        information_take_off_propertysetname = ET.SubElement(information_take_off, "PROPERTYSETNAME").text = str(entity_type)
        information_take_off_propertyname = ET.SubElement(information_take_off, "PROPERTYNAME").text = str(entity_type)
        information_take_off_operation = ET.SubElement(information_take_off, "OPERATION").text = "0"
        
    tree = ET.ElementTree(root)
    tree.write('bcsv_files/' + str(file_name), encoding="utf-8", xml_declaration=True, pretty_print=True)
   
    create_xml_declaration(file_path_xml='bcsv_files/' + str(file_name))
    

def create_xml_declaration(file_path_xml):
    
    #doorzoek bimcollab versie op het systeem

    xml_declaration =   """<?xml version="1.0"?>
    <bimcollabsmartviewfile>
    <version>5</version>
    <applicationversion>Win - Version: 3.1 (build 3.1.13.217)</applicationversion>
</bimcollabsmartviewfile>
    """
    
    with open(file_path_xml, 'r') as xml_file:
        xml_data = xml_file.readlines()
        
        
    xml_data[0] = xml_declaration + '\n'
    
    with open(file_path_xml, 'w') as xml_file:
        xml_file.writelines(xml_data)
    
    
  
    
get_ifc_entities()


if entity == 'IfcWall':
    create_quantities_bcsv(file_name='wall_takeoff_' + str(tail) + '.bcsv', bouwdeel='opp. wanden ')

if entity == 'IfcSlab':
    create_quantities_bcsv(file_name='floor_takeoff_' + str(tail) + '.bcsv', bouwdeel='opp. vloeren ')

if entity == 'IfcCovering':
    create_quantities_bcsv(file_name='covering_takeoff_' + str(tail) + '.bcsv', bouwdeel='opp. plafonds en afwerkingen ')
    
    
    