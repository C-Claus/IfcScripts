import ifcopenshell 
import lxml 
from lxml import etree as ET
import os 

#import xml.etree.ElementTree
import uuid
from datetime import datetime
from collections import defaultdict

now = datetime.now()
year = now.strftime("%Y")
month = now.strftime("%m")
day = now.strftime("%d")

fname = 'K:\\Huizen\\Projecten\\1457 Renovatie Flat 11 te Zeist\\02 Modellen, tekeningen en rapporten\\01 agNOVA architecten\\IFC bestanden\\20191211_17520_Flat 11 Zeist_AC21.ifc'

ifcfile = ifcopenshell.open(fname)
products = ifcfile.by_type('IfcProduct')
project = ifcfile.by_type('IfcProject')

head, tail = os.path.split(fname)



def get_ifc_entities():
    
    get_wall_quantities(products)
     
def get_wall_quantities(products):
    
    print ("get wall quantities") 
    
    wall_type_list = []
    wall_type_area_list = []
    
    for product in products:
        if product.is_a("IfcWall"):
            for properties in product.IsDefinedBy:
                if properties.is_a('IfcRelDefinesByProperties'):
                    if properties.RelatingPropertyDefinition.is_a('IfcElementQuantity'):
                        for quantities in properties.RelatingPropertyDefinition.Quantities:
                            if quantities.Name == 'NetSideArea':
                                if product.IsDefinedBy:
                                    for j in product.IsDefinedBy:
                                        if j.is_a("IfcRelDefinesByType"):
                                            if j.RelatingType.Name:
                                                wall_type_list.append(j.RelatingType.Name)
                                                wall_type_area_list.append([j.RelatingType.Name, quantities.AreaValue])
                            

            
    result = defaultdict(int)
    result_aslist = defaultdict(list)
    
    for k, v in wall_type_area_list:
        result[k] += v
    
    result_aslist[k].append((v))
                  
    return result



def create_wall_bcsv(file_name, wanden):
    
    root = ET.Element('SMARTVIEWSETS')
    doc = ET.SubElement(root, "SMARTVIEWSET")
 
    title = ET.SubElement(doc, "TITLE").text = "CHO smartview " + wanden + tail
    description = ET.SubElement(doc, "DESCRIPTION").text = "Omschrijving volgt hier " 
    guid = ET.SubElement(doc, "GUID" ).text = str(uuid.uuid4())
    modification_date = ET.SubElement(doc, "MODIFICATIONDATE").text = str(datetime.timestamp(now))
    
    smartviews = ET.SubElement(doc, "SMARTVIEWS")
    
    for wall_type, value  in get_wall_quantities(products).items():
        print ('gaat dit goed?:', wall_type) 
        
        
        smartview = ET.SubElement(smartviews, "SMARTVIEW")
        
        smartview_title = ET.SubElement(smartview, "TITLE").text = str(wall_type) + ' (oppervlakte: ' + str(round(value, 2)) + 'm2)'
        smartview_description = ET.SubElement(smartview, "DESCRIPTION")
        smartview_creator = ET.SubElement(smartview, "CREATOR").text = "Coen Hagedoorn Bouwgroep"
        smartview_creation_date = ET.SubElement(smartview, "CREATIONDATE").text = day + " " + month + " " + year
        smartview_modifier = ET.SubElement(smartview, "MODIFIER").text = "Coen Hagedoorn Bouwgroep"
        smartview_modification_date = ET.SubElement(smartview, "MODIFICATIONDATE").text = day + " " + month + " " + year
        smartview_guid = ET.SubElement(smartview, "GUID").text = str(uuid.uuid4())
        
       
        rules = ET.SubElement(smartview, "RULES")
        rule = ET.SubElement(rules, "RULE")
        
        ifctype = ET.SubElement(rule, "IFCTYPE").text = "Wall"
        
        property = ET.SubElement(rule, "PROPERTY")
        property_name = ET.SubElement(property, "NAME").text= "Type"
        propertyset_name = ET.SubElement(property, "PROPERTYSETNAME").text = "Summary"
        property_type = ET.SubElement(property, "TYPE").text = "Summary"
        property_value_type = ET.SubElement(property, "VALUETYPE").text = "StringValue"
        property_unit = ET.SubElement(property, "UNIT").text = "None"
        
        
        condition = ET.SubElement(rule, "CONDITION")
        condition_type = ET.SubElement(condition, "TYPE").text = "Is"
        condition_value = ET.SubElement(condition, "VALUE").text = str(wall_type)
        
        action = ET.SubElement(rule, "ACTION")
        action_type = ET.SubElement(action, "TYPE").text = "Add"
        
        information_take_off = ET.SubElement(smartview, "INFORMATIONTAKEOFF")
        information_take_off_propertysetname = ET.SubElement(information_take_off, "PROPERTYSETNAME").text = str(wall_type)
        information_take_off_propertyname = ET.SubElement(information_take_off, "PROPERTYNAME").text = str(wall_type)
        information_take_off_operation = ET.SubElement(information_take_off, "OPERATION").text = "0"
        
    tree = ET.ElementTree(root)
    tree.write('bcsv_files/' + str(file_name), encoding="utf-8", xml_declaration=True, pretty_print=True)
   
    create_xml_declaration(file_path_xml='bcsv_files/' + str(file_name))
    

def create_xml_declaration(file_path_xml):

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
create_wall_bcsv(file_name='cho_wanden.bcsv', wanden='opp. wanden ')


