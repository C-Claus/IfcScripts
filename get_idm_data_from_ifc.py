#coding=utf-8
#!/usr/bin/env python
#C.C.J. Claus & M. Berends
#2020 Coen Hagedoorn Bouwgroep

import os 
import uuid
import lxml 
import ifcopenshell 
import itertools 
import unicodedata

from datetime import datetime
from collections import defaultdict
from lxml import etree as ET
from operator import itemgetter
from collections import OrderedDict
from itertools import combinations


now = datetime.now()
year = now.strftime("%Y")
month = now.strftime("%m")
day = now.strftime("%d")


fname =  'U:\\00_ifc_modellen\\091 MaJa 20190604.ifc'

ifcfile = ifcopenshell.open(fname)
products = ifcfile.by_type('IfcProduct')
project = ifcfile.by_type('IfcProject')
site = ifcfile.by_type('IfcSite')


def get_software():
    print ('get software')
    
    for j in project:
        if j.OwnerHistory is not None:
            return (j.OwnerHistory.OwningApplication.ApplicationFullName)
               
    
def get_origins():
    print ('check origins')
    
    site = ifcfile.by_type('IfcSite')
    for j in site:
        print (j.RefLatitude,';', j.RefLongitude, ';',j.RefElevation)
      
   
          
def get_type_entities():
    print ('get type entities')
    
    product_type_list = []
 
    for product in products:  
        for relating_type in product.IsDefinedBy:
            if relating_type.is_a('IfcRelDefinesByType'):
                if relating_type.RelatingType.Name is not None:
                    product_type_list.append([product.is_a(), relating_type.RelatingType.Name])
                                   
    sorted_product_type_list = (sorted(product_type_list, key=itemgetter(0)))
    
    
    #Oplossing M. Berends
    pointer =  sorted_product_type_list[0][0]
    previous = pointer 
    content = []
    entity_list = []

    for i in sorted_product_type_list:
        pointer = i[0] 
        if pointer == previous:
            content.append(i[1])
        else: 
            entity_list.append([previous,content])
            previous=pointer
            content=[]
            content.append(i[1])
            
            
    entity_list.append([previous,content])   
    entity_dict = {}
    
    for i in entity_list:
        entity_dict[i[0]] = (set(i[1]))
      
    return entity_dict

    
def get_classification():
    print ('get classification')
    
    classification_name_list = []
    classification_list = [] 
    
    for product in products:
        for has_associations in product.HasAssociations:
            if has_associations.is_a('IfcRelAssociatesClassification'):
                if has_associations.RelatingClassification.Name is not None:
                   
                    classification_name_list.append(has_associations.Name)
                    classification_list.append([has_associations.Name, 
                                               str(has_associations.RelatingClassification.ItemReference) + ' ' +  str(has_associations.RelatingClassification.Name)
                                                ])

              
    classification_name_list = list(dict.fromkeys(classification_name_list))
    classification_list.sort()
    sorted_classification_list = list(classification_list for classification_list,_ in itertools.groupby(classification_list))
    
    if (len(sorted_classification_list)) == 0:
        sorted_classification_list.append(['Classification', 'No Classification Found'])
        
    return sorted_classification_list


def get_materials():
    print ('get materials')
    
    material_list = []
    
    for product in products:
        if product.HasAssociations:
            for j in product.HasAssociations:
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
      
   
        
    materials_list = list(dict.fromkeys(material_list))   
    
    return materials_list

def get_loadbearing():
    print ('get loadbearing')
    
    pset_list = []
    
    for product in products:
        for properties in product.IsDefinedBy:
            if properties.is_a('IfcRelDefinesByProperties'):
                if properties.RelatingPropertyDefinition:
                    if properties.RelatingPropertyDefinition.is_a('IfcPropertySet'):
                        if properties.RelatingPropertyDefinition.Name.startswith("Pset_"):
                            for loadbearing in properties.RelatingPropertyDefinition.HasProperties:
                                if loadbearing.Name == "LoadBearing":
                                    pset_list.append([properties.RelatingPropertyDefinition.Name, loadbearing.Name, loadbearing.NominalValue[0]])     
    
    loadbearing_list = [list(x) for x in set(tuple(x) for x in pset_list)]
    
    if len(loadbearing_list) == 0:
        loadbearing_list.append(['LoadBearing','No LoadBearing property found'])
    
    return loadbearing_list 

def get_isexternal():
    print ('get isexternal')
    
    pset_list = []
    for product in products:
        for properties in product.IsDefinedBy:
            if properties.is_a('IfcRelDefinesByProperties'):
                if properties.RelatingPropertyDefinition:
                    if properties.RelatingPropertyDefinition.is_a('IfcPropertySet'):
                        if properties.RelatingPropertyDefinition.Name.startswith("Pset_"):
                            for isexternal in properties.RelatingPropertyDefinition.HasProperties:
                                if isexternal.Name == "IsExternal":
                                    pset_list.append([properties.RelatingPropertyDefinition.Name, isexternal.Name, isexternal.NominalValue[0]])
        
    isexternal_list = [list(x) for x in set(tuple(x) for x in pset_list)]

    
    if len(isexternal_list) == 0:
        isexternal_list.append('IsExternal', 'No IsExternal property found')
        
    return isexternal_list

def get_firerating():
    
    pset_list = []
    
    for product in products:
        for properties in product.IsDefinedBy:
            if properties.is_a('IfcRelDefinesByProperties'):
                if properties.RelatingPropertyDefinition:
                    if properties.RelatingPropertyDefinition.is_a('IfcPropertySet'):
                        if properties.RelatingPropertyDefinition.Name.startswith("Pset_"):
                            for firerating in properties.RelatingPropertyDefinition.HasProperties:
                                if firerating.Name == "FireRating":
                                  
                                    pset_list.append([properties.RelatingPropertyDefinition.Name, firerating.Name, firerating.NominalValue[0]])
    

    fire_rating_list = [list(x) for x in set(tuple(x) for x in pset_list)]
    
    if len(fire_rating_list) == 0:
        fire_rating_list.append(['FireRating', 'No FireRating found'])
        
    return fire_rating_list

def get_phase():
    
    pset_phase_list =['Phasing','AC_Pset_RenovationAndPhasing','CPset_Phasing']
    phase_list = []
    phase_name_list = []
    
    for product in products:
        for properties in product.IsDefinedBy:
            if properties.is_a('IfcRelDefinesByProperties'):
                if properties.RelatingPropertyDefinition.is_a('IfcPropertySet'):

                    for phasing in properties.RelatingPropertyDefinition.HasProperties:
                        if properties.RelatingPropertyDefinition.Name is not None:
                            if (phasing.NominalValue):
                                
                                if properties.RelatingPropertyDefinition.Name in pset_phase_list:
                                    phase_name_list.append(properties.RelatingPropertyDefinition.Name)
                                    phase_list.append([properties.RelatingPropertyDefinition.Name, phasing.NominalValue[0]])
                
    phase_name_list = list(dict.fromkeys(phase_name_list))
    phase_list.sort()
    sorted_phase_list = list(phase_list for phase_list,_ in itertools.groupby(phase_list))
    
    if (len(sorted_phase_list)) == 0:
        sorted_phase_list.append(['Phase', 'No Phasing Found'])
        
    return sorted_phase_list

                         
def call_methods():
    
    print ('call methods')
    get_software()
    get_origins()
    get_type_entities().items():
    get_classification()
    get_phase()
    get_materials()
    get_loadbearing()
    get_isexternal()
    get_firerating()

       
call_methods()


    