# Create and assign property set
import ifcopenshell 


ifc_file = 'ifcfiles\\17411_KVN_fase1_SO.ifc'



ifcfile = ifcopenshell.open(ifc_file)
products = ifcfile.by_type("IfcProduct")

owner_history = ifcfile.by_type("IfcOwnerHistory")[0]



walls=[]
for i in products:
    if i.is_a("IfcWall"):
        walls.append(i) 


#CHB_informatiebron, ERP pakket
#CHB_informatiebron_geraadpleegd, december 2018
#CHB_BAG_ID, 
#CHB_adres, 
#CHB_typologie,
#CHB_ruimte,
#CHB_ruimte,
#CHB_element,
#CHB_gebrek,


property_values = [
    ifcfile.createIfcPropertySingleValue("CHB_informatiebron", "CHB_informatiebron", ifcfile.create_entity("IfcText", "ERP pakket"), None),
    ifcfile.createIfcPropertySingleValue("CHB_informatiebron_geraadpleegd", "CHB_informatiebron_geraadpleegd", ifcfile.create_entity("IfcText", "12 December 2018"), None),
    ifcfile.createIfcPropertySingleValue("CHB_adres", "CHB_adres", ifcfile.create_entity("IfcText", "Ergens in Rotterdam"), None),
    
    ifcfile.createIfcPropertySingleValue("CHB_ruimte", "CHB_ruimte", ifcfile.create_entity("IfcText", "niet bekend"), None),
    ifcfile.createIfcPropertySingleValue("CHB_element", "CHB_element", ifcfile.create_entity("IfcText", "een wand"), None),
    ifcfile.createIfcPropertySingleValue("CHB_gebrek", "CHB_gebrek", ifcfile.create_entity("IfcText", "lekkage"), None),
    
    #ifcfile.createIfcPropertySingleValue("CHB_1", "CHB_1", ifcfile.create_entity("IfcText", "Describe the Reference"), None),
    #ifcfile.createIfcPropertySingleValue("CHB_2", "CHB_2", ifcfile.create_entity("IfcBoolean", True), None),
    #ifcfile.createIfcPropertySingleValue("CHB_3", "CHB_3", ifcfile.create_entity("IfcReal", 2.569), None),
    #ifcfile.createIfcPropertySingleValue("CHB_4", "CHB_4", ifcfile.create_entity("IfcInteger", 2), None)
]   


for wall in walls:
    property_set = ifcfile.createIfcPropertySet(wall.GlobalId, owner_history, "CHB_Propertyset", None, property_values)
    ifcfile.createIfcRelDefinesByProperties(wall.GlobalId, owner_history, None, None, [wall], property_set)



ifcfile.write(ifc_file)  




