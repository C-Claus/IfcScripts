import ifcopenshell 

ifc_file = 'ifcfiles\\17411_KVN_fase1_SO.ifc'
ifcfile = ifcopenshell.open(ifc_file)
products = ifcfile.by_type("IfcProduct")

owner_history = ifcfile.by_type("IfcOwnerHistory")[0]

walls=[]
slabs=[]
coverings=[]
windows=[]
doors=[]
spaces=[]

for i in products:
    if i.is_a("IfcWall"):
        walls.append(i) 
        
    if i.is_a("IfcSlab"):
        slabs.append(i)
        
    if i.is_a("IfcCovering"):
        coverings.append(i)
        
    if i.is_a("IfcWindow"):
        windows.append(i)
        
    if i.is_a("IfcDoor"):
        doors.append(i)
              
    if i.is_a("IfcSpace"):
        spaces.apend(i)


property_values = [
    ifcfile.createIfcPropertySingleValue("CHB_informatiebron", "CHB_informatiebron", ifcfile.create_entity("IfcText", "ERP pakket"), None),
    ifcfile.createIfcPropertySingleValue("CHB_informatiebron_geraadpleegd", "CHB_informatiebron_geraadpleegd", ifcfile.create_entity("IfcText", "12 December 2018"), None),
    ifcfile.createIfcPropertySingleValue("CHB_BAG_ID", "CHB_BAG_ID", ifcfile.create_entity("IfcText", "0406100000007578"), None),
    ifcfile.createIfcPropertySingleValue("CHB_adres", "CHB_adres", ifcfile.create_entity("IfcText", "Adres in Rotterdam"), None),
    ifcfile.createIfcPropertySingleValue("CHB_ruimte", "CHB_ruimte", ifcfile.create_entity("IfcText", "het ruimtenummer"), None),
    ifcfile.createIfcPropertySingleValue("CHB_element", "CHB_element", ifcfile.create_entity("IfcText", "omschrijving van het element"), None),
    ifcfile.createIfcPropertySingleValue("CHB_gebrek", "CHB_gebrek", ifcfile.create_entity("IfcText", "omschrijving van het gebrek"), None),
    
]   

if len(walls) > 0:
    for wall in walls:
        property_set = ifcfile.createIfcPropertySet(wall.GlobalId, owner_history, "CHB_wand_eigenschappen", None, property_values)
        ifcfile.createIfcRelDefinesByProperties(wall.GlobalId, owner_history, None, None, [wall], property_set)

if len(slabs) > 0:
    for slab in slabs:
        property_set = ifcfile.createIfcPropertySet(slab.GlobalId, owner_history, "CHB_vloer_eigenschappen", None, property_values)
        ifcfile.createIfcRelDefinesByProperties(slab.GlobalId, owner_history, None, None, [slab], property_set)

if len(coverings) > 0:    
    for covering in coverings:
        property_set = ifcfile.createIfcPropertySet(slab.GlobalId, owner_history, "CHB_afwerking_eigenschappen", None, property_values)
        ifcfile.createIfcRelDefinesByProperties(slab.GlobalId, owner_history, None, None, [covering], property_set)

if len(windows) > 0:
    for window in windows:
        property_set = ifcfile.createIfcPropertySet(slab.GlobalId, owner_history, "CHB_kozijn_eigenschappen", None, property_values)
        ifcfile.createIfcRelDefinesByProperties(slab.GlobalId, owner_history, None, None, [window], property_set)

if len(doors) > 0:    
    for door in doors:
        property_set = ifcfile.createIfcPropertySet(slab.GlobalId, owner_history, "CHB_deur_eigenschappen", None, property_values)
        ifcfile.createIfcRelDefinesByProperties(slab.GlobalId, owner_history, None, None, [door], property_set)
  
if len(spaces) > 0:  
    for space in spaces:
        property_set = ifcfile.createIfcPropertySet(slab.GlobalId, owner_history, "CHB_ruimte_eigenschappen", None, property_values)
        ifcfile.createIfcRelDefinesByProperties(slab.GlobalId, owner_history, None, None, [space], property_set)
    
    

ifcfile.write(ifc_file)  




