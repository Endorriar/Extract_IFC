import ifcopenshell as ifcOs

model_path = 'ARC_Modell_NEST_230328.ifc'

ifc_file = ifcOs.open(model_path)

walls = ifc_file.by_type('IfcWall')

for w in walls:
    attributes = w.get_info()
    name = attributes.get('Name')
    print(attributes)
    print(name)