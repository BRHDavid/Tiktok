import pandas as pd
import comtypes.client

#Conectar a etabs
def connect_to_etabs():
    ETABSobject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
    SapModel = ETABSobject.SapModel
    return SapModel, ETABSobject

#Insertar materiales al modelo
    #Obtener data
def data_user():
    data_material = pd.read_excel("Z:\ETABS-PYTHON\datos_material_geometria.xlsx", 
                              sheet_name= 'DATA_MATERIAL')
    
    name_concrete =data_material.loc[:,['NOMBRE CONCRETO','PROP. CONCRETO']]
    name_rebar =data_material.loc[:,['NOMBRE ACERO','PROP. ACERO']]
    
    return name_concrete, name_rebar
    
def insert_materials(SapModel,a,b):
    for row in a.values[:]:
        for i, item in enumerate(row):
            if i%2 == 0:
                SapModel.PropMaterial.SetMaterial(item,2)
                
    for row in b.values[0:]:
        for j, item in enumerate(row):
            if j%2 == 0:
                SapModel.PropMaterial.SetMaterial(item,6)

if __name__ == '__main__':
    SapModel, EtabsObject = connect_to_etabs()
    name_concrete, name_rebar = data_user()
    insert_materials(SapModel, name_concrete, name_rebar)
    







    
