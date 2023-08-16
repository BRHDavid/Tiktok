import pandas as pd
import comtypes.client

#Conectar a etabs
def connect_to_etabs():
    ETABSobject = comtypes.client.GetActiveObject(
                                "CSI.ETABS.API.ETABSObject")
    SapModel = ETABSobject.SapModel
    return SapModel

def data_user():
    data_material = pd.read_excel("datos_material_geometria.xlsx", 
                              sheet_name= 'DATA_VIGA')
    
    data_beam =data_material.loc[:,['VIGAS','V.BASE','V.ALTURA']]
    print(data_beam)
    return data_beam

def insert_geometry(SapModel,a):
    SapModel.SetPresentUnits(12)
    for row in a.values[:]:
        SapModel.PropFrame.SetRectangle(row[0],
                                        "CONCRETO 210",
                                        row[2],
                                        row[1])
        SapModel.PropFrame.SetRebarBeam(row[0],
                                        "ACERO 4200",
                                        "ACERO 4200",
                                        0.06,0.06,
                                        0,0,0,0)
        
        
if __name__ == '__main__':
    SapModel = connect_to_etabs()
    data_beam = data_user()
    insert_geometry(SapModel, data_beam)




