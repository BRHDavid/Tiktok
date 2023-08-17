import pandas as pd
import comtypes.client

#Conectar a etabs 

def connect_to_etabs():
    ETABSobject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
    SapModel = ETABSobject.SapModel
    return SapModel
#Obtener cortante V2 y momento M3 de la envolvente

def v22_m33_maximos(SapModel):
    SapModel.SetPresentUnits(12)
    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
    SapModel.Results.Setup.SetComboSelectedForOutput("Envolvente")
    data = SapModel.Results.FrameForce("1",3,0)
    beam_forces = pd.DataFrame(data[3:14],index= ["Name","Station","Case Type",
                                                  "Step Type","Step Num",'P','V2',
                                                  'V3','T','M2','M3',]).transpose()
    v2_m3 = beam_forces.groupby('Name')[['V2', 'M3']].apply(lambda x: x.abs().max())
    print(v2_m3)
    return v2_m3

if __name__ == "__main__":
    SapModel = connect_to_etabs()
    v22_m33_maximos(SapModel)
    
    
    
    
    
    