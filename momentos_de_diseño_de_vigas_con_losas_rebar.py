import comtypes.client
import pandas as pd

#Conectar a etabs 
def connect_to_etabs():
    helper = comtypes.client.CreateObject('ETABSv1.Helper')
    helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
    ETABSObject = helper.GetObject("CSI.ETABS.API.ETABSObject")
    SapModel = ETABSObject.SapModel
    return SapModel

#Obtener momento M3 de la envolvente  
def m_abs_max(SapModel):
    SapModel.SetPresentUnits(12)
    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
    SapModel.Results.Setup.SetComboSelectedForOutput("Envolvente")
    data = SapModel.Results.FrameForce("",3,0)
    beam_forces = pd.DataFrame(data[1:14],index= ["Name",
                                                  "1",
                                                  "2",
                                                  "3",
                                                  "4",
                                                  'Step Type',
                                                  '6','P',
                                                  'V2','V3',
                                                  'T','M2','M3']).transpose()
    eliminar = ["1","2","3","4",'6','P','V2','V3','T','M2']
    beam_m3 = beam_forces.drop(columns=eliminar)
    names = beam_m3["Name"].unique()
    lista_momentos = {"M +":[],
                      "M -": []}
    for i in names:
        for j in range(len(beam_m3)):
            filtrado_neg = beam_m3[(beam_m3["Name"]==i)
                               & (beam_m3['Step Type'] == "Min")]
            filtrado_pos = beam_m3[(beam_m3["Name"]==i)
                               & (beam_m3['Step Type'] == "Max")]
        filas_neg = len(filtrado_neg)
        filas_pos = len(filtrado_pos)
        #Para los momentos negativos
        ini_neg = abs(filtrado_neg.iloc[0]["M3"])
        int_neg = abs(filtrado_neg.iloc[filas_neg//2]["M3"])
        end_neg = abs(filtrado_neg.iloc[filas_neg-1]["M3"])
        #Para los momentos positivos
        ini_pos = abs(filtrado_pos.iloc[0]["M3"])
        int_pos = abs(filtrado_pos.iloc[filas_pos//2]["M3"])
        end_pos = abs(filtrado_pos.iloc[filas_pos-1]["M3"])
        #Vigas de porticos especiales ACI 318-19
        ini_pos = max(ini_pos,ini_neg*0.5)
        end_pos = max(end_pos,end_neg*0.5)
        maximo = max(ini_neg,end_neg,ini_pos, end_pos)
        int_neg = max(int_neg,maximo*0.25)
        int_pos = max(int_pos,maximo*0.25)
        
        lista_momentos["M -"].extend([ini_neg,int_neg, end_neg])
        lista_momentos["M +"].extend([ini_pos,int_pos, end_pos])
    
    momentos_finales_df = pd.DataFrame(lista_momentos)
    return momentos_finales_df

#Obtener label
def get_geometry(SapModel):
    SapModel.SetPresentUnits(12)
    select = SapModel.SelectObj.GetSelected()[2]
    data = {
        "VIGA":[],
        "MATERIAL":[],
        "f'c (kg/cm2)":[],
        "BASE (cm)":[],
        "ALTURA (cm)":[]
    }
    for i in select:
        label = SapModel.FrameObj.GetLabelFromName(i)[0]
        section_label = SapModel.FrameObj.GetSection(i)[0]
        get_data = SapModel.PropFrame.GetRectangle(section_label)[1:4]
        f_c = SapModel.PropMaterial.GetOConcrete(get_data[0])[0]
        for j in range(3):
            data["VIGA"].append(label)
            data["MATERIAL"].append(get_data[0])
            data["f'c (kg/cm2)"].append(f_c/10)
            data["ALTURA (cm)"].append(get_data[1]*100)
            data["BASE (cm)"].append(get_data[2]*100)  
    data_df = pd.DataFrame(data)
    return data_df

def export_to_excel(a, b):
    try:
        tabla_final = a.join(b)
        tabla_final.to_excel("MOMENTOS_DE_DISEÑO.xlsx", index=False)
    except:
        print("No se pudo exportar a excel")
        
def main():
    SapModel = connect_to_etabs()
    momentos = m_abs_max(SapModel)
    geometria = get_geometry(SapModel)
    export_to_excel(geometria,momentos)
    
if __name__ == "__main__":
    main()
    
