import comtypes.client
import pandas as pd

#Datos del usuario
def data_user():
    diametro_de_barras = {'3/8': 0.9525, 
                          '1/2': 1.27, 
                          '5/8': 1.5875, 
                          '3/4': 1.905, 
                          '1': 2.54}
    d_long = input("Ingresa el diámetro de la barra longitudinal(plg): ")
    d_bastones ='5/8' #input("Ingresa el diámetro del baston(plg): ")
    d_estribo = input("Ingresa el diámetro del estribo(plg): ")
    recubrimiento = input("Ingresa el recubrimiento(cm): ")
    d_aceros = [
        diametro_de_barras[d_long],
        diametro_de_barras[d_bastones],
        diametro_de_barras[d_estribo],
        float(recubrimiento)
    ]
    return d_aceros

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
    beam_forces = pd.DataFrame(data[3:14],index= ["Name",
                                                  "Station",
                                                  "Case Type",
                                                  "Step Type",
                                                  "Step Num",
                                                  'P',
                                                  'V2','V3',
                                                  'T',
                                                  'M2','M3',]).transpose()
    eliminar = ["Station","Case Type","Step Num",'P','V2','V3','T','M2',]
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

#Obtener geometria
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

def calculate_area(a,b,aceros):
    tabla_global = a.join(b)
    peralte = tabla_global["ALTURA (cm)"]-aceros[3]- aceros[0]/2-aceros[2] 
    tabla_global["PERALTE"] = peralte
    
    A = peralte/5
    m_positivos = tabla_global["M +"]*100000
    m_negativos = tabla_global["M -"]*100000
    as_calculado = {"AS +":[],
                    "AS -":[]}
    for i in range(len(tabla_global)):
        dif = 1
        a = A[i]
        d = peralte[i]
        f_c_iter= tabla_global["f'c (kg/cm2)"][i]
        base = tabla_global["BASE (cm)"][i]
        acero_minimo = 14.1/4200*d*base
        #PARA LOS MOMENTOS POSITIVOS
        while dif !=0:
            ac_iteracion = m_positivos[i]/(0.9*4200*(d-a/2))
            k = (ac_iteracion*4200)/(0.85*f_c_iter*base)
            dif = (abs(k-a))
            a = k
        if ac_iteracion < acero_minimo:
            ac_iteracion = acero_minimo
        else:
            ac_iteracion = ac_iteracion
        as_calculado["AS +"].append(ac_iteracion)
        dif = 1
        #PARA LOS MOMENTOS NEGATIVOS
        while dif !=0:
            ac_iteracion = m_negativos[i]/(0.9*4200*(d-a/2))
            k = (ac_iteracion*4200)/(0.85*f_c_iter*base)
            dif = (abs(k-a))
            a = k
            ac_iteracion=round(ac_iteracion,4)
        if ac_iteracion < acero_minimo:
            ac_iteracion = acero_minimo
        else:
            ac_iteracion = ac_iteracion
        as_calculado["AS -"].append(ac_iteracion)
            
    tabla_global = tabla_global.join(pd.DataFrame(as_calculado))
    return tabla_global

def export_excel(tabla_global):
    
    try:
        tabla_global.to_excel("TABLA_RESUMEN.xlsx")
        """with pd.ExcelWriter("DATOS_TIKTOK.xlsx",mode="a"
                            ,if_sheet_exists="replace") as writer:
            tabla_global.to_excel(writer,sheet_name="DATOS")"""
        print("Se ha exportado la tabla de acero a excel")
    except:
        print("No se ha podido exportar la tabla de acero a excel")
        print("Verifica que el archivo excel no este abierto")

def main():
    aceros = data_user()
    SapModel = connect_to_etabs()
    datos_geometria = get_geometry(SapModel)
    momentos_extraidos = m_abs_max(SapModel)
    tabla_global = calculate_area(datos_geometria,momentos_extraidos,aceros )
    export_excel(tabla_global)

if __name__ == "__main__":
    main()
    
