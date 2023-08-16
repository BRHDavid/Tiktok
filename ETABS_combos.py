import comtypes.client

#Conectar a etabs
def connect_to_etabs():
    ETABSobject = comtypes.client.GetActiveObject(
                "CSI.ETABS.API.ETABSObject")
    SapModel = ETABSobject.SapModel
    return SapModel

#Crear patrones de carga
def set_load_patterns(SapModel):
    SapModel.LoadPatterns.Add("SISMO X",5)
    SapModel.LoadPatterns.Add("SISMO Y",5)

#Crear combos
def combos(SapModel):
    lista = ["1.4CM+1.7CV",
             "1.25(CM+CV)+SX",
             "1.25(CM+CV)-SX",
             "1.25(CM+CV)+SY",
             "1.25(CM+CV)-SY",
             "0.9CM + SX",
             "0.9CM - SX",
             "0.9CM + SY",
             "0.9CM - SY",
             "ENVOLVENTE"
             ]
    lista2= ["SISMO X",
                "SISMO X",
                "SISMO Y",
                "SISMO Y"]
    lista3 = [1,
                -1,
                 1,
                -1]
    for i,y in enumerate(lista):
        if i ==0:
            SapModel.RespCombo.Add(y, 0)
            SapModel.RespCombo.SetCaseList(y, 0, "Dead", 1.4)
            SapModel.RespCombo.SetCaseList(y, 0, "Live", 1.7)
            
        elif i>0 and i<5:
            SapModel.RespCombo.Add(y, 0)
            SapModel.RespCombo.SetCaseList(y, 0, "Dead", 1.25)
            SapModel.RespCombo.SetCaseList(y, 0, "Live", 1.25)
            SapModel.RespCombo.SetCaseList(y, 0, lista2[i-1], lista3[i-1])

        elif i>=5 and i <9:
            SapModel.RespCombo.Add(y, 0)
            SapModel.RespCombo.SetCaseList(y, 0, "Dead", 0.9)
            SapModel.RespCombo.SetCaseList(y, 0, lista2[i-5], lista3[i-5])
        elif i ==9:
            SapModel.RespCombo.Add(y, 1)
            for i in lista:
                SapModel.RespCombo.SetCaseList(y, 1, i, 1)

if __name__ == '__main__':
    SapModel = connect_to_etabs()
    set_load_patterns(SapModel)
    combos(SapModel)
