Sub pasarBaseCatastro()
    Dim baseTramos, baseCatastro As String
    Dim pozoInicio, pozoFin, descarga, material, flujo, orientacion, filaPozoConsulta As String
    Dim diametro, altura As Double
    Dim colPozoITramos, colPozoFTramos, colDescargaTramos, colFlujoTramos, colOrientacionTramos As Integer
    Dim colHTramos, colDiamTramos, colMaterialTramos As Integer
    Dim colPozoConsulta, colFilaPozoConsulta As Integer
    Dim colHCatastroO, colHCatastroS, colHCatastroE, colHCatastroN As Integer
    Dim colDiamCatastroO, colDiamCatastroS, colDiamCatastroE, colDiamCatastroN As Integer
    Dim colMaterialCatastroO, colMaterialCatastroS, colMaterialCatastroE, colMaterialCatastroN As Integer
    Dim colFlujoCatastroO, colFlujoCatastroS, colFlujoCatastroE, colFlujoCatastroN As Integer
    Dim colNumPozoO, colNumPozoS, colNumPozoE, colNumPozoN As Integer
    Dim colDescargaCatastro, filaConsulta As Integer
    Dim filaInicio, numDatos As Integer
    
    baseTramos = "TramosPozos.xlsx"
    baseCatastro = "PozosCatastrados.xlsx"
    Windows(baseTramos).Activate
    filaInicio = ActiveCell.Row
    filaInicio = InputBox("Fila inicio", "INICIO", filaConsulta)
    numDatos = InputBox("Numero de datos", "DATOS", 1)
    colPozoITramos = 2: colPozoFTramos = 4: colDescargaTramos = 7: colFlujoTramos = 27: colOrientacionTramos = 34
    colHTramos = 8: colDiamTramos = 10: colMaterialTramos = 11
    colPozoConsulta = 89: colFilaPozoConsulta = 109
    colHCatastroO = 53: colHCatastroS = 61: colHCatastroE = 69: colHCatastroN = 77
    colDiamCatastroO = 54: colMaterialCatastroO = 55: colDiamCatastroS = 62: colMaterialCatastroS = 63
    colDiamCatastroE = 70: colMaterialCatastroE = 71: colDiamCatastroN = 78: colMaterialCatastroN = 79
    colFlujoCatastroO = 58: colFlujoCatastroS = 66: colFlujoCatastroE = 74: colFlujoCatastroN = 82
    colNumPozoO = 59: colNumPozoS = 67: colNumPozoE = 75: colNumPozoN = 83
    colDescargaCatastro = 6: filaConsulta = 3
    
    For i = 1 To numDatos
        pozoInicio = Cells(filaInicio, colPozoITramos).Value
        pozoFin = Cells(filaInicio, colPozoFTramos).Value
'        descarga = Cells(filaInicio, colDescargaTramos).Value
        altura = Cells(filaInicio, colHTramos).Value
        diametro = Cells(filaInicio, colDiamTramos).Value
        material = Cells(filaInicio, colMaterialTramos).Value
        flujo = Cells(filaInicio, colFlujoTramos).Value
        orientacion = Cells(filaInicio, colOrientacionTramos).Value
        Windows(baseCatastro).Activate
        Cells(filaConsulta, colPozoConsulta).Value = pozoInicio
        If Cells(filaConsulta, 90).Text = "#N/A" Then
            GoTo Error
        Else
            filaPozoConsulta = Cells(filaConsulta, colFilaPozoConsulta).Value
            Cells(filaPozoConsulta, 1).Select
            Select Case orientacion
                Case "O"
                    Cells(filaPozoConsulta, colHCatastroO).Value = altura
                    Cells(filaPozoConsulta, colDiamCatastroO).Value = diametro
                    Cells(filaPozoConsulta, colMaterialCatastroO).Value = material
                    Cells(filaPozoConsulta, colNumPozoO).Value = pozoFin
                    Cells(filaPozoConsulta, colFlujoCatastroO).Value = flujo
                Case "S"
                    Cells(filaPozoConsulta, colHCatastroS).Value = altura
                    Cells(filaPozoConsulta, colDiamCatastroS).Value = diametro
                    Cells(filaPozoConsulta, colMaterialCatastroS).Value = material
                    Cells(filaPozoConsulta, colNumPozoS).Value = pozoFin
                    Cells(filaPozoConsulta, colFlujoCatastroS).Value = flujo
                Case "E"
                    Cells(filaPozoConsulta, colHCatastroE).Value = altura
                    Cells(filaPozoConsulta, colDiamCatastroE).Value = diametro
                    Cells(filaPozoConsulta, colMaterialCatastroE).Value = material
                    Cells(filaPozoConsulta, colNumPozoE).Value = pozoFin
                    Cells(filaPozoConsulta, colFlujoCatastroE).Value = flujo
                Case "N"
                    Cells(filaPozoConsulta, colHCatastroN).Value = altura
                    Cells(filaPozoConsulta, colDiamCatastroN).Value = diametro
                    Cells(filaPozoConsulta, colMaterialCatastroN).Value = material
                    Cells(filaPozoConsulta, colNumPozoN).Value = pozoFin
                    Cells(filaPozoConsulta, colFlujoCatastroN).Value = flujo
            End Select
            Cells(filaPozoConsulta, 3).Interior.ColorIndex = 4
        End If
Error:
        filaInicio = filaInicio + 1
        Windows(baseTramos).Activate
        Cells(filaInicio, 1).Select
    Next i
End Sub