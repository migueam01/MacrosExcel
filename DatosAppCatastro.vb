Sub pasarBaseApp()
    'Pasa los datos de Excel (datos en hoja formato de catastro) a los archivos de texto
    'usados para la aplicación catastro de pozos.
    'Pasa los datos directamente desde la base en Excel.
    
    Dim matrizDatos(0 To 50, 0 To 50) As Variant
    Dim OpFuncionaContiguo, OpEstadoContiguo, funcionaContiguo, estadoContiguo As String
    Dim existeSum, existeRejilla, rejillaHf, rejillaHs, rejillaMet As String
    Dim tipoRejilla, estadoSumidero, sumB, sumR, sumM As String
    Dim nivelSobre, nivelado, nivelEnterrado, nivel As String
    Dim nivelCalzadaSobre, nivelCalzadaBajo, nivelCalzada As Variant
    Dim opcionFunciona, funciona, verificarNumero As String
    Dim pozoB, pozoM, pozoR, estadoPozo, sanitario, pluvial, comb, fluido As String
    Dim escSi, escNo, escalera, escal, escalines, tipoEsc, interior, parcial, sinEnlucir, enlucidos As String
    Dim limpio, tierra, aguaEst, mantenimiento, tapaHf, tapaHs, tapaNo, tapaPiedra, tapa As String
    Dim opCadena, cadena, opRota, rota, nombrePozoCentral As String
    Dim flujo As String
    Dim rango As String
    Dim diametro, alturaBase, alturaCorona As Double
    Dim filaInicio As Integer
    
    Dim rutaPozos, rutaPozosTx As String
    Dim objetoFSO As Object
    Dim archivoTexto As Object
    Set objetoFSO = CreateObject("Scripting.FileSystemObject")
    
    filaInicio = ActiveCell.Row
    filaInicio = InputBox("Fila inicio", "INICIO", filaInicio)
'    rutaPozos = Range("B1").Value + "\"
'    rutaPozosTx = Range("C1").Value + "\"
    rutaPozosTx = "C:\PozosNuevos\"
'    Range("B2").Select
    archivo = ActiveCell.Value
'    Workbooks.OpenText Filename:=rutaPozos + archivo
'    numCaracter = InStrRev(archivo, "\")
'    archivoExcelPozo = Mid(archivo, numCaracter + 1)
    archivoMacro = "MacroCatastro.xlsm"
    
    Do While archivo <> ""
    
'        Windows(archivoExcelPozo).Activate
        rango = "CB" & filaInicio
        OpEstadoContiguo = Range(rango).Value
        Select Case OpEstadoContiguo
            Case "B"
                estadoContiguo = "Bueno"
            Case "R"
                estadoContiguo = "Regular"
            Case "M"
                estadoContiguo = "Malo"
        End Select
        
        rango = "CC" & filaInicio
        OpFuncionaContiguo = Range(rango).Value
        Select Case OpFuncionaContiguo
            Case "Si"
                funcionaContiguo = "Si"
            Case "No"
                funcionaContiguo = "No"
        End Select
        
        rango = "BY" & filaInicio
        verificarNumero = Range(rango).Value
        If verificarNumero = "" Then
            matrizDatos(1, 4) = 0
        Else
            alturaBase = Range(rango).Value
            matrizDatos(1, 4) = alturaBase
        End If
        
        'Calcula la altura de la corona
        rango = "BZ" & filaInicio
        verificarNumero = Range(rango)  'Obtiene el diámetro
        If verificarNumero = "" Then
            matrizDatos(1, 5) = 0
        Else
            diametro = Range(rango) / 1000
            alturaCorona = alturaBase - diametro
            matrizDatos(1, 5) = alturaCorona
        End If
        
        rango = "CD" & filaInicio
        flujo = Range(rango).Value
        Select Case flujo
            Case "E"
                matrizDatos(1, 9) = "Entra"
            Case "S"
                matrizDatos(1, 9) = "Sale"
            Case "I"
                matrizDatos(1, 9) = "Inicio"
        End Select
        
        rango = "CE" & filaInicio  'Nombre del pozo norte
        matrizDatos(1, 1) = 1: matrizDatos(1, 2) = Range(rango).Value: matrizDatos(1, 3) = "Norte":
        rango = "CA" & filaInicio  'Tipo de tubería
        matrizDatos(1, 6) = Range(rango).Value: matrizDatos(1, 7) = estadoContiguo: matrizDatos(1, 8) = funcionaContiguo
        
        rango = "BL" & filaInicio
        OpEstadoContiguo = Range(rango).Value
        Select Case OpEstadoContiguo
            Case "B"
                estadoContiguo = "Bueno"
            Case "R"
                estadoContiguo = "Regular"
            Case "M"
                estadoContiguo = "Malo"
        End Select
        
        rango = "BM" & filaInicio
        OpFuncionaContiguo = Range(rango).Value
        Select Case OpFuncionaContiguo
            Case "Si"
                funcionaContiguo = "Si"
            Case "No"
                funcionaContiguo = "No"
        End Select
        
        rango = "BI" & filaInicio
        verificarNumero = Range(rango).Value
        If verificarNumero = "" Then
            matrizDatos(2, 4) = 0
        Else
            alturaBase = Range(rango).Value
            matrizDatos(2, 4) = alturaBase
        End If
        
        'Calcula la altura de la corona del pozo sur
        rango = "BJ" & filaInicio
        verificarNumero = Range(rango).Value  'diámetro de la tubería
        If verificarNumero = "" Then
            matrizDatos(2, 5) = 0
        Else
            diametro = Range(rango) / 1000
            alturaCorona = alturaBase - diametro
            matrizDatos(2, 5) = alturaCorona
        End If
        
        rango = "BN" & filaInicio
        flujo = Range(rango).Value
        Select Case flujo
            Case "E"
                matrizDatos(2, 9) = "Entra"
            Case "S"
                matrizDatos(2, 9) = "Sale"
            Case "I"
                matrizDatos(2, 9) = "Inicio"
        End Select
        
        rango = "BO" & filaInicio
        matrizDatos(2, 1) = 2: matrizDatos(2, 2) = Range(rango).Value: matrizDatos(2, 3) = "Sur"
        rango = "BK" & filaInicio
        matrizDatos(2, 6) = Range(rango).Value: matrizDatos(2, 7) = estadoContiguo: matrizDatos(2, 8) = funcionaContiguo
        
        rango = "BT" & filaInicio
        OpEstadoContiguo = Range(rango).Value
        Select Case OpEstadoContiguo
            Case "B"
                estadoContiguo = "Bueno"
            Case "R"
                estadoContiguo = "Regular"
            Case "M"
                estadoContiguo = "Malo"
        End Select
        
        rango = "BU" & filaInicio
        OpFuncionaContiguo = Range(rango).Value
        Select Case OpFuncionaContiguo
            Case "Si"
                funcionaContiguo = "Si"
            Case "No"
                funcionaContiguo = "No"
        End Select
        
        rango = "BQ" & filaInicio
        verificarNumero = Range(rango).Value
        If verificarNumero = "" Then
            matrizDatos(3, 4) = 0
        Else
            alturaBase = Range(rango).Value
            matrizDatos(3, 4) = verificarNumero
        End If
        
        'Calcula la altura de la corona del pozo Este
        rango = "BR" & filaInicio
        verificarNumero = Range(rango).Value
        If verificarNumero = "" Then
            matrizDatos(3, 5) = 0
        Else
            diametro = Range(rango) / 1000
            alturaCorona = alturaBase - diametro
            matrizDatos(3, 5) = alturaCorona
        End If
        
        rango = "BV" & filaInicio
        flujo = Range(rango).Value
        Select Case flujo
            Case "E"
                matrizDatos(3, 9) = "Entra"
            Case "S"
                matrizDatos(3, 9) = "Sale"
            Case "I"
                matrizDatos(3, 9) = "Inicio"
        End Select
        
        rango = "BW" & filaInicio
        matrizDatos(3, 1) = 3: matrizDatos(3, 2) = Range(rango).Value: matrizDatos(3, 3) = "Este"
        rango = "BS" & filaInicio
        matrizDatos(3, 6) = Range(rango).Value: matrizDatos(3, 7) = estadoContiguo: matrizDatos(3, 8) = funcionaContiguo
        
        rango = "BD" & filaInicio
        OpEstadoContiguo = Range(rango).Value
        Select Case OpEstadoContiguo
            Case "B"
                estadoContiguo = "Bueno"
            Case "R"
                estadoContiguo = "Regular"
            Case "M"
                estadoContiguo = "Malo"
        End Select
        
        rango = "BE" & filaInicio
        OpFuncionaContiguo = Range(rango).Value
        Select Case OpFuncionaContiguo
            Case "Si"
                funcionaContiguo = "Si"
            Case "No"
                funcionaContiguo = "No"
        End Select
        
        rango = "BA" & filaInicio
        verificarNumero = Range(rango).Value
        If verificarNumero = "" Then
            matrizDatos(4, 4) = 0
        Else
            alturaBase = Range(rango).Value
            matrizDatos(4, 4) = alturaBase
        End If
        
        'Calcula la altura de la corona del pozo oeste
        rango = "BB" & filaInicio
        verificarNumero = Range(rango).Value  'Diámetro
        If verificarNumero = "" Then
            matrizDatos(4, 5) = 0
        Else
            diametro = Range(rango) / 1000
            alturaCorona = alturaBase - diametro
            matrizDatos(4, 5) = alturaCorona
        End If
        
        rango = "BF" & filaInicio
        flujo = Range(rango).Value
        Select Case flujo
            Case "E"
                matrizDatos(4, 9) = "Entra"
            Case "S"
                matrizDatos(4, 9) = "Sale"
            Case "I"
                matrizDatos(4, 9) = "Inicio"
        End Select
        
        rango = "BG" & filaInicio
        matrizDatos(4, 1) = 4: matrizDatos(4, 2) = Range(rango).Value: matrizDatos(4, 3) = "Oeste"
        rango = "BC" & filaInicio
        matrizDatos(4, 6) = Range(rango).Value: matrizDatos(4, 7) = estadoContiguo: matrizDatos(4, 8) = funcionaContiguo
        
        'sumideros
        existeSum = ""  'Range("R15").Value
        existeRejilla = ""  'Range("R20").Value
        existeSum = ""
        If existeSum = "X" Then
            existeSum = "Si"
        Else
            existeSum = "No"
        End If
        If existeRejilla = "X" Then
            existeRejilla = "Si"
        Else
            existeRejilla = "No"
        End If
        
        tipoRejilla = ""
        rejillaHf = "X"  'Range("R22").Value
        rejillaHs = ""  'Range("T22").Value
        rejillaMet = ""  'Range("R24").Value
        If rejillaHf = "X" Then
            tipoRejilla = "HF"
        End If
        If rejillaHs = "X" Then
            tipoRejilla = "HS"
        End If
        If rejillaMet = "X" Then
            tipoRejilla = "MET"
        End If
        
        estadoSumidero = ""
        sumB = "X"  'Range("T27").Value
        sumM = ""  'Range("T29").Value
        sumR = ""  'Range("T28").Value
        If sumB = "X" Then
            estadoSumidero = "Bueno"
        End If
        If sumM = "X" Then
            estadoSumidero = "Malo"
        End If
        If sumR = "X" Then
            estadoSumidero = "Regular"
        End If
        
        opcionFunciona = "X"  'Range("R31").Value
        funciona = ""
        If opcionFunciona = "X" Then
            funciona = "Si"
        Else
            funciona = "No"
        End If
        
        verificarNumero = "0"  'Range("S17").Value
        If verificarNumero = "" Then
            matrizDatos(5, 3) = 0
        Else
            matrizDatos(5, 3) = verificarNumero
        End If
        
        verificarNumero = "0"  'Range("T25").Value
        If verificarNumero = "" Then
            matrizDatos(5, 6) = 0
        Else
            matrizDatos(5, 6) = verificarNumero
        End If
        
        matrizDatos(5, 1) = 5: matrizDatos(5, 2) = existeSum: matrizDatos(5, 4) = existeRejilla: matrizDatos(5, 5) = tipoRejilla
        matrizDatos(5, 7) = estadoSumidero: matrizDatos(5, 8) = funciona
        
        nivelSobre = ""  'Range("X35").Value
        nivelado = "X"  'Range("X37").Value
        nivelEnterrado = ""  'Range("X39").Value
        nivel = ""
        If nivelSobre = "X" Then
            nivel = "Sobresalido"
        End If
        If nivelado = "X" Then
            nivel = "Nivelado"
        End If
        If nivelEnterrado = "X" Then
            nivel = "Enterrado"
        End If
        
        rango = "CG" & filaInicio
        matrizDatos(6, 1) = 6: matrizDatos(6, 2) = Range(rango).Value: matrizDatos(6, 3) = nivel: matrizDatos(6, 4) = 0
        
        pozoB = "X"  'Range("AI18").Value
        pozoR = "" 'Range("AI19").Value
        pozoM = ""  'Range("AI20").Value
        estadoPozo = ""
        If pozoB = "X" Then
            estadoPozo = "Bueno"
        End If
        If pozoM = "X" Then
            estadoPozo = "Malo"
        End If
        If pozoR = "X" Then
            estadoPozo = "Regular"
        End If
        
        sanitario = ""  'Range("AM18").Value
        pluvial = ""  'Range("AM19").Value
        comb = "X"  'Range("AM20").Value
        fluido = ""
        If sanitario = "X" Then
            fluido = "Sanitario"
        End If
        If pluvial = "X" Then
            fluido = "Pluvial"
        End If
        If comb = "X" Then
            fluido = "Combinado"
        End If
        
        verificarNumero = "2019"  'Range("AL9").Value
        If verificarNumero = "" Then
            matrizDatos(7, 4) = 0
        Else
            matrizDatos(7, 4) = verificarNumero
        End If
        
        matrizDatos(7, 1) = 7: matrizDatos(7, 2) = "Consultora": matrizDatos(7, 3) = "14/08/2019": matrizDatos(7, 5) = estadoPozo: matrizDatos(7, 6) = fluido
        
        escNo = ""  'Range("AM33").Value
        escSi = "X"  'Range("AM32").Value
        escalera = ""
        If escNo = "X" Then
            escalera = "No"
        End If
        If escSi = "X" Then
            escalera = "Si"
        End If
        
        escal = "X"  'Range("AM34").Value
        escalines = "" 'Range("AM35").Value
        tipoEsc = ""
        If escal = "X" Then
            tipoEsc = "Escalera"
        End If
        If escalines = "X" Then
            tipoEsc = "Escalines"
        End If
        
        interior = "X"  'Range("AM27").Value
        parcial = ""  'Range("AM29").Value
        sinEnlucir = ""  'Range("AM28").Value
        enlucidos = ""
        If interior = "X" Then
            enlucidos = "Interior"
        End If
        If parcial = "X" Then
            enlucidos = "Parcial"
        End If
        If sinEnlucir = "X" Then
            enlucidos = "Sin enlucir"
        End If
        
        verificarNumero = "1"  'Range("AM36").Value
        If verificarNumero = "" Then
            matrizDatos(8, 7) = 0
        Else
            matrizDatos(8, 7) = verificarNumero
        End If
        
        verificarNumero = "0.40"  'Range("AM37").Value
        If verificarNumero = "" Then
            matrizDatos(8, 8) = 0
        Else
            matrizDatos(8, 8) = verificarNumero
        End If
        
        matrizDatos(8, 1) = 8: matrizDatos(8, 2) = "HS": matrizDatos(8, 3) = "HS": matrizDatos(8, 4) = "HS": matrizDatos(8, 5) = escalera
        matrizDatos(8, 6) = tipoEsc: matrizDatos(8, 9) = enlucidos
        
        limpio = "X"  'Range("AI32").Value
        tierra = ""  'Range("AI33").Value
        aguaEst = ""  'Range("AI34").Value
        mantenimiento = ""
        If limpio = "X" Then
            mantenimiento = "Limpio"
        End If
        If tierra = "X" Then
            mantenimiento = "Tierra o piedra"
        End If
        If aguaEst = "X" Then
            mantenimiento = "Agua estancada"
        End If
        
        tapaHf = "X"  'Range("AI39").Value
        tapaHs = ""  'Range("AI40").Value
        tapaNo = ""  'Range("AI41").Value
        tapaPiedra = ""  'Range("AM39").Value
        tapa = ""
        If tapaHf = "X" Then
            tapa = "HF"
        End If
        If tapaHs = "X" Then
            tapa = "HS"
        End If
        If tapaPiedra = "X" Then
            tapa = "PIED"
        End If
        If tapaNo = "X" Then
            tapa = "No tiene"
        End If
        
        opCadena = "X"  'Range("AI43").Value
        cadena = ""
        If opCadena = "X" Then
            cadena = "Si"
        Else
            cadena = "No"
        End If
        
        opRota = ""  'Range("AI44").Value
        rota = ""
        If opRota = "X" Then
            rota = "Si"
        Else
            rota = "No"
        End If
        
        matrizDatos(9, 1) = 9: matrizDatos(9, 2) = mantenimiento: matrizDatos(9, 3) = tapa: matrizDatos(9, 4) = cadena: matrizDatos(9, 5) = rota
        
        verificarNumero = "0.60"  'Range("AK46").Value
        If verificarNumero = "" Then
            matrizDatos(10, 2) = 0
        Else
            matrizDatos(10, 2) = verificarNumero
        End If
        
        rango = "AW" & filaInicio
        verificarNumero = Range(rango).Value
        If verificarNumero = "" Then
            matrizDatos(10, 3) = 0
        Else
            matrizDatos(10, 3) = Round(Range(rango).Value, 2)
        End If
        
        verificarNumero = "0.90"  'Range("AK48").Value
        If verificarNumero = "" Then
            matrizDatos(10, 4) = 0
        Else
            matrizDatos(10, 4) = verificarNumero
        End If
        
        verificarNumero = "0.60"  'Range("AM47").Value
        If verificarNumero = "" Then
            matrizDatos(10, 5) = 0
        Else
            matrizDatos(10, 5) = verificarNumero
        End If
        
        matrizDatos(10, 1) = 10: matrizDatos(10, 6) = "Tramo nuevo"
        
        matrizDatos(11, 1) = 11: matrizDatos(11, 2) = "OE": matrizDatos(11, 3) = "NS"
        
        rango = "B" & filaInicio
        verificarNumero = Range(rango).Value
        If verificarNumero = "" Then
            matrizDatos(12, 5) = 0
        Else
            matrizDatos(12, 5) = verificarNumero
        End If
        
        rango = "F" & filaInicio  'Descarga
        matrizDatos(12, 1) = 12: matrizDatos(12, 2) = "COM": matrizDatos(12, 3) = "SN": matrizDatos(12, 4) = Range(rango).Value: matrizDatos(12, 6) = "": matrizDatos(12, 7) = ""
        
        'coordenada Este
        rango = "R" & filaInicio
        verificarNumero = Range(rango).Value
        If verificarNumero = "" Then
            matrizDatos(13, 3) = 0
        Else
            verificarNumero = Round(verificarNumero, 3)
            matrizDatos(13, 3) = verificarNumero
        End If
        
        'coordenada norte
        rango = "S" & filaInicio
        verificarNumero = Range(rango).Value
        If verificarNumero = "" Then
            matrizDatos(13, 4) = 0
        Else
            verificarNumero = Round(verificarNumero, 3)
            matrizDatos(13, 4) = verificarNumero
        End If
        
        'cota
        rango = "T" & filaInicio
        verificarNumero = Range(rango).Value
        If verificarNumero = "" Then
            matrizDatos(13, 5) = 0
        Else
            verificarNumero = Round(verificarNumero, 3)
            matrizDatos(13, 5) = verificarNumero
        End If
        
        matrizDatos(13, 1) = 13
        matrizDatos(13, 2) = "False"
        
        'Creamos un archivo con el método CreateTextFile
        rango = "A" & filaInicio
        nombrePozoCentral = Range(rango).Value
        Set archivoTexto = objetoFSO.CreateTextFile(rutaPozosTx & nombrePozoCentral & ".txt", True)
        For i = 1 To 13
            Select Case i
                Case 1 To 4
                    For j = 1 To 9
                        archivoTexto.Write matrizDatos(i, j) & "; "
                    Next j
                        archivoTexto.Write "0.00"
                        archivoTexto.Writeline
                Case 5
                    For j = 1 To 8
                        archivoTexto.Write matrizDatos(i, j) & "; "
                    Next j
                    archivoTexto.Writeline
                Case 6
                    For j = 1 To 4
                        archivoTexto.Write matrizDatos(i, j) & "; "
                    Next j
                    archivoTexto.Writeline
                Case 7
                    For j = 1 To 6
                        archivoTexto.Write matrizDatos(i, j) & "; "
                    Next j
                    archivoTexto.Writeline
                Case 8
                    For j = 1 To 9
                        archivoTexto.Write matrizDatos(i, j) & "; "
                    Next j
                    archivoTexto.Writeline
                Case 9
                    For j = 1 To 5
                        archivoTexto.Write matrizDatos(i, j) & "; "
                    Next j
                    archivoTexto.Writeline
                Case 10
                    For j = 1 To 6
                        archivoTexto.Write matrizDatos(i, j) & "; "
                    Next j
                    archivoTexto.Writeline
                Case 11
                    For j = 1 To 3
                        archivoTexto.Write matrizDatos(i, j) & "; "
                    Next j
                    archivoTexto.Writeline
                Case 12
                    For j = 1 To 7
                        archivoTexto.Write matrizDatos(i, j) & "; "
                    Next j
                    archivoTexto.Writeline
                Case 13
                    For j = 1 To 5
                        archivoTexto.Write matrizDatos(i, j) & "; "
                    Next j
                    archivoTexto.Writeline
            End Select
        Next i
        archivoTexto.Close
'        ActiveWorkbook.Save
'        ActiveWorkbook.Close
'        Windows(archivoMacro).Activate
        ActiveCell.Offset(1, 0).Select
        filaInicio = filaInicio + 1
        archivo = ActiveCell.Value
'        If archivo <> "" Then
'            Workbooks.OpenText Filename:=rutaPozos + archivo
'            numCaracter = InStrRev(archivo, "\")
'            archivoExcelPozo = Mid(archivo, numCaracter + 1)
'        End If
    Loop
End Sub

Function tipoTuberiaApp(ByVal material As String) As String
    Dim tipoMaterialApp As String
    Select Case material
    
    End Select
End Function