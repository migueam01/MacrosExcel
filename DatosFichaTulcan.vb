Sub colocarDatosFichaTulcan()
    Dim celdasBase(0 To 90) As Variant
    Dim celdasFicha(0 To 90) As Variant
    Dim datoLeido(0 To 90) As Variant
    
    Dim alturaCoronaO, alturaCoronaS, alturaCoronaE, alturaCoronaN, alturaBase, diametro As Double
        
    'orden de los datos:
    '2Número de hoja, 4fecha, 5año, 6descarga
    '7sistema, 8sector, 9cantidad sumideros, 18norte, 19este, 20cota
    '21estadoBueno, 22estado regular, 23estado malo, 27material pared, 28material zocalo, 29material fondo
    '30limpio, 31con tierra, 32agua estancada, 33enlucido interior, 34sin enlucir, 35parcial, 36escalera si, 37no
    '40cantidad escaleras, 41separación, 42tapa HF, 43hormigón, 44tapa piedra, 45sin tapa
    '46cadena si, 47cadena no, 48tapa, 49alto, 50ancho, 51altura HC+, 52altura HC-
    'altura base, diámetro, material, estado
    'funciona, flujo, pozo (oeste, sur, este, norte)
    '85tipo de calzada, 87tapado
    celdasBase(0) = 2: celdasBase(1) = 4: celdasBase(2) = 5: celdasBase(3) = 6
    celdasBase(4) = 7: celdasBase(5) = 8: celdasBase(6) = 9: celdasBase(7) = 18
    celdasBase(8) = 19: celdasBase(9) = 20: celdasBase(10) = 21: celdasBase(11) = 22
    celdasBase(12) = 23: celdasBase(13) = 27: celdasBase(14) = 28: celdasBase(15) = 29
    celdasBase(16) = 30: celdasBase(17) = 31: celdasBase(18) = 32: celdasBase(19) = 33
    celdasBase(20) = 34: celdasBase(21) = 35: celdasBase(22) = 36: celdasBase(23) = 37
    celdasBase(24) = 40: celdasBase(25) = 41: celdasBase(26) = 42: celdasBase(27) = 43
    celdasBase(28) = 44: celdasBase(29) = 45: celdasBase(30) = 46: celdasBase(31) = 47
    celdasBase(32) = 48: celdasBase(33) = 49: celdasBase(34) = 50: celdasBase(35) = 51
    celdasBase(36) = 52: celdasBase(37) = 53: celdasBase(38) = 54
    celdasBase(39) = 55: celdasBase(40) = 56: celdasBase(41) = 57: celdasBase(42) = 58
    celdasBase(43) = 59: celdasBase(44) = 60: celdasBase(45) = 61: celdasBase(46) = 62
    celdasBase(47) = 63: celdasBase(48) = 64: celdasBase(49) = 65: celdasBase(50) = 66
    celdasBase(51) = 67: celdasBase(52) = 68: celdasBase(53) = 69: celdasBase(54) = 70
    celdasBase(55) = 71: celdasBase(56) = 72: celdasBase(57) = 73: celdasBase(58) = 74
    celdasBase(59) = 75: celdasBase(60) = 76: celdasBase(61) = 77: celdasBase(62) = 78
    celdasBase(63) = 79: celdasBase(64) = 80: celdasBase(65) = 81: celdasBase(66) = 82
    celdasBase(67) = 83: celdasBase(68) = 84: celdasBase(69) = 85: celdasBase(70) = 87
    
    'AH5Número de hoja, AH9fecha, AL9año, AE4descarga
    'V3sistema, S4sector, S17cantidad sumideros, AI14norte, AI13este, AL11cota
    'AI18estadoBueno, AI19estado regular, AI20estado malo, AI27material pared, AI28material zocalo, AI29material fondo
    'AI32limpio, AI33con tierra, AI34agua estancada, AM27enlucido interior, AM28sin enlucir, AM29parcial
    'AM32escalera si, AM33no, AM36cantidad escaleras, AM37separación, AI39tapa HF, AI40hormigón
    'AM39tapa piedra, AI41sin tapa
    'AI43cadena si, AL43cadena no, AK46tapa, AK47alto, AK48ancho, AJ50altura HC+, AJ51altura HC-
    'altura base, diámetro, material, estado
    'funciona, flujo, pozo (oeste, sur, este, norte)
    '85tipo de calzada, 87tapado
    celdasFicha(0) = "AH5": celdasFicha(1) = "AH9": celdasFicha(2) = "AL9": celdasFicha(3) = "AE4"
    celdasFicha(4) = "V3": celdasFicha(5) = "S4": celdasFicha(6) = "S17": celdasFicha(7) = "AI14"
    celdasFicha(8) = "AI13": celdasFicha(9) = "AL11": celdasFicha(10) = "AI18"
    celdasFicha(11) = "AI19": celdasFicha(12) = "AI20": celdasFicha(13) = "AI27"
    celdasFicha(14) = "AI28": celdasFicha(15) = "AI29": celdasFicha(16) = "AI32"
    celdasFicha(17) = "AI33": celdasFicha(18) = "AI34": celdasFicha(19) = "AM27"
    celdasFicha(20) = "AM28": celdasFicha(21) = "AM29"
    celdasFicha(22) = "AM32": celdasFicha(23) = "AM33": celdasFicha(24) = "AM36"
    celdasFicha(25) = "AM37": celdasFicha(26) = "AI39"
    celdasFicha(27) = "AI40": celdasFicha(28) = "AM39": celdasFicha(29) = "AI41"
    celdasFicha(30) = "AI43": celdasFicha(31) = "AL43"
    celdasFicha(32) = "AK46": celdasFicha(33) = "AK47": celdasFicha(34) = "AK48": celdasFicha(35) = "AJ50"
    celdasFicha(36) = "AJ51"
    celdasFicha(37) = "E40": celdasFicha(38) = "F41": celdasFicha(39) = "E42"
    celdasFicha(40) = "E43": celdasFicha(41) = "E44": celdasFicha(42) = "J43": celdasFicha(43) = "G43"
    celdasFicha(44) = "B46"
    celdasFicha(45) = "M50": celdasFicha(46) = "N51": celdasFicha(47) = "M52"
    celdasFicha(48) = "M53": celdasFicha(49) = "M54": celdasFicha(50) = "K44": celdasFicha(51) = "K48"
    celdasFicha(52) = "P29"
    celdasFicha(53) = "V40": celdasFicha(54) = "W41": celdasFicha(55) = "V42"
    celdasFicha(56) = "V43": celdasFicha(57) = "V44": celdasFicha(58) = "L43": celdasFicha(59) = "N43"
    celdasFicha(60) = "B46"
    celdasFicha(61) = "L31": celdasFicha(62) = "M32": celdasFicha(63) = "L33"
    celdasFicha(64) = "L34": celdasFicha(65) = "L35": celdasFicha(66) = "K42": celdasFicha(67) = "K38"
    celdasFicha(68) = "P29"
    celdasFicha(69) = "AA34": celdasFicha(70) = "AM10"
    
    ABase = "FichasImpresion.xlsx"
    Windows(ABase).Activate
    directorio = "C:\Consultorias\Tulcan\FichasBorrar\"
    nCasos = InputBox("Numero de archivos", "Archivos", 2)
    
    For k = 1 To nCasos
        APozo = ActiveCell.Value & ".xlsx"
        fila = ActiveCell.Row
        
        'Cálculo altura de la corona
        'Pozo oeste
        alturaBase = Range("BA" & fila).Value
        diametro = Range("BB" & fila).Value
        alturaCoronaO = alturaBase - (diametro / 1000)
        'Pozo sur
        alturaBase = Range("BI" & fila).Value
        diametro = Range("BJ" & fila).Value
        alturaCoronaS = alturaBase - (diametro / 1000)
        'Pozo este
        alturaBase = Range("BQ" & fila).Value
        diametro = Range("BR" & fila).Value
        alturaCoronaE = alturaBase - (diametro / 1000)
        'Pozo norte
        alturaBase = Range("BY" & fila).Value
        diametro = Range("BZ" & fila).Value
        alturaCoronaN = alturaBase - (diametro / 1000)
        
        For j = 0 To 70
            datoLeido(j) = Cells(fila, celdasBase(j)).Value
        Next j
        Workbooks.Open Filename:=directorio + APozo
        Windows(APozo).Activate
        Range("AJ10").Select
        ActiveCell.Value = "Tapado(Si/No):"
        Selection.HorizontalAlignment = xlLeft
        Selection.Font.Bold = False
        
        'fuente para año pozo
        Range("AL9").Select
        Selection.Font.Size = 16
        Selection.Font.Bold = True
        
        'negrilla para cota
        Range("AL11").Select
        Selection.Font.Bold = True
        Columns("AL").ColumnWidth = 16
        
        'negrilla para coordenada norte
        Range("AI14").Select
        Selection.Font.Bold = False
        
        'negrilla para coordenada este
        Range("AI13").Select
        Selection.Font.Bold = False
        
        'negrilla para tapado
        Range("AM10").Select
        Selection.Font.Bold = True
        
        'negrilla para descarga
        Range("AE4").Select
        Selection.Font.Bold = True
        
        'negrilla para coordenada este
        Range("AI13").Select
        Selection.Font.Bold = True
        
        'negrilla para coordenada norte
        Range("AI14").Select
        Selection.Font.Bold = True
        
'Actualiza dato de diametro
'        Range("M32").Select
'        ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""","""",+(R[-1]C[-1]-RC[-1])*1000)"
'        Range("F41").Select
'        ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""","""",+(R[-1]C[-1]-RC[-1])*1000)"
'        Range("W41").Select
'        ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""","""",+(R[-1]C[-1]-RC[-1])*1000)"
'        Range("N51").Select
'        ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""","""",+(R[-1]C[-1]-RC[-1])*1000)"
        Range("L31").Value = ""
        Range("L32").Value = alturaCoronaN
        Range("M32").Value = ""
        Range("L33").Value = ""
        Range("L34").Value = ""
        Range("L35").Value = ""
        Range("K38").Value = ""
        Range("K42").Value = ""
        
        Range("E40").Value = ""
        Range("E41").Value = alturaCoronaO
        Range("F41").Value = ""
        Range("E42").Value = ""
        Range("E43").Value = ""
        Range("E44").Value = ""
        Range("G43").Value = ""
        Range("J43").Value = ""
        
        Range("M50").Value = ""
        Range("M51").Value = alturaCoronaS
        Range("N51").Value = ""
        Range("M52").Value = ""
        Range("M53").Value = ""
        Range("M54").Value = ""
        Range("K48").Value = ""
        Range("K44").Value = ""
        
        Range("V40").Value = ""
        Range("V41").Value = alturaCoronaE
        Range("W41").Value = ""
        Range("V42").Value = ""
        Range("V43").Value = ""
        Range("V44").Value = ""
        Range("N43").Value = ""
        Range("L43").Value = ""
        
        For i = 0 To 70
            Range(celdasFicha(i)).Value = datoLeido(i)
        Next i
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        Windows(ABase).Activate
        ActiveCell.Offset(1, 0).Range("A1").Select
    Next k
End Sub