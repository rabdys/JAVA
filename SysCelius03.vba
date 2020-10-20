Sub FormatoInicialSCV_0()
'
' Formato inicial de archivo CSV

'
    Sheets("CSV01").Select
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array _
        (20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), _
        Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array( _
        33, 1), Array(34, 1), Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), _
        Array(40, 1), Array(41, 1), Array(42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array( _
        46, 1), Array(47, 1), Array(48, 1), Array(49, 1), Array(50, 1), Array(51, 1), Array(52, 1), _
        Array(53, 1), Array(54, 1), Array(55, 1), Array(56, 1), Array(57, 1), Array(58, 1), Array( _
        59, 1), Array(60, 1), Array(61, 1), Array(62, 1), Array(63, 1), Array(64, 1), Array(65, 1), _
        Array(66, 1), Array(67, 1), Array(68, 1), Array(69, 1)), TrailingMinusNumbers:=True
End Sub

Function EstaVacio(xValor)
    If xValor = "" Then 
        EstaVacio = True
    End If
    If xValor <> "" Then
        EstaVacio = False
    End If
End Function


Sub Ins2ColxMaq_1()
' Paso 1
' Inserta dos columnas por máquina
'
'
' la fila 9 contiene los id de las máquinas que se evaluan en el archivo
    Range("A9").Select
    lMasMq = True
    Do While lMasMq = True
        Selection.EntireColumn.Insert
        Selection.EntireColumn.Insert
        ActiveCell.Offset(0, 6).Select
        IdMaquina = ActiveCell.Value
        lMasMq = (EstaVacio(IdMaquina) = False)
        ActiveCell.Offset(0, -1).Select
   Loop
   ' LLENADO DE TITULOS
   
   
End Sub

Sub InsTit1xMaq_2()
' Paso 2 Coloca los titulos en la primera fila de cada máquina

    Range("D9").Select
    PuntoIdMaq = ActiveCell.Address(False, False)
    lMasMq = True
    Do While lMasMq = True
        ' Datos a Copiar
        IdMaquina = ActiveCell.Value
        ActiveCell.Offset(2, 0).Select
        DesrMaquina = ActiveCell.Value
        
        
        ActiveCell.Offset(4, -3).Select
        ActiveCell.FormulaR1C1 = "Address"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.FormulaR1C1 = "Tag Name"
        ActiveCell.Offset(1, -1).Select
        ActiveCell.FormulaR1C1 = IdMaquina
        ActiveCell.Offset(0, 1).Select
        ActiveCell.FormulaR1C1 = DesrMaquina
        ActiveCell.Offset(1, 0).Select
        
        DescrMaqDesde = ActiveCell.Address(False, False)
        ActiveCell.Offset(0, -1).Select
        IdMaqDesde = ActiveCell.Address(False, False)
        
        'Encuentra la última fila con datos TANTO PARA ID COMO PARA DESCR
        ActiveCell.Offset(0, 2).Select
        Selection.End(xlDown).Select
        UltimaFilaMQ = ActiveCell.Address(False, False)
        ActiveCell.Offset(0, -1).Select
        DescrMaqHasta = ActiveCell.Address(False, False)
        ActiveCell.Offset(0, -1).Select
        IdMaqHasta = ActiveCell.Address(False, False)
        
        'Llenado de idMaquina y descripcion
        Rangox = IdMaqDesde & ":" & IdMaqHasta
        Range(Rangox).Value = IdMaquina
        Rangox = DescrMaqDesde & ":" & DescrMaqHasta
        Range(Rangox).Value = DesrMaquina
         
        'Punto Inicial para saltar a la otra máquina
        Range(PuntoIdMaq).Select
        ActiveCell.Offset(0, 5).Select
        PuntoIdMaq = ActiveCell.Address(False, False)
        IdMaquina = ActiveCell.Value
        lMasMq = (EstaVacio(IdMaquina) = False)
   Loop
End Sub

Sub EliminarTitulos_3()
' Paso3 : Eliminación de Titulos INNECESARIOS
'

'
    Rows("1:1").Select
    For i = 1 To 14
        Selection.Delete Shift:=xlUp
    Next
    
End Sub

Sub Grafica_LinealxMaquina_4()
'
' Generación de las Gráficas lineales.
'

'
Dim lMasMaq As Boolean

    'Ciclo de recorrido de cada máquina
    
    Range("A2").Select
    PuntoIdMaq = ActiveCell.Address(False, False)
    lMasMaq = True
    IdMaquina = Range(PuntoIdMaq).Value
    
    Do While lMasMaq
        
        If EstaVacio(IdMaquina) = False Then
            ' Calcula los rangos desde la posicion anterior
            
            ActiveCell.Offset(0, 1).Select
            DescrMaquina = ActiveCell.Value
            
            ActiveCell.Offset(0, 1).Select
            RangoDesde = ActiveCell.Address(False, False)
            ActiveCell.Offset(1000, 2).Select
            ' SECOMENTAREAR ESTA LIENA CUANDO SE HAYA TERMINADO TODO EL PROGRAMA PARA QUE TOME TODOS LOS RANGOS.
            ' Selection.End(xlDown).Select
            
            RangoHasta = ActiveCell.Address(False, False)
            
            
            ' Crea la hoja de las graficas de la máquina
            NameHoja = "Graficos_" & IdMaquina
            
            Sheets.Add After:=ActiveSheet
            Sheets(NameHoja).Select
            Sheets(NameHoja).Name = NameHoja	
            
            ActiveSheet.Shapes.AddChart2(227, xlLine).Select
            xRango = NameHoja & "!" & RangoDesde & ":" & RangoHasta
            ActiveChart.SetSourceData Source:=Range(xRango)
        End If
        'Punto Inicial para saltar a la otra máquina
        Range(PuntoIdMaq).Select
        ActiveCell.Offset(0, 5).Select
        PuntoIdMaq = ActiveCell.Address(False, False)
        IdMaquina = ActiveCell.Value
        lMasMq = (EstaVacio(IdMaquina) = False)
    
    
    Loop

End Sub
