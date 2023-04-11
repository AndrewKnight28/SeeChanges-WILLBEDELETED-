Sub Macro26()
'
' Macro26 Macro
'

'
Dim wb As Workbook
Dim wb2 As Workbook
Dim wb2_s3 As Worksheet
Dim wb_interfaz As Worksheet
Dim ultimafila As Double


Set wb = ThisWorkbook
Set wb_interfaz = wb.Worksheets(1)
Set wb2 = Workbooks.Open(wb_interfaz.Range("O11").Value)
Set wb2_s3 = wb2.Worksheets("Calculo_S3")
    
    wb2.Activate
    wb2_s3.Activate
    wb2_s3.AutoFilterMode = False
    ultimafila = wb2_s3.Range("A" & Rows.Count).End(xlUp).Row
    
    wb2_s3.Range("AE1").Select
    ActiveCell.FormulaR1C1 = "KEY"
    wb2_s3.Range("AF1").Select
    ActiveCell.FormulaR1C1 = "SALDO TOTAL"
    wb2_s3.Range("AG1").Select
    ActiveCell.FormulaR1C1 = "NUMERO KEY"
    wb2_s3.Range("AH1").Select
    ActiveCell.FormulaR1C1 = "FRACCION"
    wb2_s3.Range("AI1").Select
    ActiveCell.FormulaR1C1 = "PROVISION REAL"
    wb2_s3.Range("AJ1").Select
    ActiveCell.FormulaR1C1 = "KEY2"
    wb2_s3.Range("AK1").Select
    ActiveCell.FormulaR1C1 = "TASA A UTILIZAR"
    wb2_s3.Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]&RC[-7]&RIGHT(RC[-16],4)"
    wb2_s3.Range("AE2").Select
    Selection.AutoFill Destination:=Range("AE2:AE" & ultimafila)
    wb2_s3.Range("AE2:AE" & ultimafila).Select
    wb2_s3.Range("AJ2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-19]&RC[-16]"
    wb2_s3.Range("AJ2").Select
    Selection.AutoFill Destination:=Range("AJ2:AJ" & ultimafila)
    wb2_s3.Range("AJ2:AJ" & ultimafila).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    wb2_s3.Range("AJ2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Range("X1").Select
    Selection.AutoFilter
    wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("Q1:Q" & ultimafila), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Calculo_S3").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
' Se deja ordenado la tabla para proceder a crear la tabla dinamica
    wb2.Sheets.Add After:=ActiveSheet
Dim wb2_s3_2 As Worksheet
Set wb2_s3_2 = ActiveSheet
wb2_s3_2.Name = "Tabla1"

    wb2.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Calculo_S3!R1C1:R" & ultimafila & "C37", Version:=6).CreatePivotTable TableDestination _
        :="Tabla1!R3C1", TableName:="TablaDinámica15", DefaultVersion:=6
    wb2_s3_2.Select
    wb2_s3_2.Cells(3, 1).Select
    
    With wb2_s3_2.PivotTables("TablaDinámica15")

        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("TablaDinámica15").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    wb2_s3_2.PivotTables("TablaDinámica15").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("TablaDinámica15").PivotFields("KEY")
        .Orientation = xlRowField
        .Position = 1
    End With
    wb2_s3_2.PivotTables("TablaDinámica15").AddDataField ActiveSheet.PivotTables _
        ("TablaDinámica15").PivotFields("SALDO"), "Suma de SALDO", xlSum
    wb2_s3_2.PivotTables("TablaDinámica15").AddDataField ActiveSheet.PivotTables _
        ("TablaDinámica15").PivotFields("DOC"), "Cuenta de DOC", xlCount

    

    
    wb2_s3.Activate
    wb2_s3.Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1]," & wb2_s3_2.Name & "!C[-31]:C[-29],2,FALSE)"
    wb2_s3.Range("AF2").Select
    Selection.AutoFill Destination:=Range("AF2:AF" & ultimafila)
    wb2_s3.Range("AF2:AF" & ultimafila).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Range("AG2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2]," & wb2_s3_2.Name & "!C[-32]:C[-30],3,FALSE)"
    wb2_s3.Range("AG2").Select
    Selection.AutoFill Destination:=Range("AG2:AG" & ultimafila)
    wb2_s3.Range("AG2:AG" & ultimafila).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Range("AH2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-11]/RC[-2],0)"
    wb2_s3.Range("AH2").Select
    Selection.AutoFill Destination:=Range("AH2:AH" & ultimafila)
    wb2_s3.Range("AH2:AH" & ultimafila).Select
    wb2_s3.Range("AH2").Select

    wb2_s3.Range("AH2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    wb2_s3.Range("AI2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-11]"
    wb2_s3.Range("AI2").Select
    Selection.AutoFill Destination:=Range("AI2:AI" & ultimafila)
    wb2_s3.Range("AI2:AI" & ultimafila).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

'aca iba el copia correcto
    
    wb2.Sheets.Add After:=ActiveSheet
Dim wb2_s3_3 As Worksheet
Set wb2_s3_3 = ActiveSheet
wb2_s3_3.Name = "Tasas a calcular"
    
    wb2_s3.Select
    wb2_s3.Activate
    wb2_s3.Columns("Q:Q").Select
    Selection.Copy
    wb2_s3_3.Select
    wb2_s3_3.Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3_3.Range("$A$1:$A$1031784").RemoveDuplicates Columns:=1, Header:= _
        xlNo
    wb2_s3_3.Range("A1").Select

'Se realiza LA CORRECCION de la tasa del campo BPE
Dim primfila1 As Double
Dim primfila2 As Double
Dim primfila3 As Double
Dim primfila4 As Double
Dim primfila5 As Double
Dim primfila6 As Double
Dim primfila7 As Double
Dim primfila8 As Double
Dim primfila9 As Double

Dim ultfila1 As Double
Dim ultfila2 As Double
Dim ultfila3 As Double
Dim ultfila4 As Double
Dim ultfila5 As Double
Dim ultfila6 As Double
Dim ultfila7 As Double
Dim ultfila8 As Double
Dim ultfila9 As Double
    
Dim primfilaINMOBI As Double
Dim ultfilaINMOBI As Double
Dim primfilaINMOBI2 As Double
'primera fila BPE
  primfila1 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A2").Value, _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfila1
  ultfila1 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A2").Value, _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfila1
'primera fila Consumo
  primfila2 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A3").Value, _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfila1
  ultfila2 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A3").Value, _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfila2
'primera fila Convenios
  primfila3 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A4").Value, _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfila3
  ultfila3 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A4").Value, _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfila3
'primera fila Corporativos
  primfila4 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A5").Value, _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfila4
  ultfila4 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A5").Value, _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfila4
'primera fila Empresas
  primfila5 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A6").Value, _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfila5
  ultfila5 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A6").Value, _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfila5
'primera fila Hipotecas
  primfila6 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A7").Value, _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfila6
  ultfila6 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A7").Value, _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfila6
'primera fila Inmobiliarias
  primfila7 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A8").Value, _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfila7
  ultfila7 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A8").Value, _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfila7
'primera fila Tarjetas
  primfila8 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A9").Value, _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfila8
  ultfila8 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A9").Value, _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfila8
'primera fila Vehicular
  primfila9 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A10").Value, _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfila9
  ultfila9 = wb2_s3.Columns("Q").Find(wb2_s3_3.Range("A10").Value, _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfila9
'fila de inmobiliaria
  primfilaINMOBI = wb2_s3.Columns("Q").Find("Inmobiliarias", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfilaINMOBI
  ultfilaINMOBI = wb2_s3.Columns("Q").Find("Inmobiliarias", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfilaINMOBI
  primfilaINMOBI2 = wb2_s3_3.Columns("A").Find("Inmobiliarias", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfilaINMOBI2
  
    wb2_s3_3.Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(Calculo_S3!R" & primfila1 & "C[21]:R" & ultfila1 & "C[21])"
    wb2_s3_3.Columns("B:B").Select
    Selection.NumberFormat = "General"
 'empieza la suma producto de las carteras
    wb2_s3_3.Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(Calculo_S3!R" & primfila1 & "C[17]:R" & ultfila1 & "C[17],Calculo_S3!R" & primfila1 & "C[20]:R" & ultfila1 & "C[20])"
    wb2_s3_3.Range("D2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-2]"
    wb2_s3_3.Range("C3").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(Calculo_S3!R" & primfila2 & "C[17]:R" & ultfila2 & "C[17],Calculo_S3!R" & primfila2 & "C[20]:R" & ultfila2 & "C[20])"
    wb2_s3_3.Range("C4").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(Calculo_S3!R" & primfila3 & "C[17]:R" & ultfila3 & "C[17],Calculo_S3!R" & primfila3 & "C[20]:R" & ultfila3 & "C[20])"
    wb2_s3_3.Range("C5").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(Calculo_S3!R" & primfila4 & "C[17]:R" & ultfila4 & "C[17],Calculo_S3!R" & primfila4 & "C[20]:R" & ultfila4 & "C[20])"
    wb2_s3_3.Range("C6").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(Calculo_S3!R" & primfila5 & "C[17]:R" & ultfila5 & "C[17],Calculo_S3!R" & primfila5 & "C[20]:R" & ultfila5 & "C[20])"
    wb2_s3_3.Range("C7").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(Calculo_S3!R" & primfila6 & "C[17]:R" & ultfila6 & "C[17],Calculo_S3!R" & primfila6 & "C[20]:R" & ultfila6 & "C[20])"
    wb2_s3_3.Range("C8").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(Calculo_S3!R" & primfila7 & "C[17]:R" & ultfila7 & "C[17],Calculo_S3!R" & primfila7 & "C[20]:R" & ultfila7 & "C[20])"
    wb2_s3_3.Range("C9").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(Calculo_S3!R" & primfila8 & "C[17]:R" & ultfila8 & "C[17],Calculo_S3!R" & primfila8 & "C[20]:R" & ultfila8 & "C[20])"
    wb2_s3_3.Range("C10").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(Calculo_S3!R" & primfila9 & "C[17]:R" & ultfila9 & "C[17],Calculo_S3!R" & primfila9 & "C[20]:R" & ultfila9 & "C[20])"
'continuamos con la suma del resto de las carteras
    wb2_s3_3.Range("B3").Select
    ActiveCell.FormulaR1C1 = "=SUM(Calculo_S3!R" & primfila2 & "C[21]:R" & ultfila2 & "C[21])"
    wb2_s3_3.Range("B4").Select
    ActiveCell.FormulaR1C1 = "=SUM(Calculo_S3!R" & primfila3 & "C[21]:R" & ultfila3 & "C[21])"
    wb2_s3_3.Range("B5").Select
    ActiveCell.FormulaR1C1 = "=SUM(Calculo_S3!R" & primfila4 & "C[21]:R" & ultfila4 & "C[21])"
    wb2_s3_3.Range("B6").Select
    ActiveCell.FormulaR1C1 = "=SUM(Calculo_S3!R" & primfila5 & "C[21]:R" & ultfila5 & "C[21])"
    wb2_s3_3.Range("B7").Select
    ActiveCell.FormulaR1C1 = "=SUM(Calculo_S3!R" & primfila6 & "C[21]:R" & ultfila6 & "C[21])"
    wb2_s3_3.Range("B8").Select
    ActiveCell.FormulaR1C1 = "=SUM(Calculo_S3!R" & primfila7 & "C[21]:R" & ultfila7 & "C[21])"
    wb2_s3_3.Range("B9").Select
    ActiveCell.FormulaR1C1 = "=SUM(Calculo_S3!R" & primfila8 & "C[21]:R" & ultfila8 & "C[21])"
    Range("B10").Select
    ActiveCell.FormulaR1C1 = "=SUM(Calculo_S3!R" & primfila9 & "C[21]:R" & ultfila9 & "C[21])"
    
 
    
    wb2_s3_3.Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D10"), Type:=xlFillDefault
    Range("D2:D10").Select
    Columns("D:D").Select
    Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"


    wb2_s3.Activate
    wb2_s3.Range("AK2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(Calculo_S3!RC[-17]=0,LOOKUP(RC[-1],'Tasas a calcular'!C[-36]:C[-33]),Calculo_S3!RC[-17])"
    wb2_s3.Range("AK2").Select
    Selection.AutoFill Destination:=Range("AK2:AK" & ultimafila)
    wb2_s3.Range("AK2:AK" & ultimafila).Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
'para inmobiliarias
    wb2_s3.Range("AL1").Select
    ActiveCell.FormulaR1C1 = "KEY 3"
    wb2_s3.Range("AM1").Select
    ActiveCell.FormulaR1C1 = "TASA INMBILIARIA"

    Selection.AutoFilter
    wb2_s3.Range("$A$1:$AM$" & ultimafila).AutoFilter Field:=17, Criteria1:= _
        "Inmobiliarias"
    ActiveWorkbook.Worksheets("Calculo_S3").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Calculo_S3").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("M1:M" & ultimafila), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Calculo_S3").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim primfilaIM As Double
    Dim ultfilaIM As Double
    Dim primfilaLPC As Double
    Dim ultfilaLPC As Double
  
  primfilaIM = wb2_s3.Columns("M").Find("$IM", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfilaIM
  ultfilaIM = wb2_s3.Columns("M").Find("$IM", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfilaIM
  primfilaLPC = wb2_s3.Columns("M").Find("LPC", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfilaLPC
  ultfilaLPC = wb2_s3.Columns("M").Find("LPC", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfilaLPC
    
    wb2_s3.Range("$A$1:$AM$" & ultimafila).AutoFilter Field:=13, Criteria1:="$IM"
    wb2_s3.Range("AN" & primfilaIM).Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(R" & primfilaLPC & "C[-20]:R" & ultfilaLPC & "C[-20],R" & primfilaLPC & "C[-17]:R" & ultfilaLPC & "C[-17])/SUM(R" & primfilaLPC & "C[-17]:R" & ultfilaLPC & "C[-17])"
    Range("AN" & primfilaIM).Select
    Selection.Copy
    Range("AM" & primfilaIM & ":AM" & ultfilaIM).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Selection.Copy
    
    Range("AN" & primfilaIM).Select
    Selection.Copy
    wb2_s3_3.Select
    wb2_s3_3.Range("D" & primfilaINMOBI2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Activate
    wb2_s3.Range("$A$1:$AM$" & ultimafila).AutoFilter Field:=13

    wb2_s3.Range("AK" & primfilaIM).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(Calculo_S3!RC[-17]=0,LOOKUP(Calculo_S3!RC[-1],'Tasas a calcular'!C[-36]:C[-33]),Calculo_S3!RC[-17])"
    Range("AK" & primfilaIM).Select
    Selection.AutoFill Destination:=Range("AK" & primfilaINMOBI & ":AK" & ultfilaINMOBI), Type:= _
        xlFillDefault
    Range("AK" & primfilaINMOBI & ":AK" & ultfilaINMOBI).Select
    wb2_s3.Range("AM" & primfilaIM & ":AM" & ultfilaIM).Select
    Selection.Copy
    wb2_s3.Range("AK" & primfilaIM & ":AK" & ultfilaIM).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AK" & primfilaINMOBI & ":AK" & ultfilaINMOBI).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.AutoFilter
    
    
    
    wb2_s3.Range("AK2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    wb2_s3.Range("U2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

'se completa la tasa
    wb2_s3.Range("AI2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wb2_s3.Range("X2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    

    wb2_s3.Range("AK2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    wb2_s3.Range("U2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    wb2_s3.Columns("AE:AN").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Application.DisplayAlerts = False
    wb2_s3_2.Select
    ActiveWindow.SelectedSheets.Delete
    wb2_s3_3.Select
    ActiveWindow.SelectedSheets.Delete
    wb2_s3.Activate
    
    wb2_s3.Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    wb2_s3.Range("Y2").Select
    Selection.AutoFill Destination:=Range("Y2:Y" & ultimafila)
    wb2_s3.Range("Y2:Y" & ultimafila).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


' TABLA DINAMICA PARA VALIDACION

    wb_interfaz.Activate


End Sub
