Sub Macro20()
'
' Macro20 Macro
'

'
Dim wb As Workbook
Dim wb2 As Workbook
Dim wb2_s3 As Worksheet
Dim wb_interfaz As Worksheet
Dim ultimafila As Double
Dim ultimafila3 As String
Dim primfilazztarjetas As Double
Dim ultfilatarjetas As Double
Dim primfilatarjetas As Double
Set wb = ThisWorkbook
Set wb_interfaz = wb.Worksheets(1)
Set wb2 = Workbooks.Open(wb_interfaz.Range("O11").Value)
Set wb2_s3 = wb2.Worksheets("Calculo S3")
    
    
    wb2_s3.Activate
    wb2_s3.Range("AE1").Select
    ActiveCell.FormulaR1C1 = _
        "=[" & wb.Name & "]" & wb_interfaz.Name & "!R15C19-[" & wb.Name & "]" & wb_interfaz.Name & "!R14C19"
    wb2_s3.Range("AE1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    Dim diasvalidador As Variant
    diasvalidador = wb2_s3.Range("AE1").Value
    Debug.Print diasvalidador

    ultimafila = wb2_s3.Range("A" & Rows.Count).End(xlUp).Row
    wb2_s3.Activate
    Cells.Replace What:="Tarjetas", Replacement:="zzTarjetas", LookAt:=xlPart _
        , searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    wb2_s3.Range("Q1").Select
    Selection.AutoFilter
    wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("Q1:Q" & ultimafila), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter

    
    wb2.Activate
    wb2_s3.Activate
    wb2_s3.Range("AD1").Select
    ActiveCell.FormulaR1C1 = "TRAMO"
    wb2_s3.Range("AF1").Select
    ActiveCell.FormulaR1C1 = "=""MENOR A ""&RC[-1]"
    wb2_s3.Range("AG1").Select
    ActiveCell.FormulaR1C1 = "=""MAYOR A ""&RC[-2]"
    
    

        
    wb2_s3.Range("Q1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AG$" & ultimafila).AutoFilter Field:=17, Criteria1:= _
        "zzTarjetas"
  
  primfilazztarjetas = wb2_s3.Columns("Q").Find("zzTarjetas", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfilazztarjetas
  
    wb2_s3.Rows(primfilazztarjetas & ":" & primfilazztarjetas + 1).Select
    wb2_s3.Range("E" & primfilazztarjetas).Activate
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Cells.Replace What:="zzTarjetas", Replacement:="Tarjetas", LookAt:=xlPart _
        , searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
  primfilatarjetas = wb2_s3.Columns("Q").Find("Tarjetas", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
  Debug.Print primfilatarjetas & "Z"
  ultfilatarjetas = wb2_s3.Columns("Q").Find("Tarjetas", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  Debug.Print ultfilatarjetas & "Z"
    
    Range("A1:AC1").Select
    Range("AC1").Activate
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    wb2_s3.Range("A" & primfilatarjetas - 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.AutoFilter
    
    wb2_s3.Range("AD2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-16]<=30,""TRAMO1"",IF(RC[-16]<=90,""TRAMO2"",""TRAMO3""))"
    Range("AD2").Select
    Selection.AutoFill Destination:=Range("AD2:AD" & primfilatarjetas - 3)
    wb2_s3.Range("AD2:AD" & primfilatarjetas - 3).Select
    wb2_s3.Range("AC2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-15]<R1C31,R1C32,R1C33)"
    wb2_s3.Range("AC2").Select
    Selection.AutoFill Destination:=Range("AC2:AC" & primfilatarjetas - 3)
    wb2_s3.Range("AD2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("AD" & primfilatarjetas - 3).Select
    Selection.End(xlUp).Select
    Range("AD2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    

    wb2_s3.Range("AD1").Select
    Selection.AutoFilter
    wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("AD1:AD" & primfilatarjetas - 3), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    wb2_s3.Range("AC2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Range("AF1:AG1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("N1:N" & ultfilatarjetas - 3), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
 
 
 'Añadimos los filtros correspondientes por tramo
'TRAMO 1 - HASTA 30 DIAS
Dim primfilatr1 As Double
Dim ultfilatr1 As Double
Dim primfilatr2 As Double
Dim ultfilatr2 As Double
Dim primfilatr3 As Double
Dim ultfilatrvalidador As Double
Dim primfilatrvalidador As Double
Dim ultfilatrvalidador2 As Double

primfilatr1 = wb2_s3.Columns("AD").Find("TRAMO1", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row

ultfilatr1 = wb2_s3.Columns("AD").Find("TRAMO1", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row

primfilatr2 = wb2_s3.Columns("AD").Find("TRAMO2", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row

ultfilatr2 = wb2_s3.Columns("AD").Find("TRAMO2", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    
primfilatr3 = wb2_s3.Columns("AD").Find("TRAMO3", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row

ultfilatrvalidador2 = wb2_s3.Columns("AC").Find(wb2_s3.Range("AF1").Value, _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
primfilatrvalidador = wb2_s3.Columns("AC").Find(wb2_s3.Range("AG1").Value, _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
ultfilatrvalidador = wb2_s3.Columns("AC").Find(wb2_s3.Range("AG1").Value, _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
'tramo 1
    wb2_s3.Range("Z2").Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-5])^(RC[-12]/360)-1)*RC[-3]"
    wb2_s3.Range("Z2").Select
    Selection.AutoFill Destination:=Range("Z2:Z" & ultfilatr1), Type:=xlFillDefault
    Range("Z2:Z" & ultfilatr1).Select
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("AB2").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("AA2").Select
    Selection.AutoFill Destination:=Range("AA2:AA" & ultfilatr1), Type:=xlFillDefault
    Range("AA2:AA26").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("AB2").Select
    Selection.AutoFill Destination:=Range("AB2:AB" & ultfilatr1), Type:=xlFillDefault
    Range("AB2:AB" & ultfilatr1).Select
'tramo 2
    wb2_s3.Range("Z" & primfilatr2).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-5])^(30/360)-1)*RC[-3]"
    wb2_s3.Range("AA" & primfilatr2).Select
    ActiveCell.FormulaR1C1 = "=(((1+RC[-6])^(RC[-13]/360)-1)*RC[-4])-RC[-1]"
    wb2_s3.Range("AB" & primfilatr2).Select
    ActiveCell.FormulaR1C1 = "0"
    wb2_s3.Range("Z" & primfilatr2).Select
    Selection.AutoFill Destination:=Range("Z" & primfilatr2 & ":Z" & ultfilatr2), Type:=xlFillDefault
    wb2_s3.Range("Z" & primfilatr2 & ":Z" & ultfilatr2).Select
    wb2_s3.Range("AA" & primfilatr2).Select
    Selection.AutoFill Destination:=Range("AA" & primfilatr2 & ":AA" & ultfilatr2), Type:=xlFillDefault
    wb2_s3.Range("AA" & primfilatr2 & ":AA" & ultfilatr2).Select
    wb2_s3.Range("AB" & primfilatr2).Select
    Selection.AutoFill Destination:=Range("AB" & primfilatr2 & ":AB" & ultfilatr2), Type:=xlFillDefault
    wb2_s3.Range("AB" & primfilatr2 & ":AB" & ultfilatr2).Select

'tramo 3
    wb2_s3.Range("Z" & primfilatr3).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-5])^(30/360)-1)*RC[-3]"
    wb2_s3.Range("AA" & primfilatr3).Select
    ActiveCell.FormulaR1C1 = "=(((1+RC[-6])^(90/360)-1)*RC[-4])-RC[-1]"
    wb2_s3.Range("AB" & primfilatr3).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-7])^(RC[-6]/360)-1)*RC[-3]"
    wb2_s3.Range("Z" & primfilatr3).Select
    Selection.AutoFill Destination:=Range("Z" & primfilatr3 & ":Z" & ultfilatrvalidador2), Type:=xlFillDefault
    wb2_s3.Range("Z" & primfilatr3 & ":Z" & ultfilatrvalidador2).Select
    wb2_s3.Range("AA" & primfilatr3).Select
    Selection.AutoFill Destination:=Range("AA" & primfilatr3 & ":AA" & ultfilatrvalidador2), Type:=xlFillDefault
    wb2_s3.Range("AA" & primfilatr3 & ":AA" & ultfilatrvalidador2).Select
    wb2_s3.Range("AB" & primfilatr3).Select
    Selection.AutoFill Destination:=Range("AB" & primfilatr3 & ":AB" & ultfilatrvalidador2), Type:=xlFillDefault
    wb2_s3.Range("AB" & primfilatr3 & ":AB" & ultfilatrvalidador2).Select
    
' TRAMO 4
    wb2_s3.Range("Z" & primfilatrvalidador).Select
    ActiveCell.FormulaR1C1 = _
        "=(+(1+RC[-5])^((IF((RC[-12]-R1C31)<=30,30-(RC[-12]-R1C31),0))/360)-1)*RC[-3]"
    Range("AA" & primfilatrvalidador).Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = _
        "=((+(1+RC[-6])^((IF(RC[-5]<=" & diasvalidador & "," & diasvalidador & "-RC[-5],0))/360)-1)*RC[-4])-RC[-1]"
    Range("AB" & primfilatrvalidador).Select
    ActiveCell.FormulaR1C1 = _
        "=(+(1+RC[-7])^((IF(RC[-6]<=R1C31,RC[-6],R1C31))/360)-1)*RC[-3]"

    wb2_s3.Range("Z" & primfilatrvalidador).Select
    Selection.AutoFill Destination:=Range("Z" & primfilatrvalidador & ":Z" & ultfilatrvalidador), Type:=xlFillDefault
    wb2_s3.Range("Z" & primfilatrvalidador & ":Z" & ultfilatrvalidador).Select
    wb2_s3.Range("AA" & primfilatrvalidador).Select
    Selection.AutoFill Destination:=Range("AA" & primfilatrvalidador & ":AA" & ultfilatrvalidador), Type:=xlFillDefault
    wb2_s3.Range("AA" & primfilatrvalidador & ":AA" & ultfilatrvalidador).Select
    wb2_s3.Range("AB" & primfilatrvalidador).Select
    Selection.AutoFill Destination:=Range("AB" & primfilatrvalidador & ":AB" & ultfilatrvalidador), Type:=xlFillDefault
    wb2_s3.Range("AB" & primfilatrvalidador & ":AB" & ultfilatrvalidador).Select

    
'Calculo de intereses para neto negativo
    
    wb2_s3.Range("AH1").Select
    ActiveCell.FormulaR1C1 = "PROVISION"
    wb2_s3.Range("AH2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-9]<0,""MENOR QUE 0"",""MAYOR QUE 0"")"
    wb2_s3.Range("AH2").Select
    Selection.AutoFill Destination:=Range("AH2:AH" & primfilatarjetas - 3)
    wb2_s3.Range("AH2:AH" & ultfilatarjetas - 3).Select
    wb2_s3.Range("AH2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    wb2_s3.Range("AH1").Select
    Selection.AutoFilter
    wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("Y1:Y" & primfilatarjetas - 3), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim primmenor As Double
    Dim ultmenor As Double
primmenor = wb2_s3.Columns("AH").Find("MENOR QUE 0", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print primmenor
ultmenor = wb2_s3.Columns("AH").Find("MENOR QUE 0", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ultmenor
    
    wb2_s3.Range("Z" & primmenor).Select
    ActiveCell.FormulaR1C1 = "0"
    wb2_s3.Range("AA" & primmenor).Select
    ActiveCell.FormulaR1C1 = "0"
    wb2_s3.Range("AB" & primmenor).Select
    ActiveCell.FormulaR1C1 = "0"
    wb2_s3.Range("Z" & primmenor).Select
    Selection.AutoFill Destination:=Range("Z" & primmenor & ":Z" & ultmenor), Type:=xlFillDefault
    Range("Z" & primmenor & ":Z" & ultmenor).Select

    Range("AA" & primmenor).Select
    Selection.AutoFill Destination:=Range("AA" & primmenor & ":AA" & ultmenor), Type:=xlFillDefault
    Range("AA" & primmenor & ":AA" & ultmenor).Select

    Range("AB" & primmenor).Select
    Selection.AutoFill Destination:=Range("AB" & primmenor & ":AB" & ultmenor), Type:=xlFillDefault
    Range("AB" & primmenor & ":AB" & ultmenor).Select

    Selection.AutoFilter
    Range("AB1").Select
    Selection.AutoFilter
    wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("N1:N" & ultfilatarjetas - 3), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter

'CALCULAR LOS INTERESES DE LA CARTERA TARJETAS

    wb2_s3.Range("AC" & primfilatarjetas).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(RC[-15]<R1C30,R1C32,R1C33)"
    wb2_s3.Range("AC" & primfilatarjetas).Select
    Selection.AutoFill Destination:=Range("AC" & primfilatarjetas & ":AC" & ultfilatarjetas)
    wb2_s3.Range("AC" & primfilatarjetas & ":AC" & ultfilatarjetas).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Range("AD" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "TOTAL DE INTERES"
    wb2_s3.Range("AE" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "PROV INTERES"
    wb2_s3.Range("AF" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "INTERES NETO"
    wb2_s3.Range("AG" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "TIPO"
    wb2_s3.Range("AH" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "CRUCE MES"
    wb2_s3.Range("AI" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "STAGE MES"
    wb2_s3.Range("AJ" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "SALDO MES"
    wb2_s3.Range("AK" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "DIAS MES"
    wb2_s3.Range("AL" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "TASA MES"
    wb2_s3.Range("AM" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "DIAS MES-MES"
    wb2_s3.Range("AN" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "CRUCE MES"
    wb2_s3.Range("AO" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "TIPO DOC"
    wb2_s3.Range("AD" & primfilatarjetas).Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3])"
    wb2_s3.Range("AD" & primfilatarjetas).Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-4]:RC[-2])"
    wb2_s3.Range("AD" & primfilatarjetas).Select
    Selection.AutoFill Destination:=Range("AD" & primfilatarjetas & ":AD" & ultfilatarjetas)
    wb2_s3.Range("AD" & primfilatarjetas & ":AD" & ultfilatarjetas).Select
    wb2_s3.Range("AF" & primfilatarjetas).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    wb2_s3.Range("AF" & primfilatarjetas).Select
    Selection.AutoFill Destination:=Range("AF" & primfilatarjetas & ":AF" & ultfilatarjetas)

'Abrir los archivos de s1 y s3
Dim Dialogo1 As FileDialog
Dim ArchivoS1Seleccionado As Variant

Set Dialogo1 = Application.FileDialog(msoFileDialogFilePicker)
algo = False
Dialogo1.Title = "Escoger Archivo S1 y S2 del mes anterior"
If Dialogo1.Show = -1 Then

    For Each ArchivoS1Seleccionado In Dialogo1.SelectedItems
        
        wb_interfaz.Range("O16").FormulaR1C1 = ArchivoS1Seleccionado
        Workbooks.Open Filename:= _
        ArchivoS1Seleccionado
    
    Next ArchivoS1Seleccionado
    
Else
MsgBox "No se ha escogido el archivo "

End If

Dim Dialogo2 As FileDialog
Dim ArchivoS3Seleccionado As Variant

Set Dialogo2 = Application.FileDialog(msoFileDialogFilePicker)
algo = False
Dialogo1.Title = "Escoger Archivo S3 del mes anterior"
If Dialogo1.Show = -1 Then

    For Each ArchivoS3Seleccionado In Dialogo1.SelectedItems
        
        wb_interfaz.Range("O17").FormulaR1C1 = ArchivoS3Seleccionado
        Workbooks.Open Filename:= _
        ArchivoS3Seleccionado
    
    Next ArchivoS3Seleccionado
    
Else
MsgBox "No se ha escogido el archivo "

End If
'Nombrar los archivos y las hojas

Dim wbs1 As Workbook
Dim wbs3 As Workbook

Dim wbs1_detalles As Worksheet
Dim wbs3_calculoss3 As Worksheet

Set wbs1 = Workbooks.Open(wb_interfaz.Range("O16").Value)
Set wbs3 = Workbooks.Open(wb_interfaz.Range("O17").Value)

Set wbs1_detalles = wbs1.Worksheets("DETALLE 8104 0-90 días")
Set wbs3_calculoss3 = wbs3.Worksheets("Calculo S3")
    
Dim ults1tarjetas As Double
Dim prims1tarjetas As Double

prims1tarjetas = wbs1_detalles.Columns("Q").Find("Tarjetas", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print prims1tarjetas & XD
ults1tarjetas = wbs1_detalles.Columns("Q").Find("Tarjetas", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ults1tarjetas & XD
    
    
    wb2_s3.Activate
    wb2_s3.Range("AH" & primfilatarjetas).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-19],'[" & wbs1.Name & "]" & wbs1_detalles.Name & "'!R" & prims1tarjetas & "C15:R" & ults1tarjetas & "C17,3,0)"
    wb2_s3.Range("AH" & primfilatarjetas).Select
    Selection.AutoFill Destination:=Range("AH" & primfilatarjetas & ":AH" & ultfilatarjetas)
    wb2_s3.Range("AH" & primfilatarjetas & ":AH" & ultfilatarjetas).Select
    wb2_s3.Range("AH" & primfilatarjetas & ":AH" & ultfilatarjetas).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Range("AH" & primfilatarjetas - 1).Select
    Selection.AutoFilter
    wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("AH" & primfilatarjetas & ":AH" & ultfilatarjetas), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    wb2_s3.Range("AG" & primfilatarjetas).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(RC[1]=""Tarjetas"",""NUEVO"","" ""),"" "")"
    wb2_s3.Range("AG" & primfilatarjetas).Select
    Selection.AutoFill Destination:=Range("AG" & primfilatarjetas & ":AG" & ultfilatarjetas)
    wb2_s3.Range("AG" & primfilatarjetas & ":AG" & ultfilatarjetas).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

Dim ults3tarjetas As Double
Dim prims3tarjetas As Double
Dim primcrucetarjetas As Double
Dim ultcrucetarjetas As Double

prims3tarjetas = wbs3_calculoss3.Columns("Q").Find("Tarjetas", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print prims3tarjetas & "XD"
ults3tarjetas = wbs3_calculoss3.Columns("Q").Find("Tarjetas", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ults3tarjetas & "XD"
ultcrucetarjetas = wb2_s3.Columns("AH").Find("Tarjetas", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ultcrucetarjetas & "XD"


    wb2_s3.Range("AH" & ultcrucetarjetas + 1).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-33],'[" & wbs3.Name & "]" & wbs3_calculoss3.Name & "'!R" & prims3tarjetas & "C1:R" & ults3tarjetas & "C13,13,0)"
    wb2_s3.Range("AH" & ultcrucetarjetas + 1).Select
    Selection.AutoFill Destination:=Range("AH" & ultcrucetarjetas + 1 & ":AH" & ultfilatarjetas)
    wb2_s3.Range("AH" & ultcrucetarjetas + 1 & ":AH" & ultfilatarjetas).Select
    wb2_s3.Range("AH" & ultcrucetarjetas + 1 & ":AH" & ultfilatarjetas).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Range("AG" & ultcrucetarjetas + 1).Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[1],""NUEVO"")"
    wb2_s3.Range("AG" & ultcrucetarjetas + 1).Select
    Selection.AutoFill Destination:=Range("AG" & ultcrucetarjetas + 1 & ":AG" & ultfilatarjetas)
    wb2_s3.Range("AG" & ultcrucetarjetas + 1 & ":AH" & ultfilatarjetas).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

Dim ults3nuevos As Double
    wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("AG" & primfilatarjetas & ":AG" & ultfilatarjetas), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
  ults3nuevos = wb2_s3.Columns("AG").Find("NUEVO", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ults3nuevos & "XD"
    Rows(ults3nuevos + 1 & ":" & ults3nuevos + 2).Select
    Range("T" & ults3nuevos + 1).Activate
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AO" & primfilatarjetas - 1).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.Copy
    Range("A" & ults3nuevos + 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    wb2_s3.Range("AP" & primfilatarjetas - 1).Select
    ActiveCell.FormulaR1C1 = "TRAMO"
    Range("AP" & primfilatarjetas).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-28]<=30,""TRAMOA1"",IF(RC[-28]<=90,""TRAMOA2"",""TRAMOA3""))"
    Range("AP" & primfilatarjetas).Select
    Selection.AutoFill Destination:=Range("AP" & primfilatarjetas & ":AP" & ults3nuevos)
    Range("AP" & primfilatarjetas & ":AP" & ults3nuevos).Select
    Range("AP" & primfilatarjetas).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("AP" & primfilatarjetas).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Selection.AutoFilter
    wb2_s3.Range("AP" & primfilatarjetas - 1).Select
    Selection.AutoFilter
    wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("AP" & primfilatarjetas & ":AP" & ults3nuevos), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("N" & primfilatarjetas & ":N" & ults3nuevos), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
Dim primtramoa1 As Double
Dim primtramoa2 As Double
Dim primtramoa3 As Double
Dim ulttramoa1 As Double
Dim ulttramoa2 As Double
Dim ulttramoa3 As Double

primtramoa1 = wb2_s3.Columns("AP").Find("TRAMOA1", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print prims3tarjetas & "XD"
ulttramoa1 = wb2_s3.Columns("AP").Find("TRAMOA1", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print prims3tarjetas & "XD"
primtramoa2 = wb2_s3.Columns("AP").Find("TRAMOA2", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print prims3tarjetas & "XD"
ulttramoa2 = wb2_s3.Columns("AP").Find("TRAMOA2", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print prims3tarjetas & "XD"
primtramoa3 = wb2_s3.Columns("AP").Find("TRAMOA3", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print prims3tarjetas & "XD"
ulttramoa3 = wb2_s3.Columns("AP").Find("TRAMOA3", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print prims3tarjetas & "XD"

 'TRAMOA1
    wb2_s3.Range("Z" & primfilatarjetas).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-5])^(RC[-12]/360)-1)*RC[-3]"
    Selection.AutoFill Destination:=Range("Z" & primfilatarjetas & ":Z" & ulttramoa1), Type:=xlFillDefault
    Range("Z" & primfilatarjetas & ":Z" & ulttramoa1).Select
    wb2_s3.Range("AA" & primfilatarjetas).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AA" & primfilatarjetas & ":AA" & ulttramoa1), Type:=xlFillDefault
    Range("AA" & primfilatarjetas & ":AA" & ulttramoa1).Select
    wb2_s3.Range("AB" & primtramoa1).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AB" & primtramoa1 & ":AB" & ulttramoa1), Type:=xlFillDefault
    Range("AB" & primtramoa1 & ":AB" & ulttramoa1).Select
'TRAMOA2
    Range("Z" & primtramoa2).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-5])^(30/360)-1)*RC[-3]"
    Range("Z" & primtramoa2).Select
    Selection.AutoFill Destination:=Range("Z" & primtramoa2 & ":Z" & ulttramoa3)
    Range("Z" & primtramoa2 & ":Z" & ulttramoa3).Select
    Range("AA" & primtramoa2).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-6])^(RC[-13]/360)-1)*RC[-4]-RC[-1]"
    Range("AA" & primtramoa2).Select
    Selection.AutoFill Destination:=Range("AA" & primtramoa2 & ":AA" & ulttramoa2), Type:= _
        xlFillDefault
    Range("AA" & primtramoa2 & ":AA" & ulttramoa2).Select
    wb2_s3.Range("AB" & primtramoa2).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AB" & primtramoa2 & ":AB" & ulttramoa2), Type:=xlFillDefault
    Range("AB" & primtramoa2 & ":AB" & ulttramoa2).Select

'TRAMOA3
    Range("AA" & primtramoa3).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-6])^(90/360)-1)*RC[-4]-RC[-1]"
    Range("AA" & primtramoa3).Select
    Selection.AutoFill Destination:=Range("AA" & primtramoa3 & ":AA" & ulttramoa3)
    Range("AA" & primtramoa3 & ":AA" & ulttramoa3).Select
    Range("AB" & primtramoa3).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-7])^(RC[-6]/360)-1)*RC[-3]"
    Range("AB" & primtramoa3).Select
    Selection.AutoFill Destination:=Range("AB" & primtramoa3 & ":AB" & ulttramoa3)
    Range("AB" & primtramoa3 & ":AB" & ulttramoa3).Select

'COMPLETAR LOS ANTIGUOS
    wbs3_calculoss3.Activate
    wbs3_calculoss3.Columns("M:N").Select
    Range("M11166").Activate
    Application.CutCopyMode = False
    Selection.Copy
    wbs3_calculoss3.Columns("AP:AQ").Select
    Range("AP11166").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AP11196").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    
    wb2_s3.Activate
    wb2_s3.Range("AH" & ults3nuevos + 3).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-19],'[" & wbs3.Name & "]" & wbs3_calculoss3.Name & "'!R" & prims3tarjetas & "C15:R" & ults3tarjetas & "C43,28,0)"
    wb2_s3.Range("AH" & ults3nuevos + 3).Select
    Selection.AutoFill Destination:=Range("AH" & ults3nuevos + 3 & ":AH" & ultfilatarjetas + 2)
    wb2_s3.Range("AH" & ults3nuevos + 3 & ":AH" & ultfilatarjetas + 2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Range("AJ" & ults3nuevos + 3).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-21],'[" & wbs3.Name & "]" & wbs3_calculoss3.Name & "'!R" & prims3tarjetas & "C15:R" & ults3tarjetas & "C23,9,0)"
    wb2_s3.Range("AJ" & ults3nuevos + 3).Select
    Selection.AutoFill Destination:=Range("AJ" & ults3nuevos + 3 & ":AJ" & ultfilatarjetas + 2)
    wb2_s3.Range("AJ" & ults3nuevos + 3 & ":AJ" & ultfilatarjetas + 2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Range("AK" & ults3nuevos + 3).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-22],'[" & wbs3.Name & "]" & wbs3_calculoss3.Name & "'!R" & prims3tarjetas & "C15:R" & ults3tarjetas & "C43,29,0)"
    wb2_s3.Range("AK" & ults3nuevos + 3).Select
    Selection.AutoFill Destination:=Range("AK" & ults3nuevos + 3 & ":AK" & ultfilatarjetas + 2)
    wb2_s3.Range("AK" & ults3nuevos + 3 & ":AK" & ultfilatarjetas + 2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Range("AL" & ults3nuevos + 3).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-23],'[" & wbs3.Name & "]" & wbs3_calculoss3.Name & "'!R" & prims3tarjetas & "C15:R" & ults3tarjetas & "C23,7,0)"
    wb2_s3.Range("AL" & ults3nuevos + 3).Select
    Selection.AutoFill Destination:=Range("AL" & ults3nuevos + 3 & ":AL" & ultfilatarjetas + 2)
    wb2_s3.Range("AL" & ults3nuevos + 3 & ":AL" & ultfilatarjetas + 2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb2_s3.Range("AM" & ults3nuevos + 3).Select
    ActiveCell.FormulaR1C1 = "=RC[-25]-RC[-2]"
    Range("AM" & ults3nuevos + 3).Select
    Selection.AutoFill Destination:=Range("AM" & ults3nuevos + 3 & ":AM" & ultfilatarjetas + 2)
    wb2_s3.Range("AM" & ults3nuevos + 3 & ":AM" & ultfilatarjetas + 2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
  
'Intereses para antiguos clientes
    wb2_s3.Range("AG" & ults3nuevos + 3).Select
    ActiveCell.FormulaR1C1 = "=IF(RC[6]<30,""0.negativos"","" "")"
    wb2_s3.Range("AG" & ults3nuevos + 3).Select
    Selection.AutoFill Destination:=Range("AG" & ults3nuevos + 3 & ":AG" & ultfilatarjetas + 2)
    wb2_s3.Range("AG" & ults3nuevos & ":AG" & ultfilatarjetas + 2).Select
    wb2_s3.Range("AP" & ults3nuevos + 2).Select
    ActiveCell.FormulaR1C1 = "muestra"
    wb2_s3.Range("AP" & ults3nuevos + 3).Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-7],""NO UBICADOS POR DOC"")"
    wb2_s3.Range("AP" & ults3nuevos + 3).Select
    Selection.AutoFill Destination:=Range("AP" & ults3nuevos + 3 & ":AP" & ultfilatarjetas + 2)
    Range("AP" & ults3nuevos + 3 & ":AP" & ultfilatarjetas + 2).Select
    wb2_s3.Range("AP" & ults3nuevos + 3).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-9],""NO UBICADOS POR DOC"")"
    wb2_s3.Range("AP" & ults3nuevos + 3).Select
    Selection.AutoFill Destination:=Range("AP" & ults3nuevos + 3 & ":AP" & ultfilatarjetas + 2)
    Range("AP" & ults3nuevos + 3 & ":AP" & ultfilatarjetas + 2).Select
    Selection.Copy
    Range("AG" & ults3nuevos + 3).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    wb2_s3.Range("AP" & ults3nuevos + 3).Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-9]="" "",""SI LOS UBICA 31 DIAS"",RC[-9])"
    wb2_s3.Range("AP" & ults3nuevos + 3).Select
    Selection.AutoFill Destination:=Range("AP" & ults3nuevos + 3 & ":AP" & ultfilatarjetas + 2)
    Range("AP" & ults3nuevos + 3 & ":AP" & ultfilatarjetas + 2).Select
    Selection.Copy
    Range("AG" & ults3nuevos + 3).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AP" & ults3nuevos + 3).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Selection.AutoFilter
    Range("AP" & ults3nuevos + 2).Select
    Selection.ClearContents
    Range("AH" & ults3nuevos + 2).Select
    Selection.AutoFilter
    wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("AG" & ults3nuevos + 2 & ":AG" & ultfilatarjetas + 2), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Dim prim0negativo As Double
Dim primNoubicdoc As Double
Dim prim31dias As Double
Dim ult0negativo As Double
Dim ultNoubicdoc As Double
Dim ult31dias As Double

prim0negativo = wb2_s3.Columns("AG").Find("0.negativos", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print prim0negativo & "zz"
ult0negativo = wb2_s3.Columns("AG").Find("0.negativos", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ult0negativo & "zzzz"
primNoubicdoc = wb2_s3.Columns("AG").Find("NO UBICADOS POR DOC", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print pprimNoubicdoc & "XD1"
ultNoubicdoc = wb2_s3.Columns("AG").Find("NO UBICADOS POR DOC", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ultNoubicdoc & "XD2"
prim31dias = wb2_s3.Columns("AG").Find("SI LOS UBICA 31 DIAS", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print prim31dias & "XD3"
ult31dias = wb2_s3.Columns("AG").Find("SI LOS UBICA 31 DIAS", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ult31dias & "XD34"

'Intereses para los 0 negativos

    wb2_s3.Range("Z" & prim0negativo).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("Z" & prim0negativo & ":Z" & ult0negativo), Type:=xlFillDefault
    Range("Z" & prim0negativo & ":Z" & ult0negativo).Select
    wb2_s3.Range("AA" & prim0negativo).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AA" & prim0negativo & ":AA" & ult0negativo), Type:=xlFillDefault
    Range("AA" & prim0negativo & ":AA" & ult0negativo).Select
    wb2_s3.Range("AB" & prim0negativo).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AB" & prim0negativo & ":AB" & ult0negativo), Type:=xlFillDefault
    Range("AB" & prim0negativo & ":AB" & ult0negativo).Select
    
''Intereses para los no ubicados por documento
    
    wb2_s3.Range("$A$" & ults3nuevos + 2 & ":$AO$" & ultfilatarjetas + 2).AutoFilter Field:=33, Criteria1:= _
        "NO UBICADOS POR DOC"
    wb2_s3.AutoFilter.Sort.SortFields.Clear
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("N" & ults3nuevos + 2 & ":N" & ultfilatarjetas + 2), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    wb2_s3.Range("AP" & primNoubicdoc).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-28]<=30,""TRAMOB1"",IF(RC[-28]<=90,""TRAMOB2"",IF(RC[-28]<=120,""TRAMOB3"",""TRAMOB4"")))"
    Range("AP" & primNoubicdoc).Select
    Selection.AutoFill Destination:=Range("AP" & primNoubicdoc & ":AP" & ultNoubicdoc), Type:=xlFillDefault
    Range("AP" & primNoubicdoc & ":AP" & ultNoubicdoc).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
Dim primtramob1 As Double
Dim primtramob2 As Double
Dim primtramob3 As Double
Dim primtramob4 As Double
Dim ulttramob1 As Double
Dim ulttramob2 As Double
Dim ulttramob3 As Double
Dim ulttramob4 As Double


primtramob1 = wb2_s3.Columns("AP").Find("TRAMOB1", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print primtramob1 & "aea"
ulttramob1 = wb2_s3.Columns("AP").Find("TRAMOB1", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ulttramob1 & "aea1"
primtramob2 = wb2_s3.Columns("AP").Find("TRAMOB2", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print primtramob1 & "aea2"
ulttramob2 = wb2_s3.Columns("AP").Find("TRAMOB2", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ulttramob2 & "aea3"
primtramob3 = wb2_s3.Columns("AP").Find("TRAMOB3", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print primtramob3 & "aea4"
ulttramob3 = wb2_s3.Columns("AP").Find("TRAMOB3", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ulttramob3 & "aea5"
primtramob4 = wb2_s3.Columns("AP").Find("TRAMOB4", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print primtramob4 & "aea6"
ulttramob4 = wb2_s3.Columns("AP").Find("TRAMOB4", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print ulttramob4 & "aea7"
    Dim dias As Variant
    dias = Split(wb_interfaz.Range("S15").Value, "/")(0)
    
'TRAMO B1
    wb2_s3.Range("Z" & primtramob1).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-5])^(RC[-12]/360)-1)*RC[-3]"
    wb2_s3.Range("Z" & primtramob1).Select
    Selection.AutoFill Destination:=Range("Z" & primtramob1 & ":Z" & ulttramob1)
    Range("Z" & primtramob1 & ":Z" & ulttramob1).Select
    wb2_s3.Range("AA" & primtramob1).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AA" & primtramob1 & ":AA" & ulttramob1), Type:=xlFillDefault
    Range("AA" & primtramob1 & ":AA" & ulttramob1).Select
    wb2_s3.Range("AB" & primtramob1).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AB" & primtramob1 & ":AB" & ulttramob1), Type:=xlFillDefault
    Range("AB" & primtramob1 & ":AB" & ulttramob1).Select

 'TRAMO B2
 
    wb2_s3.Range("Z" & primtramob2).Select
    ActiveCell.FormulaR1C1 = "0"
    wb2_s3.Range("Z" & primtramob2).Select
    Selection.AutoFill Destination:=Range("Z" & primtramob2 & ":Z" & ulttramob4)
    Range("Z" & primtramob2 & ":Z" & ulttramob4).Select
    wb2_s3.Range("AA" & primtramob2).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-6])^(" & dias & "/360)-1)*RC[-4]"
    Selection.AutoFill Destination:=Range("AA" & primtramob2 & ":AA" & ulttramob2), Type:=xlFillDefault
    Range("AA" & primtramob2 & ":AA" & ulttramob2).Select
    wb2_s3.Range("AB" & primtramob2).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AB" & primtramob2 & ":AB" & ulttramob2), Type:=xlFillDefault
    Range("AB" & primtramob2 & ":AB" & ulttramob2).Select


 
 'TRAMO B3
    
    wb2_s3.Range("AA" & primtramob3).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-6])^(((90-(RC[-13]-" & dias & "))/360))-1)*RC[-4]"
    Selection.AutoFill Destination:=Range("AA" & primtramob3 & ":AA" & ulttramob3), Type:=xlFillDefault
    Range("AA" & primtramob3 & ":AA" & ulttramob3).Select
    wb2_s3.Range("AB" & primtramob3).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-7])^((" & dias & "-(90-(RC[-14]-" & dias & ")))/360)-1)*RC[-3]"
    Selection.AutoFill Destination:=Range("AB" & primtramob3 & ":AB" & ulttramob3), Type:=xlFillDefault
    Range("AB" & primtramob3 & ":AB" & ulttramob3).Select
 
 
 'TRAMO B4
    
    wb2_s3.Range("AA" & primtramob4).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AA" & primtramob4 & ":AA" & ulttramob4), Type:=xlFillDefault
    Range("AA" & primtramob4 & ":AA" & ulttramob4).Select
    wb2_s3.Range("AB" & primtramob4).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-7])^(" & dias & "/360)-1)*RC[-3]"
    Selection.AutoFill Destination:=Range("AB" & primtramob4 & ":AB" & ulttramob4), Type:=xlFillDefault
    Range("AB" & primtramob4 & ":AB" & ulttramob4).Select
    
    
 ''Intereses para los UBICA 31 DIAS
    
 
    wb2_s3.Range("$A$" & ults3nuevos + 2 & ":$AO$" & ultfilatarjetas + 2).AutoFilter Field:=33, Criteria1:= _
        "SI LOS UBICA 31 DIAS"
    wb2_s3.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("N30692:N49076"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With wb2_s3.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    wb2_s3.Range("AP" & prim31dias).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-28]<=30,""TRAMOC1"",IF(RC[-28]<=90,""TRAMOC2"",IF(RC[-28]<=120,""TRAMOC3"",""TRAMOC4"")))"
    Range("AP" & prim31dias).Select
    Selection.AutoFill Destination:=Range("AP" & prim31dias & ":AP" & ult31dias), Type:=xlFillDefault
    Range("AP" & prim31dias & ":AP" & ult31dias).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Dim primtramoc2 As Double
Dim primtramoc3 As Double
Dim primtramoc4 As Double

Dim ulttramoc2 As Double
Dim ulttramoc3 As Double
Dim ulttramoc4 As Double



primtramoc2 = wb2_s3.Columns("AP").Find("TRAMOC2", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print primC2 & "XD"
ulttramoc2 = wb2_s3.Columns("AP").Find("TRAMOC2", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print prims3tarjetas & "XD"
primtramoc3 = wb2_s3.Columns("AP").Find("TRAMOC3", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print prims3tarjetas & "XD"
ulttramoc3 = wb2_s3.Columns("AP").Find("TRAMOC3", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print prims3tarjetas & "XD"
primtramoc4 = wb2_s3.Columns("AP").Find("TRAMOC4", _
  searchorder:=xlByRows, searchdirection:=xlNext).Row
Debug.Print prims3tarjetas & "XD"
ulttramoc4 = wb2_s3.Columns("AP").Find("TRAMOC4", _
  searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Debug.Print prims3tarjetas & "XD"

  'TRAMO c2
 
    wb2_s3.Range("Z" & primtramoc2).Select
    ActiveCell.FormulaR1C1 = "0"
    wb2_s3.Range("Z" & primtramoc2).Select
    Selection.AutoFill Destination:=Range("Z" & primtramoc2 & ":Z" & ulttramoc4)
    Range("Z" & primtramoc2 & ":Z" & ulttramoc4).Select
    wb2_s3.Range("AA" & primtramoc2).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-6])^(RC[12]/360)-1)*RC[-4]"
    Selection.AutoFill Destination:=Range("AA" & primtramoc2 & ":AA" & ulttramoc2), Type:=xlFillDefault
    Range("AA" & primtramoc2 & ":AA" & ulttramoc2).Select
    wb2_s3.Range("AB" & primtramoc2).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AB" & primtramoc2 & ":AB" & ulttramoc2), Type:=xlFillDefault
    Range("AB" & primtramoc2 & ":AB" & ulttramoc2).Select
 
 'TRAMO c3

    wb2_s3.Range("AA" & primtramoc3).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-6])^((90-RC[10])/360)-1)*RC[-4]"
    Selection.AutoFill Destination:=Range("AA" & primtramoc3 & ":AA" & ulttramoc3), Type:=xlFillDefault
    Range("AA" & primtramoc3 & ":AA" & ulttramoc3).Select
    wb2_s3.Range("AB" & primtramoc3).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-7])^((RC[11]-(90-RC[9]))/360)-1)*RC[-3]"
    Selection.AutoFill Destination:=Range("AB" & primtramoc3 & ":AB" & ulttramoc3), Type:=xlFillDefault
    Range("AB" & primtramoc3 & ":AB" & ulttramoc3).Select
 
 
 'TRAMO c4
    
    wb2_s3.Range("AA" & primtramoc4).Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AA" & primtramoc4 & ":AA" & ulttramoc4), Type:=xlFillDefault
    Range("AA" & primtramoc4 & ":AA" & ulttramoc4).Select
    wb2_s3.Range("AB" & primtramoc4).Select
    ActiveCell.FormulaR1C1 = "=((1+RC[-7])^(RC[11]/360)-1)*RC[-3]"
    Selection.AutoFill Destination:=Range("AB" & primtramoc4 & ":AB" & ulttramoc4), Type:=xlFillDefault
    Range("AB" & primtramoc4 & ":AB" & ulttramoc4).Select
      
    

    wb2_s3.Columns("AP:AP").Select
    Selection.Delete Shift:=xlToLeft
    
    Selection.AutoFilter
    
    wb2_s3.Range("AH" & primfilatarjetas - 1).Select


    wb2_s3.Activate
    wb2_s3.Range("AH1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlToLeft
    
    wb2_s3.Range("AD1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlToLeft
    wb2_s3.Range("AE1:AF1").Select
    Selection.Delete Shift:=xlToLeft
    MsgBox " Se calcularon los intereses"
    
    wbs1.Close
    wbs3.Close
    wb2.Save
   wb_interfaz.Activate
End Sub
