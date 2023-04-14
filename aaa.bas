Sub Culminarexcelriesgos()

Dim Dialogo1 As FileDialog
'nFilas=filas de la hoja DATA
'nColumnas=columnas de la hoja DATA
Dim nFilas, nColumnas As String
'ultimafila2=filas de la hoja CTA_CTBL
Dim ultimafila2 As String
  Dim t As Single 'Inicia el cronómetro
  t = Timer
  'NUESTRO CÓDIGO
Dim algo As Boolean
Dim nombre As String
Dim wb As Workbook
Dim wb_interfaz As Worksheet
Dim ArchivoBCSeleccionado As Variant

Set wb = ThisWorkbook
Set wb_interfaz = wb.Sheets(1)
'Abrir cuadro para escoger el BC

Set Dialogo1 = Application.FileDialog(msoFileDialogFilePicker)
algo = False
Dialogo1.Title = "Escoger Archivo Balance de Comprobacion"
If Dialogo1.Show = -1 Then

    For Each ArchivoBCSeleccionado In Dialogo1.SelectedItems
        
        wb_interfaz.Range("F11").FormulaR1C1 = ArchivoBCSeleccionado
        On Error GoTo Adios
        Workbooks.Open Filename:= _
        ArchivoBCSeleccionado
    
    If ArchivoBCSeleccionado <> " " Then
        algo = True
    End If
    
    
    Next ArchivoBCSeleccionado
    
Else
MsgBox "No se ha escogido el archivo "
    If algo = False Then
        GoTo Despedida
    End If


End If
' Seteo los archivos


Dim nFilas1, ultimafila As Double
'ultimafila2=filas de la hoja CTA_CTBL


'wb1 = libro destino riesgos
Dim wb1 As Workbook
Dim wb1_data As Worksheet
Dim wb1_ctactbl As Worksheet
Dim wb1_tipo As Worksheet
'wb2 =libro origen _bc
Dim wb2 As Workbook
Dim wb2_saldosmen As Worksheet


wb.Activate
Set wb1 = Workbooks.Open(wb_interfaz.Range("F7").Value)
wb.Activate


Set wb2 = Workbooks.Open(wb_interfaz.Range("F11").Value)

'abrir el archivo de riesgos

' Seteo las hojas
Set wb1_data = wb1.Worksheets("DATA")
Set wb1_ctactbl = wb1.Worksheets("CTA_CTBL")
Set wb2_saldosmen = wb2.Worksheets("SaldosMensuales")
    
    
    nFilas1 = wb1_data.Range("A" & Rows.Count).End(xlUp).Row
    
    ultimafila = wb1_ctactbl.Range("A" & Rows.Count).End(xlUp).Row
    
    

    'Para la concatenacion CT_CTBL
    wb1_ctactbl.Activate
    wb1_ctactbl.AutoFilterMode = False
    wb1_ctactbl.Range("E1").Value = "CODIGO"
    wb1_ctactbl.Range("F1").Value = "CUENTA"
    wb1_ctactbl.Range("G1").Value = "NOMBRE CUENTA"
    wb1_ctactbl.Range("E2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-3],RC[-2])"
    wb1_ctactbl.Range("E2").Select
    Selection.AutoFill Destination:=wb1_ctactbl.Range("E2:E" & ultimafila)
    
    wb1_ctactbl.Range("E2:E" & ultimafila).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    'para la izquierda
    wb1_ctactbl.Range("F2").Select
    ActiveCell.FormulaR1C1 = "=Left(RC[-2],4)"
    wb1_ctactbl.Range("F2").Select
    Selection.AutoFill Destination:=wb1_ctactbl.Range("F2:F" & ultimafila)
    
    wb1_ctactbl.Range("F1:F" & ultimafila).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False


    
    
    'Para la CONCATENACION DATA
    wb1_data.Activate
    'Completar los inputs
    '
    wb1_data.AutoFilterMode = False
    wb1_data.Columns("L:L").Select
    Selection.Copy
    wb1_data.Columns("O:O").Select
    wb1_data.Paste
    Application.CutCopyMode = False
    wb1_data.Columns("D:D").Select
    Selection.Copy
    wb1_data.Columns("N:N").Select
    wb1_data.Paste
    Application.CutCopyMode = False
    'completar los demas datos
    wb1_data.Range("P2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-15],RC[-14])"
    wb1_data.Range("P2").Select
    Selection.AutoFill Destination:=wb1_data.Range("P2:P" & nFilas1)
    
    wb1_data.Range("P1:P" & nFilas1).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    

    wb1_data.Range("P1").Value = "CODIGO"
    wb1_data.Range("Q1").Value = "CUENTA"
    wb1_data.Range("R1").Value = "TIPO"
    wb1_data.Range("S1").Value = "NOMBRE CUENTA"


      
      'Para la BUSQUEDA CT_CTBL
      'wb2.name& "]"&wb2_saldosmen.name
    
    wb1_ctactbl.Activate
    wb1_ctactbl.Range("G2").FormulaR1C1 = _
        "=VLOOKUP(RC[-3],[" & wb2.Name & "]" & wb2_saldosmen.Name & "!C10:C11,2,0)"
    wb1_ctactbl.Range("G2").Select
    
    Selection.AutoFill Destination:=wb1_ctactbl.Range("G2:G" & ultimafila)
        
    wb1_ctactbl.Range("G1:G" & ultimafila).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    wb2.Close
      
      'Para la BUSQUEDA DATA CUENTA
      'wb2.name& "]"&wb2_saldosmen.name
    
    wb1_data.Activate
    
    wb1_data.Range("Q2").FormulaR1C1 = _
        "=VLOOKUP(RC[-1]," & wb1_ctactbl.Name & "!C[-12]:C[-10],2,FALSE)"
    wb1_data.Range("Q2").Select
  ' "=VLOOKUP(RC[-1],[" & wb1.Name & "]" & wb1_ctactbl.Name & "!C[-12]:C[-10],2,FALSE)"
    Selection.AutoFill Destination:=wb1_data.Range("Q2:Q" & nFilas1)
        
    wb1_data.Range("Q1:Q" & nFilas1).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

      'Para la BUSQUEDA DATA TIPO
      'wb2.name& "]"&wb2_saldosmen.name
    
    wb1_data.Activate
    wb1_data.Range("S2").FormulaR1C1 = _
        "=VLOOKUP(RC[-3]," & wb1_ctactbl.Name & "!C[-14]:C[-11],3,0)"
    wb1_data.Range("S2").Select
    
    Selection.AutoFill Destination:=wb1_data.Range("S2:S" & nFilas1)
        
    wb1_data.Range("S1:S" & nFilas1).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
      'Para la BUSQUEDA TIPO CUENTA
    wb1_data.Columns("Q:Q").Select
    Selection.TextToColumns Destination:=Range("Q1"), DataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=False, Tab:=True, Semicolon _
        :=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, _
        1), TrailingMinusNumbers:=True
    

    
    wb1_data.Select
    wb1_data.Range("R2").Select
    wb1_data.Range("R2").FormulaR1C1 = "=VLOOKUP(RC[-1],[" & wb.Name & "]" & wb_interfaz.Name & "!R14C6:R20C7,2,0)"
    wb1_data.Range("R2").Select
    
    Selection.AutoFill Destination:=wb1_data.Range("R2:R" & nFilas1)
    wb1_data.Range("R1:R" & nFilas1).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

'Ordenar
    wb1_ctactbl.Activate
    wb1_ctactbl.Range("F1").Select
    Selection.AutoFilter

    wb1_ctactbl.AutoFilter.Sort.SortFields.Clear
    wb1_ctactbl.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("F1:F" & ultimafila), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With wb1_ctactbl.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With



    Dim bFileSaveAs As Boolean
    bFileSaveAs = Application.Dialogs(xlDialogSaveAs).Show
    
    Dim rutawb3 As Variant
    Dim wb3 As Workbook
    Set wb3 = ActiveWorkbook
    
    rutawb3 = wb3.Path
    wb.Activate
    wb_interfaz.Activate
    wb_interfaz.Range("O7").FormulaR1C1 = rutawb3 & "\" & wb3.Name
    
    Debug.Print nombre
    


  '"=VLOOKUP(RC[-1],[" & wb1.Name & "]" & wb1_ctactbl.Name & "!C[-12]:C[-10],2,0)"
GoTo Despedida
Adios:
    If ActiveWorkbook.Name <> nombre Then

    ActiveWorkbook.Close
    End If

    MsgBox "Error encontrado"

Despedida:
End Sub


