Sub ExportarArchivoparaTI()
Dim Dialogo As FileDialog
Dim ArchivoRIESGOSSeleccionado As Variant
Dim fso As FileSystemObject
Dim ArchivoRIESGOS As Workbook
Dim NombreArchivo, RutaArchivo As String
Dim txtArchivo As Scripting.TextStream
'nFilas=filas de la hoja DATA
'nColumnas=columnas de la hoja DATA
Dim nFilas, nColumnas As String
'ultimafila2=filas de la hoja CTA_CTBL
Dim ultimafila2 As String
Dim ultimafila3 As String
Dim algo As Boolean
Dim nombre As String
Dim i, j As Integer
Dim wb As Workbook
Dim t As Single
Dim wb_interfaz As Worksheet

'Inicia el cronómetro
t = Timer
  'NUESTRO CÓDIGO
Set wb = ThisWorkbook
Set wb_interfaz = ThisWorkbook.Sheets(1)
nombre = wb.Name
Set Dialogo = Application.FileDialog(msoFileDialogFilePicker)
algo = False
Dialogo.Title = "Escoger Archivo a Exportar"
If Dialogo.Show = -1 Then

    For Each ArchivoRIESGOSSeleccionado In Dialogo.SelectedItems
        
        wb_interfaz.Range("F7").FormulaR1C1 = ArchivoRIESGOSSeleccionado
        On Error GoTo Adios
        Workbooks.Open Filename:= _
        ArchivoRIESGOSSeleccionado
    

    If ArchivoRIESGOSSeleccionado <> " " Then
        algo = True
    End If
    
    Next ArchivoRIESGOSSeleccionado
Else

 MsgBox "No se ha escogido el archivo "
    If algo = False Then
        GoTo Despedida
    End If
End If


    NombreArchivo = ThisWorkbook.Sheets("INTERFAZ").Cells(6, 6).Value
    RutaArchivo = ActiveWorkbook.Path & "\" & NombreArchivo & ".txt"
    
    Set fso = New FileSystemObject
    Set txtArchivo = fso.CreateTextFile(RutaArchivo, True)
    On Error GoTo Adios
    
    nFilas = ActiveWorkbook.Worksheets("DATA").Range("A" & Rows.Count).End(xlUp).Row
    nColumnas = ActiveWorkbook.Worksheets("DATA").Cells(1, Columns.Count).End(xlToLeft).Column
    ultimafila2 = Sheets("CTA_CTBL").Cells(Rows.Count, 1).End(xlUp).Row
    'i=Filas
    'j=Columnas
    
    
    For i = 2 To nFilas
        strCelda = ""
        For j = 1 To 1
            strCelda = strCelda & ActiveWorkbook.Worksheets("DATA").Cells(i, j).Text
        Next j
        txtArchivo.WriteLine Right(strCelda, Len(strCelda))
    Next i
    
    txtArchivo.Close
    

 'Para la IZQUIERDA CT_CTBL

'wb1 = libro origen RIESGOS
Dim wb1 As Workbook
Set wb1 = Workbooks.Open(wb_interfaz.Range("F7").Value)
' Seteo las hojas
Dim wb1_data As Worksheet
Dim wb1_ctactbl As Worksheet

Set wb1_data = wb1.Worksheets("DATA")
Set wb1_ctactbl = wb1.Worksheets("CTA_CTBL")

    wb1_ctactbl.Activate
    wb1_ctactbl.Range("F2").Select
    ActiveCell.FormulaR1C1 = "=Left(RC[-2],4)"
    wb1_ctactbl.Range("F2").Select
    Selection.AutoFill Destination:=wb1_ctactbl.Range("F2:F" & ultimafila2)
    
    wb1_ctactbl.Range("F1:F" & ultimafila2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    wb1_ctactbl.Range("F1").Value = "TIPO"
 'Para completar las condiciones
    wb1_ctactbl.Columns("F:F").Select
    wb1_ctactbl.Range("$F$1:$F$" & ultimafila2).RemoveDuplicates Columns:=1, Header:= _
        xlNo
    wb1_ctactbl.Range("F2").Select
    wb1_ctactbl.Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    wb.Activate
    wb_interfaz.Range("F14").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    wb1.Activate
    wb1_ctactbl.Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    wb.Activate
    ultimafila3 = wb_interfaz.Range("F" & Rows.Count).End(xlUp).Row
    wb_interfaz.Range("F14:F" & ultimafila3).Select
    Selection.TextToColumns Destination:=Range("F14"), DataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=False, Tab:=True, Semicolon _
        :=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, _
        1), TrailingMinusNumbers:=True
        
    Application.DisplayAlerts = False
    wb1.Close
    wb_interfaz.Activate
    
    wb_interfaz.Range("G14").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(INTERFAZ!RC[-1],CRITERIOS!C[-6]:C[-5],2)"
    Selection.AutoFill Destination:=Range("G14:G" & ultimafila3)
    wb_interfaz.Range("G14:G" & ultimafila3).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    wb_interfaz.Range("A1").Select

 

GoTo Despedida
Adios:
    If ActiveWorkbook.Name <> nombre Then

    ActiveWorkbook.Close
    End If

    MsgBox "Error encontrado"

Despedida:

End Sub
