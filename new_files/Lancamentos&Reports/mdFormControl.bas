Attribute VB_Name = "mdFormControl"
Option Explicit
Public Enum Report
    Consultas = 0
    Procedimentos = 1
End Enum
Public Function ValidateEmptyControls(ByRef FRM As UserForm) As Boolean
    Dim xControl As MSForms.Control
    Dim sList    As String
    
    For Each xControl In FRM.Controls
        Select Case TypeName(xControl)
            Case "TextBox", "ComboBox"
                If xControl.Value = vbNullString Or xControl.Value = 0 Then
                    If Not ValidateEmptyControls Then _
                        ValidateEmptyControls = True
                    sList = sList & vbNewLine & xControl.Tag
                End If
        End Select
    Next xControl
    
    If ValidateEmptyControls Then MsgBox "Preencha os campos abaixo:" _
        & vbNewLine & sList, vbExclamation
End Function

Public Sub ClearFields(FRM As MSForms.UserForm)
    Dim xField As MSForms.Control

    For Each xField In FRM.Controls
        Select Case VBA.TypeName(xField)
            Case "TextBox", "ComboBox"
                If xField.Name <> "txt_databpa" Then xField.Value = vbNullString
        End Select
    Next xField
    
End Sub

Public Sub ExportReport(ReportName As Report, Year As String, Month As String, GeraPDF As XlYesNoGuess)
    Dim ws      As Excel.Worksheet
    Dim pTable  As Excel.PivotTable
    Dim fDialog As Office.FileDialog
    Dim xPath   As String
    
    On Error GoTo err
    'Set ws = VBA.IIf(ReportName = Consultas, wsReportConsultas, wsReportProcedimentos)
    Set pTable = ws.PivotTables(1)
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)

    Application.ScreenUpdating = False

    With pTable
        .ClearAllFilters
        .RefreshTable
        .PivotCache.Refresh
        .PivotFields("YEAR").ClearAllFilters
        .PivotFields("MONTH").ClearAllFilters
        .PivotFields("YEAR").CurrentPage = Year
        .PivotFields("MONTH").CurrentPage = Month
    End With
    
    If GeraPDF = xlYes Then
        With fDialog
            .Title = "Salvar o relatório PDF em ..."
            .ButtonName = "Salvar aqui"
            If .Show Then
            xPath = .SelectedItems(1) & Application.PathSeparator & ws.Name & "-" & Month & Year
                With ws
                    .ExportAsFixedFormat Type:=xlTypePDF, Filename:=xPath, _
                        Quality:=xlQualityStandard, IgnorePrintAreas:=True, OpenAfterPublish:=True
                End With
            Else
                MsgBox "Nenhum local foi selecionado, operação cancelada", vbExclamation
            End If
        End With
    Else: ws.PrintOut
    End If
    
    Application.ScreenUpdating = True
    Call ClearAllFilterinPivotTables
    MsgBox "Relatório exportado com sucesso.", vbInformation, "[CONCLUÍDO]"
    Exit Sub
    
err: MsgBox "Erro ao inserir dados na tabela dinâmica.", vbCritical, "Exportação de Relatório"
     Call ClearAllFilterinPivotTables
     Application.ScreenUpdating = True
End Sub

Public Sub SortListObject(lo As ListObject, iCol As Integer, Order As XlSortOrder, Header As XlYesNoGuess)
    With lo
        .Sort.SortFields.Clear
        .Sort.SortFields.Add _
            Key:=lo.ListColumns(iCol).Range, SortOn:=xlSortOnValues, _
            Order:=Order, DataOption:=xlSortNormal
        With .Sort
            .Header = Header
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
End Sub
 
Function ColumnsWidhtsToListBox(DataList) As String
    Dim iRow    As Integer
    Dim iCol    As Integer
    Dim arr, item, Widths
    Dim nCaract As Integer
    
    arr = DataList
    ReDim Widths(1 To UBound(arr, 2)) As String
    
    For iCol = 1 To UBound(arr, 2)
        For iRow = 1 To UBound(arr, 1)
            item = arr(iRow, iCol)
            If VBA.Len(item) Then
                Widths(iCol) = VBA.Len(item)
            End If
        Next iRow
    Next iCol
    
End Function

Sub ClearAllFilterinPivotTables()
    Dim pTable As Excel.PivotTable
    Dim ws     As Excel.Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        For Each pTable In ws.PivotTables
            pTable.ClearAllFilters
        Next pTable
    Next ws
End Sub

Public Function StartDateBPA() As Date
    Dim vMes As Integer
    
    Select Case VBA.Day(VBA.Date)
        Case 1 To 20
            vMes = VBA.Month(VBA.Date) - 1
        Case Else
            vMes = VBA.Month(VBA.Date)
    End Select

    StartDateBPA = VBA.DateSerial(VBA.Year(VBA.Date), vMes, 21)
End Function
