Attribute VB_Name = "mdAux"
Option Explicit
Option Private Module
Public Enum RotineType
    Desligado = 0
    Ligado = 1
End Enum

Public Const sPass As String = "INKMASD876T673KN5656762"

Public Sub CallForm(): frmLancamentos.Show: End Sub
Public Sub GoToHome(): wsHome.Select: End Sub
Public Sub GoToDados(): wsDados.Select: End Sub
Public Sub GoToReport(): wsReport.Select: End Sub

Public Sub DateToPivotTable()
On Error GoTo Err
    Dim ws          As Worksheet
    Dim pTable      As PivotTable
    Dim YearField   As PivotField
    Dim MonthField  As PivotField
    Dim DayField    As PivotField
    Dim wDayField   As PivotField
    
    Set ws = wsReport
    Set pTable = ws.PivotTables("tbDyn_ForProduct")
    
    wsReport.PivotTables("tbDyn_ForMonth").PivotCache.Refresh
    
    With pTable
        Set YearField = .PivotFields("ANO")
        Set MonthField = .PivotFields("MÊS")
        Set DayField = .PivotFields("DIA")
        Set wDayField = .PivotFields("DIA_SEMANA")
    
        .PivotCache.Refresh
        .ClearAllFilters
        
        YearField.CurrentPage = VBA.DateTime.Year(Date)
        MonthField.CurrentPage = VBA.MonthName(VBA.DateTime.Month(Date))
        DayField.CurrentPage = VBA.DateTime.Day(Date)
        wDayField.CurrentPage = VBA.WeekdayName(VBA.DateTime.Weekday(Date))
        
        MsgBox "Relatório atualizado com os dados de: " & Date & " (" & StrConv(VBA.WeekdayName(Weekday(Date), True), vbProperCase) & ")" & vbNewLine & vbNewLine & _
                "Última atualização: " & .RefreshDate, vbInformation, wsReport.Name
    Exit Sub
Err:
    MsgBox "Erro ao atualizar o relatório, verifique se há dados lançados hoje.", vbCritical
    End With
End Sub

Public Sub ClearAllFiltersInPivotTable()
    wsReport.PivotTables("tbDyn_ForProduct").ClearAllFilters
End Sub

Public Sub FullScreenMode(Status As RotineType)
    Dim booConfig As Boolean
    Dim ws        As Worksheet
    Dim sName     As String
    
    sName = ActiveSheet.Name
    
    booConfig = Not VBA.IIf(Status = Ligado, True, False)
    
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
        .DisplayStatusBar = booConfig
        .DisplayStatusBar = booConfig
        .DisplayFormulaBar = booConfig
        For Each ws In .ThisWorkbook.Sheets
            ws.Activate
            With .ActiveWindow
                .DisplayHeadings = booConfig
                .DisplayWorkbookTabs = booConfig
                .DisplayVerticalScrollBar = booConfig
                .DisplayHorizontalScrollBar = booConfig
            End With
        Next ws
        .ScreenUpdating = True
        .ThisWorkbook.Sheets(sName).Select
        .EnableEvents = True
    End With
End Sub

Public Sub ClearFields(ByRef FRM As MSForms.UserForm)
    Dim Ctrl As Control
    
    For Each Ctrl In FRM.Controls
        Select Case VBA.TypeName(Ctrl)
            Case "TextBox", "ComboBox"
                Ctrl.Value = vbNullString
        End Select
    Next Ctrl
    
End Sub

Public Function ErrRaise(ByVal sCustomMsg As String) As String
    ThisWorkbook.Save
    
    With VBA.Err
        ErrRaise = sCustomMsg & vbNewLine & vbNewLine & _
                    "NÚMERO DO ERRO: " & .Number & vbNewLine & vbNewLine & _
                    "DESCRIÇÃO DO ERRO: " & .Description
    End With
    
    MsgBox ErrRaise, vbCritical
End Function

Public Function EmptyFields(ByRef FRM As MSForms.UserForm) As Boolean
    Dim xControl As Control, sFields As String
    
    For Each xControl In FRM.Controls
        If VBA.TypeName(xControl) = "TextBox" Or _
            VBA.TypeName(xControl) = "ComboBox" Then
            If xControl.Value = "" Then
            sFields = sFields & vbNewLine & "Campo: " & xControl.Tag
            EmptyFields = True
            End If
        End If
    Next xControl
    
    If EmptyFields Then MsgBox "Favor preencher os seguintes campos:" & _
        vbNewLine & sFields, vbExclamation
End Function

Public Function Filtermtz(ByVal mtz, _
                          ByVal iCol As Integer)
    Dim mtzResult   As Variant
    Dim index       As Long
    Dim RowCounter  As Long
    Dim ColCounter  As Integer
    Dim mtzSize     As Long
    Dim mtzValue    As Date
    Dim ValueDate   As Date
    
    For index = LBound(mtz, 1) To UBound(mtz, 1)
        On Error Resume Next
        mtzValue = mtz(index, iCol)
        On Error GoTo 0
        ValueDate = DateSerial(Year(mtzValue), Month(mtzValue), Day(mtzValue))
        If ValueDate = Date Then mtzSize = mtzSize + 1
    Next index
    
    mtzValue = 0
    ValueDate = 0
    
    ReDim mtzResult(1 To mtzSize + 1, 1 To 4) As String
    
    mtzResult(1, 1) = "ID"
    mtzResult(1, 2) = "DATA"
    mtzResult(1, 3) = "PRODUTO"
    mtzResult(1, 4) = "VALOR"
    
    For index = LBound(mtz, 1) To UBound(mtz, 1)
        On Error Resume Next
        mtzValue = mtz(index, iCol)
        On Error GoTo 0
        ValueDate = DateSerial(Year(mtzValue), Month(mtzValue), Day(mtzValue))
        If ValueDate = Date Then
            RowCounter = RowCounter + 1
            For ColCounter = 2 To UBound(mtzResult, 2)
                mtzResult(RowCounter + 1, 1) = index
                mtzResult(RowCounter + 1, ColCounter) = mtz(index, ColCounter)
            Next ColCounter
        End If
    Next index
    
    Filtermtz = FormatArray(mtzResult, "Currency", _
        wsDados.ListObjects(1).ListColumns("VALOR").index)
    
End Function

Public Function FormatArray(ByVal mtz, _
                            ByVal sFormat As String, _
                            ByVal iCol As Integer)
    Dim RowCounter  As Long
    Dim ColCounter  As Long
    Dim mtzResult   As Variant
    
    mtzResult = mtz
    
    For RowCounter = LBound(mtzResult, 1) To UBound(mtzResult, 1)
        For ColCounter = LBound(mtzResult, 2) To UBound(mtzResult, 2)
            If ColCounter = iCol Then
                mtzResult(RowCounter, ColCounter) = Format(mtzResult(RowCounter, ColCounter), sFormat)
            End If
        Next ColCounter
    Next RowCounter
    
    FormatArray = mtzResult
End Function

Public Function ValidateNonPayment() As String
    Dim iCounter As Integer
    Dim vDate    As Date
    Dim ws       As Excel.Worksheet
    Dim lo       As ListObject
    Dim vStatus  As String
    
    Set ws = wsPagamentos
    Set lo = ws.ListObjects("tbPagamentos")
    
    For iCounter = 1 To lo.ListRows.Count
        vDate = lo.DataBodyRange(iCounter, _
            lo.ListColumns("DATA").index).Value2 + 1
        vStatus = lo.DataBodyRange(iCounter, _
            lo.ListColumns("Valida").index).Value2
        If vDate <= VBA.Date And vStatus = "NÃO PAGO" Then
            ValidateNonPayment = lo.DataBodyRange(iCounter, _
                lo.ListColumns("ID").index).Text
            Exit Function
        End If
    Next iCounter
    
End Function

Public Sub FullAccessControl(ByVal mode As RotineType)
    If mode = Ligado Then
        FullScreenMode Desligado
        Application.EnableEvents = False
        wsPagamentos.Unprotect sPass
    Else
        Application.EnableEvents = True
        FullScreenMode Ligado
        wsPagamentos.Protect sPass
    End If
End Sub
