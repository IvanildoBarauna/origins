Attribute VB_Name = "Módulo1"
Option Explicit
Public Enum RotineMode
    desligado = 0
    Ligado = 1
End Enum
Public Sub CriarLOG(ByVal sAcao As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    
    Set ws = shLOG
    Set lo = ws.ListObjects("tbLOG")
    Set lr = lo.ListRows.Add
    
    With lr
        .Range(lo.ListColumns("DATA").Index).Value2 = VBA.Format(VBA.Now, "dd/mm/yyyy hh:mm:ss")
        .Range(lo.ListColumns("COMPUTADOR").Index).Value2 = VBA.Environ("computername")
        .Range(lo.ListColumns("USUARIO").Index).Value2 = VBA.Environ("username")
        .Range(lo.ListColumns("REGISTRO").Index).Value2 = sAcao
    End With
    ws.Columns.AutoFit
End Sub

Public Sub CriarLOG2(ByVal sAcao As String)
    Dim ws As Worksheet
    Dim LastRow As Long
    
    Set ws = shLOG

    With ws
         LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        .Cells(LastRow, 1).Value2 = Now
        .Cells(LastRow, 2).Value2 = VBA.Environ("computername")
        .Cells(LastRow, 3).Value2 = VBA.Environ("username")
        .Cells(LastRow, 4).Value2 = sAcao
    End With
End Sub

Public Sub FullScreenMode(Status As RotineMode)
'Rotina original de Fábio Gatti com adaptações conforme necessidade do projeto atual
    Dim sCount  As Integer
    Dim sRibbon As String
    Dim booConf As Boolean
    
    booConf = IIf(Status = 1, True, False)
    sRibbon = VBA.IIf(Not booConf, "True", "False")
    
    With Application
        .ScreenUpdating = False
        .DisplayFullScreen = booConf
        .DisplayFormulaBar = Not booConf
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon""," & sRibbon & ")"
        .DisplayStatusBar = Not booConf
        .DisplayScrollBars = Not booConf
        For sCount = 1 To ThisWorkbook.Sheets.Count
            ThisWorkbook.Sheets(sCount).Activate
            If ThisWorkbook.Sheets(sCount).Visible = False Then GoTo NextSheet
            With .ActiveWindow
                .DisplayWorkbookTabs = Not booConf
                .DisplayHorizontalScrollBar = Not booConf
                .DisplayVerticalScrollBar = Not booConf
                .DisplayHeadings = Not booConf
            End With
NextSheet:
        Next sCount
        .ScreenUpdating = False
    End With
End Sub

