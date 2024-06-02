Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub Main()
    Const ScreenTip As String = "Clique para entrar no grupo"
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim iCounter As Integer
    Dim item As String
    
    Set ws = Planilha1
    Set lo = ws.ListObjects(1)
    
    For iCounter = 2 To lo.ListRows.Count
        item = lo.DataBodyRange(iCounter, 3).Value2
        ws.Hyperlinks.Add lo.DataBodyRange(iCounter, 3), _
                          lo.DataBodyRange(iCounter, 3).Value2, , ScreenTip & ": " & _
                          lo.DataBodyRange(iCounter, 2).Value2, _
                          lo.DataBodyRange(iCounter, 2).Value2
    Next iCounter
End Sub

Public Sub Main2()
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    
    ws.Hyperlinks.Add Selection, ws.Range("H7").Value, , "teste", "teste2"
End Sub
