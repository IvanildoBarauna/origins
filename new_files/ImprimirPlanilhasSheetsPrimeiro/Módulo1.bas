Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub PrintSheets()
    Dim ws      As Worksheet
    Dim oCol    As Collection
    Dim item    As Integer
    
    Set oCol = New Collection
    
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.Visible Then ws.Visible = True
        If Not ws.Index Mod 2 = 0 Then
            ws.PrintOut
        Else
            oCol.Add ws.Name
        End If
    Next ws
    
    For item = 1 To oCol.Count
        ThisWorkbook.Sheets(oCol.item(item)).PrintOut
    Next item
End Sub

Public Function GetPlanName() As String
    GetPlanName = ActiveSheet.Name
End Function
