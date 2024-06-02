Attribute VB_Name = "Módulo1"
Option Explicit

Sub t()

'      Dim lo As ListObject
      Dim ws As Worksheet
'
      Set ws = ActiveSheet
'      Set lo = ws.ListObjects("tblBASE")

      Dim usrprofile    As String
      usrprofile = Environ("USERPROFILE") & "\Desktop\"
      
      Application.DisplayAlerts = 0
      Application.ScreenUpdating = 0

'      Workbooks.Open Filename:= _
'            usrprofile & "BACKLOG.xls"
'
'      Range("A2:P2", Selection.End(xlDown)).Copy
'
'      Windows(ThisWorkbook.Name).Activate
'
'      shtBACKLOG.Range("A5").PasteSpecial xlPasteValuesAndNumberFormats
'
'      ws.Columns.EntireColumn.AutoFit
'
'      ActiveWindow.ActivateNext
'
'      ActiveWindow.Close False
'
      Workbooks.Open Filename:= _
            usrprofile & "SLA.xls"
            
      Range("A2:M2", Selection.End(xlDown)).Copy
      
      Windows(ThisWorkbook.Name).Activate
      
      shtSLA.Range("A5").PasteSpecial xlPasteValuesAndNumberFormats
      
      ws.Columns.EntireColumn.AutoFit
      
      ActiveWindow.ActivateNext
      
      ActiveWindow.Close False
      
      MsgBox "Concluído!", vbInformation
      
End Sub



