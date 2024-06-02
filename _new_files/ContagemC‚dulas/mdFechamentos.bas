Attribute VB_Name = "mdFechamentos"
Option Explicit
Private Sub InserirFechamento()
    Dim lo      As Excel.ListObject
    Dim lr      As Excel.ListRow
    Dim xCtrl   As String, VendaReal, VendaEsperada As Double
    
    Application.ScreenUpdating = False
    xCtrl = wsMain.Range("Type").Value2
    VendaReal = wsMain.Range("Venda").Value2
    
    If xCtrl = "VENDA" And VendaReal > 0 Then
        If VBA.MsgBox("Tem certeza que deseja realizar o lançamento com os valores atuais?", vbQuestion + vbYesNo) = vbYes Then
            VendaEsperada = VendaTotal()
            If VendaEsperada <= 0 Then Exit Sub
            Set lo = wsFechamentos.ListObjects(1)
            Set lr = lo.ListRows.Add
            With lr
                .Range(lo.ListColumns("Data").Index).Value2 = VBA.DateTime.Date
                .Range(lo.ListColumns("VendaReal").Index).Value2 = VendaReal
                .Range(lo.ListColumns("VendaEsperada").Index).Value2 = VendaEsperada
                .Range(lo.ListColumns("Perda").Index).Value2 = IIf(VendaEsperada > VendaReal, VendaReal - VendaEsperada, 0)
                MsgBox "Fechamento realizado com sucesso.", vbInformation
            End With
        End If
    End If
    Application.ScreenUpdating = True
End Sub

Private Function isOpenWb(wbFullName As String) As Excel.Workbook
    Dim wb      As Excel.Workbook
    Dim booAux  As Boolean
    
    For Each wb In Excel.Application.Workbooks
        If wb.FullName = wbFullName Then
            booAux = True
            Exit For
        End If
    Next wb
    
    If booAux Then
        Set isOpenWb = Application.Workbooks(GetName(wbFullName))
    Else
        Application.EnableEvents = False
        Set isOpenWb = Application.Workbooks.Open(wbFullName)
        ThisWorkbook.Sheets(1).Activate
        Application.EnableEvents = True
    End If
    
End Function

Private Function VendaTotal() As Double
    Const sPath  As String = "D:\OneDrive\wbSalesManager.xlsm"
    Dim iCounter As Long
    Dim wb       As Excel.Workbook
    Dim ws       As Excel.Worksheet
    Dim lo       As Excel.ListObject
    Dim vSum     As Double
    
    Set wb = isOpenWb(sPath)
    
    If Not wb Is Nothing Then
        Set ws = wb.Sheets("Lançamentos")
        Set lo = ws.ListObjects(1)
        For iCounter = 1 To lo.ListRows.Count
            If lo.DataBodyRange(iCounter, lo.ListColumns("DataLançamento").Index).Value2 = VBA.DateTime.Date And _
                lo.DataBodyRange(iCounter, lo.ListColumns("Lançamento").Index).Value2 = "RECEITA" Then
                vSum = vSum + lo.DataBodyRange(iCounter, lo.ListColumns("Venda").Index).Value2
            End If
        Next iCounter
        VendaTotal = vSum
    End If
End Function

Private Function GetName(wbFullName As String) As String
    Dim wf            As WorksheetFunction: Set wf = Application.WorksheetFunction
    Dim ReplaceString As String
    Dim extString     As String
    Dim FinalString   As String
    
    ReplaceString = VBA.Replace(wbFullName, "\", "-", 1, 1)
    extString = VBA.Mid(ReplaceString, wf.Find("\", ReplaceString))
    GetName = VBA.Replace(extString, "\", "")
End Function

