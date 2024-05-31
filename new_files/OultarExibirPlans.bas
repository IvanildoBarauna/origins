Attribute VB_Name = "OultarExibirPlans"
Sub MostraTodasPlan()
 
  
 Dim Resultado As VbMsgBoxResult
    Resultado = MsgBox("Tem certeza que deseja exibir todas as guia?", vbYesNo, "Tomar uma descisão")
    
       
    If Resultado = vbYes Then
 
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "HOME" Then
            ws.Visible = xlSheetVisible 'Mostar todas
        End If
    Next
 
    Else
       ' MsgBox ("Ação Cancelada")
    End If
 
End Sub

Sub OcultarTodasPlan()
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "HOME" Then
            ws.Visible = xlSheetVeryHidden 'Ocultar todas
        End If
    Next
 Sheets("HOME").Select
 Range("M4").Select
End Sub


Sub Teste()
Application.Speech.Speak ("Semana ") & Folha7.Range("F5")
End Sub
