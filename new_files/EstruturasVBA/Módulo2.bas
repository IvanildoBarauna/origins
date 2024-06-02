Attribute VB_Name = "Módulo2"
Option Explicit

Sub YourConfidence()
      Dim reply As Integer
      
      reply = MsgBox("Você confia nas suas próprias forças?", vbQuestion + vbYesNo, "YourConfidence")
      
      If reply = vbYes Then
            MsgBox "Maldito o homem que confia no homem!", vbExclamation
      Else
            MsgBox "Mantenha sua confiança sempre em Deus", vbInformation
      End If
      
End Sub
