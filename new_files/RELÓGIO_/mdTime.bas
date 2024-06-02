Attribute VB_Name = "mdTime"
Option Explicit
Dim Ligado As Boolean

Sub Relógio()
    If Ligado Then
        ActiveSheet.Calculate
        Application.OnTime Now() + TimeValue("00:00:01"), "Relógio"
    End If
End Sub

Sub Ligar()
    Ligado = True
    Call Relógio
End Sub

Sub Desligar()
    Ligado = False
End Sub


