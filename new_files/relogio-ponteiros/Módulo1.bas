Attribute VB_Name = "Módulo1"

Dim Ligado As Boolean

Sub Relogio()

If Ligado Then
    ActiveSheet.Calculate
    Application.OnTime Now() + TimeValue("00:00:01"), "Relogio"
End If

End Sub

Sub LigarRelogio()

Ligado = True
Relogio

End Sub


Sub PararRelogio()

Ligado = False

End Sub
