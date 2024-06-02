Attribute VB_Name = "Módulo1"
Sub Macro2()

i = 4
Do While (i < 107)
If Plan11.Range("B" & i) = 0 Then
        Plan11.Rows(i).Hidden = True
    Else: Plan11.Rows(i).Hidden = False
End If
i = i + 1
Loop


End Sub


Sub Macro1()

i = 4
Do While (i < 107)
If Plan12.Range("B" & i) = 0 Then
        Plan12.Rows(i).Hidden = True
    Else: Plan12.Rows(i).Hidden = False
End If
i = i + 1
Loop


End Sub
