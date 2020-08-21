Attribute VB_Name = "vbaBancaVazame"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
   FullScreenMode Desligado
   ThisWorkbook.Save
End Sub

Private Sub Workbook_Open()
    Dim sResult As String
    
    sResult = ValidateNonPayment()
    If sResult <> "" Then
        MsgBox "A " & sResult & " parcela está atrasada ou com status não pago., SISTEMA BLOQUEADO!" _
            , vbCritical, Me.Name
        'ThisWorkbook.Close savechanges:=True
    Else
        FullScreenMode Ligado
        MsgBox "Bem vindo ao sistema gerencial de lançamentos." _
            , vbInformation, Me.Name
    End If
End Sub


