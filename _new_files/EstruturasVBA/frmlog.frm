Attribute VB_Name = "frmlog"
Attribute VB_Base = "0{836935E7-FCFD-4F48-B69F-5A0731F4023B}{F18D701C-C838-4C1A-AA37-3FE6857D36E8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit


Private Sub cmdcancel_Click()
      Unload Me
End Sub

Private Sub cmdvalida_Click()
      Dim usuario As String
      Dim senha As String
      Dim ultln As Long
      Dim i As Long
      
      Application.ScreenUpdating = False
      
      ultln = shtLST.Range("A1047576").End(xlUp).Row
      shtLST.Select
      shtLST.Range("A2").Select
      
      For i = 1 To ultln
            usuario = ActiveCell.Value
            senha = ActiveCell.Offset(0, 1).Value
            If usuario = Me.txtusr And senha = Me.txtpass Then
                  Unload Me
                  MsgBox "Autorizado com sucesso", vbInformation
                  sht.Select
                  Exit Sub
            ElseIf ActiveCell.Row > ultln Then
                  MsgBox "Login não autorizado, reveja sua credencial de acesso", vbCritical
                  Me.txtusr = Empty
                  Me.txtpass = Empty
                  Me.txtusr.SetFocus
            Else
                  ActiveCell.Offset(1, 0).Select
            End If
      Next
      Application.ScreenUpdating = True
End Sub
