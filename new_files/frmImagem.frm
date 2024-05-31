Attribute VB_Name = "frmImagem"
Attribute VB_Base = "0{94F49245-D9C0-4ED8-84EE-FC7386309DC2}{BC718099-7820-47C7-AE3A-3DD72AD48EA9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub cmdSair_Click()
    Me.Hide
End Sub

Private Sub imgSQL_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Hide
End Sub

Private Sub UserForm_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    frmImagem.Caption = "Tabelas da base " & frmSQL.txtArquivoDados
End Sub
