Attribute VB_Name = "frmCadClientes"
Attribute VB_Base = "0{C3C8A410-6FDD-4FD8-81B4-93DDDD8CEF2D}{2B87E5B9-0E22-40ED-B20F-D20A3847CDF4}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private booNewReg As Boolean
Private Sub btnCancela_Click()
     Unload Me
End Sub

Public Sub btnSalvar_Click()
     ValidaCampos Me
     SalvarDados_SemTabela shClientes, Me, "cad_clientes"
End Sub

Private Sub txtPesquisa_Enter()
     If txtPesquisa.Value = "PESQUISE AQUI ..." Then txtPesquisa.Value = ""
End Sub

Private Sub txtPesquisa_Exit(ByVal Cancel As MSForms.ReturnBoolean)
     If txtPesquisa.Value = vbNullString Then txtPesquisa.Value = "PESQUISE AQUI ..."
End Sub

