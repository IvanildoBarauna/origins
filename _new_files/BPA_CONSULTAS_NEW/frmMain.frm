Attribute VB_Name = "frmMain"
Attribute VB_Base = "0{48514690-F42F-449D-B33C-E0FB876F45CD}{59C69820-D65E-416C-83EF-51F2B558F8A9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub btn_frmprof_Click()
     frmMain.Hide
     frmProfissionais.Show
End Sub

Private Sub btn_lançamentos_Click()
     frmMain.Hide
     frmLançamentos.Show
End Sub

Private Sub btn_print_Click()
     frmMain.Hide
     frmPrint.Show
End Sub

Private Sub btn_sair_Click()
     shDados.Activate
     Unload Me
End Sub
Private Sub btn_save_Click()
     ThisWorkbook.Save
     MsgBox "Arquivo Salvo com sucesso.", vbInformation, Me.Caption
End Sub

Private Sub UserForm_Initialize()
Dim iData As Date
Dim fData As Date

iData = DateSerial(Year(Date), Month(Date) - 1, 21)
fData = DateSerial(Year(Date), Month(Date), 20)
     Me.lbperiodo = "De: " & iData & " Á " & fData
     shApoio.Activate
     ModoApp True
End Sub

Private Sub UserForm_Terminate()
     ModoApp False
     shDados.Activate
End Sub
