Attribute VB_Name = "frmProfissionais"
Attribute VB_Base = "0{29FF193D-1A98-4A72-8DFB-01A4B8C3D554}{181F6201-3211-477A-9A25-3BBA7D4046C9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub btn_back_Click()
     Unload Me
     frmMain.Show
End Sub

Private Sub btn_cadastrar_Click()
     Dim ws As Worksheet
     Dim lo As ListObject
     Dim lr As ListRow
     
    With Me
          If .txt_prof.Value = "" Then
               .lbValida.Caption = "#ERRO PROF = VAZIO"
               .lbValida.ForeColor = &HFF&
           ElseIf .txt_cod = "" Then
                .lbValida.Caption = "#ERRO CÓD. = VAZIO"
               .lbValida.ForeColor = &HFF&
           ElseIf .txt_cbo = "" Then
                .lbValida.Caption = "#ERRO CBO. = VAZIO"
               .lbValida.ForeColor = &HFF&
           Else
                Set ws = shListas
                Set lo = ws.ListObjects("LISTA_PROCED")
                Set lr = lo.ListRows.Add
                With lr
                     .Range(1).Value2 = Me.txt_prof.Value
                     .Range(2).Value2 = Me.txt_cod.Value * 1
                     .Range(3).Value2 = Me.txt_cbo.Value * 1
                End With
          .txt_prof.Value = ""
          .txt_cod.Value = ""
          .txt_cbo.Value = ""
          .lbValida = ""
          MsgBox "Cadastro Realizado com Sucesso", vbInformation, .Caption
          End If
      End With
End Sub

Private Sub btn_voltar_Click()
     Unload Me
     frmMain.Show
End Sub

Private Sub txt_cbo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
     Me.txt_cbo.Value = Format(Me.txt_cbo.Value, "000000")
End Sub

Private Sub txt_cod_Exit(ByVal Cancel As MSForms.ReturnBoolean)
     Me.txt_cod.Value = Format(Me.txt_cod.Value, "0000000000")
End Sub

Private Sub txt_prof_Exit(ByVal Cancel As MSForms.ReturnBoolean)
     Me.txt_prof.Value = UCase(Me.txt_prof.Value)
End Sub

