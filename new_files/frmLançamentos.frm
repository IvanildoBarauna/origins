Attribute VB_Name = "frmLançamentos"
Attribute VB_Base = "0{3B8D75A9-E6F5-4818-BAAE-CC478C2AC338}{6085E063-A600-4574-85DD-B6DCE2A3F5AD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub btn_voltar_Click()
     Unload Me
     frmMain.Show
End Sub

Private Sub UserForm_Initialize()
     Dim ws As Worksheet
     Dim lo As ListObject
     Dim lr As Integer
     Dim i As Integer
     
     Set ws = shListas
     Set lo = ws.ListObjects("LISTA_PROCED")
     lr = lo.ListRows.Count + 1
          
     For i = 2 To lr
          cbo_prof.AddItem ws.Cells(i, 1).Value2
     Next i
     Me.btnLan.Enabled = False
     Me.txt_databpa = DateSerial(Year(Date), Month(Date) - 1, 21)
End Sub
Private Sub cbo_prof_Exit(ByVal Cancel As MSForms.ReturnBoolean)
     With Me
          If .cbo_prof.Value = "" Then
                    .lbValida.Caption = "#ERRO PROF = VAZIO"
                    .lbValida.ForeColor = &HFF&
                    .btnLan.Enabled = False
          Else
               .lbValida = ""
          End If
     End With
End Sub

Private Sub txt_nascto_Exit(ByVal Cancel As MSForms.ReturnBoolean)
     With Me
          If .txt_nascto = "" Then
               .lbValida.ForeColor = &HFF&
               .lbValida = "#ERRO DATA NASCTO = VAZIO"
               Cancel = True
               .btnLan.Enabled = False
          ElseIf Len(Replace(.txt_nascto, "/", "")) <> 8 Then
               .lbValida.ForeColor = &HFF&
               .lbValida = "#ERRO DATA NASCTO = INVÁIDA."
               Cancel = True
               txt_nascto.Value = ""
          Else
               .lbValida = ""
               .btnLan.Enabled = True
               .txt_nascto.Value = Format(Replace(Me.txt_nascto.Value, "/", ""), "00\/00\/0000")
          End If
     End With
End Sub

Private Sub txt_databpa_Exit(ByVal Cancel As MSForms.ReturnBoolean)
     With Me
     If .txt_databpa = "" Then
               .lbValida.ForeColor = &HFF&
               .lbValida = "#ERRO DATA BPA = VAZIO"
               Cancel = True
               .btnLan.Enabled = False
          ElseIf Len(Replace(.txt_databpa, "/", "")) <> 8 Then
               .lbValida.Caption = "#ERRO DATA BPA = INVÁLIDA"
               .txt_databpa.Value = ""
               Cancel = True
               .lbValida.ForeColor = &HFF&
               .btnLan.Enabled = False
          Else
               .lbValida = ""
               .btnLan.Enabled = True
               .txt_databpa.Value = Format(Replace(.txt_databpa, "/", ""), "00\/00\/0000")
          End If
     End With
End Sub

Private Sub btnLan_Click()
Dim ws As Worksheet
Dim lo As ListObject
Dim lr As ListRow

     With Me
          If .cbo_prof.Value = "" Then
               .lbValida.Caption = "ERRO CAMPO PROFISSIONAL"
               .btnLan.Enabled = False
               ElseIf .txt_nascto.Value = "" Then
                    .lbValida.Caption = "ERRO CAMPO DATA NASCTO."
                    .btnLan.Enabled = False
               ElseIf .txt_databpa.Value = "" Then
                    .lbValida.Caption = "ERRO CAMPO DATA BPA"
                    .btnLan.Enabled = False
               Else
                    .btnLan.Enabled = True
                    Set ws = shDados
                    Set lo = ws.ListObjects("DIGITAÇÃO")
                    Set lr = lo.ListRows.Add
                    With lr
                         .Range(1).Value2 = Me.cbo_prof.Value
                         .Range(2).Value2 = Replace(Me.txt_nascto.Value, "/", "")
                         .Range(3).Value2 = Replace(Me.txt_databpa.Value, "/", "")
                         Me.txt_nascto.Value = ""
                         Me.cbo_prof.Value = ""
                         Me.cbo_prof.SetFocus
                         Me.lbValida.ForeColor = &HFF00&
                         Me.lbValida = "FICHA LANÇADA!"
                    End With
          End If
     End With
End Sub
