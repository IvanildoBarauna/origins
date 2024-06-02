Attribute VB_Name = "frmPrint"
Attribute VB_Base = "0{B490E190-2164-4113-BA63-6EE2F411660B}{00581CA3-B429-4607-8311-E0AD932DBE39}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub btn_imprimir_Click()
     If shDados.Range("A6").Value = vbNullString Then
          MsgBox "Não há dados para impressão.", vbExclamation, Me.Caption
     Else
          With shDyn
               .PivotTables("dyn_bpa").PivotFields("ANO").CurrentPage = Me.txt_ano.Value * 1
               .PivotTables("dyn_bpa").PivotFields("MÊS").CurrentPage = Me.cbo_mês.Value
               .PivotTables("dyn_bpa").PivotCache.Refresh
               .Visible = True
               .PrintOut
               .Visible = False
               MsgBox "O relatório atualizado foi enviado para a fila de impressão da impressora padrão!" & _
                    vbNewLine & vbNewLine & "Verifique o arquivo impresso.", vbInformation, Me.Caption
          End With
     End If
     Unload Me
     frmMain.Show
End Sub

Private Sub btn_voltar_Click()
     Unload Me
     frmMain.Show
End Sub

Private Sub UserForm_Initialize()
Dim i As Integer
     With Me
          .txt_ano.Value = Format(Date, "YYYY")
          .cbo_mês.Value = StrConv(Format(Date, "MMMM"), vbProperCase)
          For i = 1 To 12
               .cbo_mês.AddItem StrConv(MonthName(i), vbProperCase)
          Next i
     End With
End Sub
