Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    wsProdutos.Visible = xlSheetVeryHidden
    wsMarkups.Visible = xlSheetVeryHidden
    
    Me.Save
    Excel.Application.OnKey "^+~"
End Sub

Private Sub Workbook_Open()
    Dim strSaudacao     As String
    Dim ActualHour      As Integer
    
    ActualHour = VBA.Hour(VBA.Now)
    
    Select Case ActualHour
        Case 0 To 12
            strSaudacao = "OLÁ, BOM DIA!"
        Case 12 To 18
            strSaudacao = "OLÁ, BOA TARDE!"
        Case Else
            strSaudacao = "OLÁ, BOA NOITE!"
    End Select
    
    'ATALHO CTRL+SHIFT+ENTER
    
    Excel.Application.OnKey "^+~", "ShowProducts"
    MsgBox strSaudacao & vbNewLine & vbNewLine & "Seja muito bem vindo(a) a nossa Tabela de Preços, sinta-se á vontade para navegar entre os produtos e escolher a melhor opção para você e para sua casa.", vbInformation, "Tabela de Preços"
End Sub
