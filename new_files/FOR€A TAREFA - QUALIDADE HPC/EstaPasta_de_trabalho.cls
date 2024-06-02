Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)

LOG "ARQUIVO FECHADO"

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

LOG "ARQUIVO SALVO"

End Sub


Private Sub Workbook_NewSheet(ByVal Sh As Object)

LOG "NOVA PLANILHA INSERIDA"

End Sub

Private Sub Workbook_Open()

Dim USER As String
USER = Environ("USERNAME")

      'AUTORIZADOS A VER A BARRA DE SHEETS
      
      Select Case USER
      
      Case "ijuni002", "jsilv061", "Junior&Cris"
      
      Plan5.Visible = True

Plan5.Unprotect ["ftqualidadehpi222"]
         
      ActiveWindow.DisplayWorkbookTabs = True
    
      
      Case Else
      
      ActiveWindow.DisplayWorkbookTabs = False
     
      
    Plan5.Protect ["ftqualidadehpi222"]

    Plan5.Visible = False
      
      End Select

      LOG "ARQUIVO ABERTO"

Worksheets("HOME").Activate
Range("A1").Select

        If Hour(Now()) > 0 And Hour(Now()) <= 12 Then
            MsgBox "Bom dia " & Environ("USERNAME") & "! Seja bem vindo(a)!"
        End If
        
        If Hour(Now()) > 12 And Hour(Now()) < 18 Then
            MsgBox "Boa tarde " & Environ("USERNAME") & "! Seja bem vindo(a)!"
        End If
        
        If Hour(Now()) >= 18 Then
            MsgBox "Boa noite " & Environ("USERNAME") & "! Seja bem vindo(a)!"
        End If
        
End Sub

