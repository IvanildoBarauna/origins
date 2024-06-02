Attribute VB_Name = "wbk"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()

Dim User As String
      User = Environ("USERNAME")
Dim Validar As String
      Validar = Range("Validação").Value
Dim booLogin      As Boolean
      booLogin = Validar
      
      Range("Usuário_Logado").Value = User
      shtUSR.Calculate
           
If booLogin Then

      MsgBox "Seja bem vindo!", vbInformation, "Validação de Usuários"
      
Else

      MsgBox "Usuário não inserido na lista de autorizados.", vbCritical, "Acesso negado"

      'ThisWorkbook.Close False   Comentei esta linha para caso vc não esteja na lista, não fosse barrado e não conseguisse abrir a wbk
      
End If

End Sub
