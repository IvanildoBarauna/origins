Attribute VB_Name = "twbk"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
     If Cancel Then
          ModoApp False
          Application.StatusBar = ""
          Me.Save
     End If
End Sub

Private Sub Workbook_Open()
Dim strUname As String
Dim strCname  As String
Dim strBar        As String

strUname = Environ("Username")
strCname = Environ("computername")
strBar = "Todos os direitos reservados à Ivanildo Junior.  |  Nome do Usuário logado: " _
               & strUname & "  |  Nome do Computador: " & strCname

     If strUname = "cristiane.barauna" Or strUname = "Junior&Cris" Then
          Application.StatusBar = strBar
          frmMain.Show
     Else
          MsgBox "Usuário não autorizado." & vbNewLine _
               & "Entre em contato com o administrador da planilha para conceder permissão" & vbNewLine _
                    & vbNewLine & strBar, vbCritical, "Validação"
          ThisWorkbook.Close SaveChanges:=False
     End If
End Sub
