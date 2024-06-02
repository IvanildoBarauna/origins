Attribute VB_Name = "wbk"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Dim user            As String
Dim autorizaçao     As String
Dim ws              As Worksheet

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Application.Visible = True
End Sub

Private Sub Workbook_Open()

On Error GoTo N
    Dim lin         As Integer
    
    '    ThisWorkbook.Application.Visible = False
    MsgBox "Aqui aplicar o visible FALSE para ocultar o excel" & vbNewLine & _
        "Mas quando fechar ele volta ao normal!", vbInformation
        
    user = Environ("USERNAME")
    
    Set ws = ThisWorkbook.Worksheets(1)
        lin = 2
        
    Do Until ws.Range("a" & lin) = ""
        
        If ws.Cells(lin, 1) = user Then
        autorizaçao = ws.Cells(lin, 6)
            If autorizaçao = "Verdadeiro" Then
                MsgBox "Seja bem vindo " & user, vbInformation
                    ThisWorkbook.Application.Visible = True
                                Exit Sub
            Else
                MsgBox "Você não está autorizado para utilizar essa planilha!", vbExclamation
                GoTo N
                    'ThisWorkbook.Close False
                        Exit Sub
            End If
        End If
    lin = lin + 1
    Loop
    
    MsgBox "O usuário " & user & " não foi encontrado!", vbCritical
N:
         Set ws = ThisWorkbook.Worksheets(2)
                        lin = ws.UsedRange.Rows.Count + 1
                        
                        With ws
                            .Cells(lin, 1) = user
                            .Cells(lin, 2) = FormatDateTime(Now, vbShortDate)
                            .Cells(lin, 3) = FormatDateTime(Now, vbShortTime)
                        End With
End Sub

'----------------------------------------------------------------------------------------------------------------
'Private Sub Workbook_Open()
'
'Dim usr As String
'      usr = Environ("USERNAME")
'      Range("USERNAME") = usr
'Dim boologin As Boolean
'      boologin = shtAPOIO.Range("VALIDAR").Value
'
'If boologin Then
'
'      MsgBox "Seja bem vindo " & usr, vbInformation, "Acesso Autorizado"
'
'Else
'
'      MsgBox "Usuário não autorizado", vbCritical, "Acesso Negado"
'
'End If
'
'End Sub
