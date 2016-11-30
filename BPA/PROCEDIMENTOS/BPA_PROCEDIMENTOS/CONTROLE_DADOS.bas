Attribute VB_Name = "CONTROLE_DADOS"
Public qtd As Long
Public acr As Long
Public at As Long
Public resp As String
Public proced As String
Public cod As Long
Public prof As String
Public cbo As Long
Public filtro As String
Public filtro2 As String
Public Const App As String = "| Gerencial BPA - UBS Santo Onofre 2017 |"

Sub ACRESCENTAR_QTD()

Application.ScreenUpdating = False

On Error GoTo ErroFiltro

filtro = InputBox("Digite o nome do procedimento que deseja acrescentar quantidade:", App)

ErroFiltro:

If filtro = Empty Or IsNumeric(filtro) = True Then

MsgBox "Dados inválidos, digite o nome do procedimento exatamente como mostra na lista de procedimentos..", vbCritical, App

Exit Sub

Else

On Error GoTo ErroFiltro2

filtro2 = InputBox("Digite o nome do profissional que deseja acrescentar procedimentos", App)

ErroFiltro2:

If filtro2 = Empty Or IsNumeric(filtro2) = True Then

MsgBox "Dados inválidos, digite o nome do profissional exatamente como mostra na lista de profissionais..", vbCritical, App

Exit Sub

Else

    ActiveSheet.ListObjects("tbDIGITAÇÃO").Range.AutoFilter Field:=2, Criteria1 _
        :=filtro
        
        ActiveSheet.ListObjects("tbDIGITAÇÃO").Range.AutoFilter Field:=1, Criteria1 _
        :=filtro2
        
        Range("A4").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(0, 4).Select
        at = ActiveCell.Value
        
On Error GoTo Erro1

acr = InputBox("Digite a quantidade a ser acrescentada em " & filtro & " para o profissional: " & filtro2 & ":", App)

Erro1:

If acr = Empty Or IsNumeric(acr) = False Then

MsgBox "Dados inválidos, digite a quantidade a ser acrescentada no procedimento.", vbCritical, App

Exit Sub

Else

ActiveCell.Value = at + acr

End If

End If

End If

        shtDIGITAÇÃO.ShowAllData

        Range("A4").Select
        Selection.End(xlDown).Select
        
        MsgBox "FORAM ADICIONADOS " & acr & " PROCEDIMENTOS DE " & UCase(filtro) & " PARA O PROFISSIONAL: " & filtro2 & "!", vbInformation, App

acr = Empty
filtro = Empty

Application.ScreenUpdating = True

End Sub

Sub LIMPEZA_DADOS()

Application.ScreenUpdating = False

resp = MsgBox("Atenção! Todos os dados anteriormente inseridos nesta tabela serão excluídos, deseja realmente iniciar uma nova digitação?", vbExclamation + vbYesNo, App)

If resp = vbYes Then

    Range("A6:E6").Select
    Range(Selection, Selection.End(xlDown)).Rows.Delete
    Range("A5:B5,E5").ClearContents
    
    Range("A5").Select

Else

Exit Sub

End If

Application.ScreenUpdating = True

End Sub

Sub INSERÇÃO_PROCEDIMENTO()

Application.ScreenUpdating = False

On Error GoTo Erro2

proced = UCase(InputBox("Insira o nome do procedimento:", App))

Erro2:

If proced = "" Or IsNumeric(proced) = True Then

    MsgBox "Nome do procedimento inválido.", vbCritical, App

    Exit Sub

Else

On Error GoTo Erro3

cod = InputBox("Insira o código do procedimento " & UCase(proced) & ":", App)

Erro3:

If cod = Empty Or IsNumeric(cod) = False Then

    MsgBox "Código do procedimento inválido.", vbCritical, App
    
    Exit Sub
    
    Else

shtPROCED.Visible = True
shtPROCED.Activate
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
ActiveCell = proced
ActiveCell.Offset(0, 1) = cod

    On Error Resume Next
    If Not Intersect(Target, Range("A:A")) Is Nothing Then
        Range("A1").Sort Key1:=Range("A2"), _
          Order1:=xlAscending, Header:=xlYes, _
          OrderCustom:=1, MatchCase:=False, _
          Orientation:=xlTopToBottom
    End If

shtDIGITAÇÃO.Activate

Application.ScreenUpdating = True

MsgBox "PROCEDIMENTO: " & UCase(proced) & " DE CÓDIGO: " & cod & " FOI INSERIDO COM SUCESSO!", vbInformation, App

End If

End If

proced = Empty
cod = Empty

End Sub

Sub ABRIR_PROCEDIMENTOS()

Application.ScreenUpdating = False

shtPROCED.Visible = True
shtPROCED.Activate
Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub ABRIR_PROFISSIONAIS()

Application.ScreenUpdating = False

shtPROF.Visible = True
shtPROF.Activate
Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub INSERÇÃO_PROFISSIONAL()

Application.ScreenUpdating = False

On Error GoTo Erro3

prof = UCase(InputBox("Insira o nome do profissional:", App))

Erro3:

If prof = "" Or IsNumeric(prof) = True Then

    MsgBox "Nome do profissional inválido.", vbCritical, App

    Exit Sub

Else

On Error GoTo Erro4

cbo = InputBox("Insira o número de CBO referente à: " & UCase(prof) & ":", App)

Erro4:

If cbo = Empty Or IsNumeric(cbo) = False Then

    MsgBox "Número de CBO do profissional inválido.", vbCritical, App
    
    Exit Sub
    
    Else

shtPROF.Visible = True
shtPROF.Activate
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
ActiveCell = prof
ActiveCell.Offset(0, 1) = cbo

    On Error Resume Next
    If Not Intersect(Target, Range("A:A")) Is Nothing Then
        Range("A1").Sort Key1:=Range("A2"), _
          Order1:=xlAscending, Header:=xlYes, _
          OrderCustom:=1, MatchCase:=False, _
          Orientation:=xlTopToBottom
    End If

shtDIGITAÇÃO.Activate

Application.ScreenUpdating = True

MsgBox "PROFISSIONAL: " & UCase(prof) & " NÚM CBO: " & cbo & " FOI INSERIDO COM SUCESSO!", vbInformation, App

End If

End If

prof = Empty
cbo = Empty

End Sub
