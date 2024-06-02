Attribute VB_Name = "mdMainAux"
Option Explicit
Public Const SW_SHOW As Integer = 1
Public Const SW_SHOWMAXIMIZED  As Integer = 3
Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" _
                      (ByVal hWnd As Long, _
                       ByVal lpOperation As String, _
                       ByVal lpFile As String, _
                       ByVal lpParameters As String, _
                       ByVal lpDirectory As String, _
                       ByVal nShowCmd As Long) As Long
                       
Public booAux   As Boolean
Public Const sCredits As String = "Todos os Direitos Reservados à Ivanildo Junior | Contato: +55 11 940758369, ivanildo.jnr@outlook.com"
                                     
Public Enum RotineMode
    Desligado = 0
    Ligado = 1
End Enum
    
Public Sub btn_caixa(): shCaixa.Select: End Sub
Public Sub btn_cedulas(): shContagem.Select: End Sub
Public Sub btn_pedidos(): shPedidos.Select: End Sub
Public Sub btn_apoio(): sApoio.Select: End Sub
Public Sub btn_exit(): ThisWorkbook.Close SaveChanges:=True: End Sub
Public Sub AbrirForm(): frmlançamentos.Show: End Sub
Public Sub AbrirCalculadora(): Shell "CALC.EXE": End Sub
Public Sub BackToHome(): shCaixa.Select: End Sub

Public Sub btn_relatorio()
    Dim fDialog     As FileDialog
    Dim sArquivo    As String
    Dim sPath       As String
    
    If MsgBox("O arquivo atual será fechado, salvo e o painel do PowerBI será aberto," & _
        "deseja prosseguir?", vbQuestion + vbYesNo) = vbYes Then
        
        Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
        sPath = ThisWorkbook.Path
        ThisWorkbook.Save
            
        With fDialog
            .Title = "Abrir Relatório do PowerBI"
            .Filters.Add "Arquivos do PowerBI", "*.PBIX"
            .InitialFileName = sPath
             If .Show Then
                sArquivo = .SelectedItems(1)
                Call ShellExecute(0, "open", sArquivo, "", _
                              sPath, SW_SHOWMAXIMIZED)
                ThisWorkbook.Close
                Application.Quit
            Else
                MsgBox "Nenhum arquivo foi selecionado, operação cancelada", _
                    vbCritical, "Abrir Relatório do PBI"
            End If
        End With
    End If
End Sub

Public Sub AlternateFullScreen()
    If booAux Then ModoTelaCheia Desligado Else ModoTelaCheia Ligado
End Sub

Public Sub ClearData()
    Dim ws  As Worksheet
    Dim lo  As ListObject
    
    Set ws = shContagem
    Set lo = ws.ListObjects(1)
    
    If lo.ListRows.Count < 2 Then
        MsgBox "Não há dados para serem apagados.", vbExclamation
    Else
        With lo.ListRows(2)
            .Application.Range(.Range(1, 1), _
                .Range(lo.ListRows.Count - 1, lo.ListColumns.Count)).Rows.Delete
            lo.Application.Range(lo.DataBodyRange(1, 1), lo.DataBodyRange(1, 2)).ClearContents
            lo.DataBodyRange(1, 1).Select
        End With
        MsgBox "Valores reiniciados", vbInformation
    End If
End Sub
Public Sub ModoTelaCheia(ByVal Status As RotineMode)
    Dim booConfig As Boolean
    
    booConfig = Not VBA.IIf(Status = Ligado, True, False)
    booAux = Not booConfig
    
    With Application
        .ScreenUpdating = False
        .DisplayFullScreen = Not booConfig
        .DisplayFormulaBar = booConfig
        .DisplayScrollBars = booConfig
        .DisplayStatusBar = booConfig
        .Caption = VBA.IIf(Not booConfig, sCredits, VBA.vbNullString)
        With .ActiveWindow
            .DisplayHeadings = booConfig
            .DisplayWorkbookTabs = booConfig
            .DisplayVerticalScrollBar = booConfig
            .DisplayHorizontalScrollBar = booConfig
        End With
        .ScreenUpdating = True
    End With
End Sub
