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
                       
Public booAux         As Boolean
Public Const sCredits As String = "Todos os Direitos Reservados à Ivanildo Junior | Contato: +55 11 940758369, ivanildo.jnr@outlook.com"

Public Enum ReturnyFormulaType
    NÚMERO = 0
    TEXTO = 1
End Enum
                                     
Public Enum RotineMode
    Desligado = 0
    Ligado = 1
End Enum
    
Public Sub btn_caixa(): shCaixa.Select: End Sub
Public Sub btn_pedidos(): shPedidos.Select: End Sub
Public Sub btn_apoio(): sApoio.Select: End Sub
Public Sub btn_exit(): ThisWorkbook.Close SaveChanges:=True: End Sub
Public Sub AbrirForm(): frmlançamentos.Show: End Sub
Public Sub AbrirCalculadora(): Shell "CALC.EXE": End Sub
Public Sub BackToHome(): shCaixa.Select: End Sub

Public Sub btn_relatorio()
    Dim fDialog     As Office.FileDialog
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

Public Sub ModoTelaCheia(Status As RotineMode)
    Dim booConfig As Boolean
    
    booConfig = Not VBA.IIf(Status = Ligado, True, False)
    booAux = Not booConfig
    
    With Application
        .DisplayFullScreen = Not booConfig
        .DisplayFormulaBar = booConfig
        .DisplayScrollBars = booConfig
        With .ActiveWindow
            .DisplayHeadings = booConfig
            .DisplayWorkbookTabs = booConfig
        End With
    End With
End Sub

Function GeraID(vDate As Date)
    Dim lo As Excel.ListObject
    Dim iCounter As Integer
    
    Set lo = shPedidos
    
    For iCounter = 1 To lo.ListRows.Count
        
    Next iCounter
End Function

Public Sub CorrectFormulaErros()
    Dim iCell   As Excel.Range
    Dim FX      As String

    For Each iCell In Selection
        If VBA.IsError(iCell.Value2) Then
            FX = VBA.Replace(iCell.FormulaLocal, "=", "")
            iCell.FormulaLocal = "=SEERRO(" & FX & ";)"
        End If
    Next iCell
    
End Sub
