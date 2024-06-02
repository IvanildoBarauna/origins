Attribute VB_Name = "mdAux"
Option Private Module
Option Explicit
Public booAux   As Boolean
Public Const sCredits As String = _
    "Todos os Direitos Reservados à Ivanildo Junior | Contato: (11 949982337), ivanildo.jnr@outlook.com"
Public Enum RotineMode
    Desligado = 0
    Ligado = 1
End Enum

' Modulo    : mdAux / MóduloComum
' Rotina    : ModoTelaCheia(Status = Ligado / Desligado) / Sub
' Autor     : Ivanildo Junior (ivanildo.jnr@outlook.com)
' Data      : 25/11/2017 22:39
' Proposta  : Deixar a aplicação em tela cheia, desativando todos as barras e menus.
'---------------------------------------------------------------------------------------

Public Sub ModoTelaCheia(ByVal Status As RotineMode)
    Dim booConfig As Boolean
    Dim sRiboon   As String
    Dim sCount    As Integer
    
    booConfig = VBA.IIf(Status = 1, True, False)
    booAux = booConfig
    booConfig = Not booConfig
    sRiboon = VBA.IIf(booConfig, "True", "False")
    
    With Application
        .ScreenUpdating = False
        .DisplayFullScreen = Not booConfig
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon""," & sRiboon & ")"
        .DisplayFormulaBar = booConfig
        .DisplayScrollBars = booConfig
        .DisplayStatusBar = booConfig
        .Caption = VBA.IIf(Not booConfig, sCredits, VBA.vbNullString)
        With .ActiveWindow
            For sCount = 1 To ThisWorkbook.Sheets.Count
                If ThisWorkbook.Sheets(sCount).Visible = False Then GoTo NextSheet
                ThisWorkbook.Sheets(sCount).Activate
                .DisplayGridlines = booConfig
                .DisplayHeadings = booConfig
                .DisplayWorkbookTabs = booConfig
                .DisplayVerticalScrollBar = booConfig
                .DisplayHorizontalScrollBar = booConfig
                .WindowState = xlMaximized
NextSheet:  Next sCount
        End With
        .ScreenUpdating = True
        .Sheets(1).Activate
    End With
End Sub

' Modulo    : mdAux / MóduloComum
' Rotina    : ToogleFullScreenMode / Sub
' Autor     : Ivanildo Junior (ivanildo.jnr@outlook.com)
' Data      : 25/11/2017 22:39
' Proposta  : Alternar a aplicação entre ativado e desativado para a rotina ModoTelaCheia
'---------------------------------------------------------------------------------------

Public Sub ToogleFullScreenMode()
    If booAux Then
        ModoTelaCheia Desligado
    Else
        ModoTelaCheia Ligado
    End If
End Sub

' Modulo    : mdAux / MóduloComum
' Rotina    : LIMPAR_MEMORIA() / Sub
' Autor     : Jefferson Dantas (jefferson@tecnun.com.br)
' Data      : 07/11/2012 - 16:42
' Revisão   : Fernando Fernandes (fernando@tecnun.com.br)
' Data      : 07/01/2013 (mdy)
' Revisão   : Ivanildo Junior (ivanildo.jnr@outlook.com)
' Data      : 25/11/2017 21:57 (dmy)
' Proposta  : Remove objetos da memória
'---------------------------------------------------------------------------------------

Public Sub Limpar_Memoria(ParamArray Objects() As Variant)
    On Error GoTo TRATAR_ERRO
    Dim Counter As Integer
    
    For Counter = 0 To UBound(Objects) Step 1
        Select Case TypeName(Objects(Counter))
            Case "Boolean"
                Objects(Counter) = False
            Case "Variant"
                If VBA.IsArray(Objects(Counter)) Then Erase Objects(Counter)
                Objects(Counter) = Empty
            Case "String"
             Objects(Counter) = vbNullString
            Case "Worksheet"
                Set Objects(Counter) = Nothing
            Case "Workbook"
                Objects(Counter).Close SaveChanges:=False
            Set Objects(Counter) = Nothing
                Case "Connection", "Database", "Recordset2", "Recordset"
                Objects(Counter).Close
            Set Objects(Counter) = Nothing
            
            Case Else
            
            Set Objects(Counter) = Nothing
            If VBA.IsObject(Objects(Counter)) Then
                Set Objects(Counter) = Nothing
            Else
                Objects(Counter) = Empty
            End If
        End Select
    Next Counter
    Exit Sub
    On Error GoTo 0
TRATAR_ERRO:
    Resume Next
    Debug.Print "ERRO NA LIMPEZA DE MEMÓRIA: " & VBA.TypeName(Objects(Counter))
End Sub


Public Function GetColumnOfListObjetc(ByVal ws As Worksheet, ByVal sTabela As String, ByVal sColumn As String) As Long
    On Error GoTo ErrRaise
    Dim lo As ListObject
    
    Set lo = ws.ListObjects(sTabela)
    
    GetColumnOfListObjetc = Application.WorksheetFunction.Match(sColumn, lo.HeaderRowRange, 0)
    Exit Function
ErrRaise:
    GetColumnOfListObjetc = -1
End Function

Public Function GetRowOfListObjetc(ByVal ws As Worksheet, ByVal sTabela As String, ByVal sColumn As String, ByVal Value As Variant)
    Dim lo       As Excel.ListObject
    Dim ColumnID As Long
    Dim rngDATA  As Excel.Range
    
    On Error GoTo ErrRaise
    
    Set lo = ws.ListObjects(sTabela)
    ColumnID = lo.ListColumns(sColumn).Index
    
    Set rngDATA = lo.ListColumns(ColumnID).DataBodyRange
    
    GetRowOfListObjetc = Application.WorksheetFunction.Match(Value, rngDATA, 0)
    Exit Function
ErrRaise:
    GetRowOfListObjetc = -1
 End Function
