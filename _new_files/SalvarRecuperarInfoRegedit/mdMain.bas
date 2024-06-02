Attribute VB_Name = "mdMain"
Option Explicit
Public Const APPNAME   As String = "wbVendas"
Public Const mdName    As String = "mdMain"

Public Enum RotineMode
    DESLIGAR = 0
    LIGAR = 1
End Enum

Public Sub ModoTelaCheia(Mode As RotineMode)
    On Error GoTo err:
    Const RotineName As String = "ModoTelaCheia"
    Dim booAux       As Boolean
    Dim reginfo      As String
    
    booAux = Not VBA.IIf(Mode = LIGAR, True, False)
    reginfo = VBA.IIf(Mode = LIGAR, "ATIVO", "INATIVO")
    
    With Application
        .DisplayFullScreen = Not booAux
        .DisplayFormulaBar = booAux
        .DisplayScrollBars = booAux
        .DisplayStatusBar = booAux
        With .ActiveWindow
            .DisplayHeadings = booAux
            .DisplayWorkbookTabs = booAux
        End With
    End With
    
    VBA.SaveSetting APPNAME, mdName, RotineName, reginfo
    Exit Sub
err:
End Sub

Sub StatusAlternate()
    Dim reginfo As String
    
    reginfo = VBA.GetSetting(APPNAME, mdName, "ModoTelaCheia")
    
    If reginfo = "INATIVO" Then ModoTelaCheia LIGAR Else ModoTelaCheia DESLIGAR
End Sub
