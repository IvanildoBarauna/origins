Attribute VB_Name = "mAux"
Option Explicit
Public Enum RotineMode
    Desligado = 0
    Ligado = 1
End Enum

Public Sub ModoApp(Mode As RotineMode)
    Dim booAux As Boolean
    
    booAux = Not VBA.IIf(Mode = Ligado, True, False)
    
    With Application
        .DisplayFullScreen = Not booAux
        .DisplayFormulaBar = booAux
        .DisplayStatusBar = booAux
        With .ActiveWindow
            .DisplayHeadings = booAux
            .DisplayWorkbookTabs = booAux
        End With
    End With
End Sub

