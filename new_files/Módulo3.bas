Attribute VB_Name = "Módulo3"
Public Enum FullScreenMode
    Ligar = 1
    Desligar = 0
End Enum

Public Sub FullScreen(eMode As FullScreenMode)
    Dim booAux      As Boolean
    Dim sbooAux     As String
    
    booAux = Not VBA.IIf(eMode = Ligar, True, False)
    sbooAux = VBA.IIf(booAux, "True", "False")
    
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon""," & sbooAux & ")"
    
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
End Sub
