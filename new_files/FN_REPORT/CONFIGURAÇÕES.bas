Attribute VB_Name = "CONFIGURAÇÕES"
Public Const AppName As String = "| Necessary Feedback Report |"
Public Const BarDefault As String = AppName & " 2017 - Todos os direitos Reservados. "
Option Private Module

Public Sub ACTIVATE_()

With Application

    .DisplayAlerts = False
    .ScreenUpdating = False
    .StatusBar = BarDefault

End With

End Sub

Public Sub DEACTIVATE_()

With Application

    .DisplayAlerts = True
    .ScreenUpdating = True
    .StatusBar = BarDefault

End With


End Sub


Public Sub ACTIVATE_APP()

With ActiveWindow

    .DisplayWorkbookTabs = False
    
End With

Application.StatusBar = BarDefault


End Sub

Public Sub DEACTIVATE_APP()

With ActiveWindow


    .DisplayWorkbookTabs = True

End With

Application.StatusBar = BarDefault

End Sub


