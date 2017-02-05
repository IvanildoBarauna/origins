Attribute VB_Name = "CONFIGURAÇÕES"
Public Sub ACTIVATE_()

With Application

    .DisplayAlerts = False
    .ScreenUpdating = False

End With

End Sub

Public Sub DEACTIVATE_()

With Application

    .DisplayAlerts = True
    .ScreenUpdating = True
    .StatusBar = False

End With


End Sub


Public Sub ACTIVATE_APP()

With ActiveWindow

    .DisplayWorkbookTabs = False
    .DisplayHorizontalScrollBar = False
    .DisplayVerticalScrollBar = False
    
End With

With Application

    .DisplayFullScreen = True
    .DisplayFormulaBar = False
    .DisplayStatusBar = False
    
End With

End Sub

Public Sub DEACTIVATE_APP()

With ActiveWindow

    .DisplayWorkbookTabs = True
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
    
End With
    
With Application

    .DisplayFullScreen = False
    .DisplayFormulaBar = True
    .DisplayStatusBar = True

    
End With

End Sub


