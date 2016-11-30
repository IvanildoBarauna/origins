Attribute VB_Name = "mdConfig"
Public Sub ModoApp(OnOff As Boolean)
     With Application
          .DisplayFullScreen = OnOff
          .DisplayFormulaBar = Not OnOff
          .DisplayScrollBars = Not OnOff
          .DisplayStatusBar = Not OnOff
     End With

     With ActiveWindow
          .WindowState = xlMaximized
          .DisplayHeadings = Not OnOff
          .DisplayWorkbookTabs = Not OnOff
     End With
End Sub
