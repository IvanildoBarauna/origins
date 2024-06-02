Attribute VB_Name = "mdCONT"
Option Explicit
Dim Botão As Boolean

Sub Contador()

If Botão Then
CONT.Calculate
Application.OnTime Now() + TimeValue("00:00:01"), "Contador"

End If

End Sub

Sub Ligar()

Botão = True
Call Contador

End Sub

Sub Desligar()

Botão = False

End Sub

Sub Modo_App()

ActiveWindow.DisplayWorkbookTabs = 0
ActiveWindow.DisplayHeadings = 0
ActiveWindow.DisplayHorizontalScrollBar = 0
ActiveWindow.DisplayVerticalScrollBar = 0

Application.DisplayFullScreen = 1
Application.DisplayStatusBar = 1
Application.DisplayFormulaBar = 0

End Sub

Sub Modo_App_Desliga()

ActiveWindow.DisplayWorkbookTabs = 1
ActiveWindow.DisplayHeadings = 1
ActiveWindow.DisplayHorizontalScrollBar = 1
ActiveWindow.DisplayVerticalScrollBar = 1

Application.DisplayFormulaBar = 1
Application.DisplayFullScreen = 0
Application.DisplayStatusBar = 1

End Sub




