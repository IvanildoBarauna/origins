Attribute VB_Name = "Módulo1"
Option Explicit
Sub GlebasApp()
    Dim bot         As Selenium.WebDriver
    Dim chk         As By
    Dim IglebaInput As Variant
    Dim sText       As String
    Dim iLinAtual   As Integer
    Dim iLinUltima  As Integer
    Dim allInputs   As Object
    Dim element     As Object
    Dim xpath       As String
    Dim destino     As Excel.Range
    Dim ws          As Excel.Worksheet
    Dim iniTime     As Single
    
    iniTime = VBA.Timer
    Set ws = Planilha1
    Set destino = ws.Range("B16")
    Set chk = New By
    Set bot = New Selenium.ChromeDriver
    
    With bot
        .AddArgument "--disable-plugins-discovery"
        .AddArgument "--disable-extensions"
        .AddArgument "--disable-infobars"
        .SetPreference "plugins.plugins_disabled", Array("Adobe Flash Player")
        .Get ("http://www.glebas.com.br/")
        .Window.Maximize
        .FindElementByXPath("/html/body/header/div/div/div[1]/button").Click
        .ExecuteScript ("showGlebaModal();")
        
        With ws
            iLinUltima = .Cells(.Rows.Count, 2).End(xlUp).Row
            For iLinAtual = 10 To iLinUltima + 8
                sText = sText & "1 " & .Cells(iLinAtual, 2) & " " & _
                    .Cells(iLinAtual, 3) & " " & .Cells(iLinAtual, 4) & VBA.vbCrLf
            Next iLinAtual
        End With
        
        .Wait 5000
        .FindElementById("glebaInput").SendKeys sText
        .ExecuteScript ("showGlebaMap();")
        .Wait 5000
        .ExecuteScript ("showArea();")
        On Error Resume Next
        .FindElementByXPath("//*[@id=""map""]/div/div/div[9]/div[2]").Click
        .WindowHandles.Last
        .SwitchTo.Window
        MsgBox .FindElementByClass("table").Text
    End With
    Debug.Print "Tempo Total: " & VBA.Format(VBA.Timer - iniTime, "0.00 segundos")
End Sub
