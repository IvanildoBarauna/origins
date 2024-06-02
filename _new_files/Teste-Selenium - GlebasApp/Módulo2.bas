Attribute VB_Name = "Módulo2"
Option Explicit

Dim driver As WebDriver
Dim bot As WebDriver
Dim chk As By
Dim IglebaInput         As Variant
 Sub GlebasApp()
 
 'Variáveis
 
Set driver = New ChromeDriver
Set chk = New By
'Dim destino As Range
'Dim tabela As WebElement
Dim sText As String
Dim iLinAtual As Integer
Dim iLinUltima As Integer
Dim allInputs   As Object
Dim element     As Object
Dim xpath As String
Dim destino As Range
Set destino = Range("B16")

Dim iniTime As Single

iniTime = VBA.Timer



'Abre chrome e a página específica
Set bot = New ChromeDriver

With bot
    .AddArgument "--disable-plugins-discovery"
    .AddArgument "--disable-extensions"
    .AddArgument "--disable-infobars"
    .SetPreference "plugins.plugins_disabled", Array("Adobe Flash Player")
    .Get ("http://www.glebas.com.br/")
    .Window.Maximize
  
End With


 'Abre o dropdown
     
         bot.FindElementByXPath("/html/body/header/div/div/div[1]/button").Click
                  
   'Clica em abrir Glebas
   
bot.ExecuteScript ("showGlebaModal();")

      'Insegir gleba na caixa de texto:

  
  iLinUltima = Application.WorksheetFunction.CountA(ThisWorkbook.Worksheets(1).Range("B:B"))
  With Worksheets(1)
    For iLinAtual = 10 To iLinUltima + 8
      sText = sText & "1 " & .Cells(iLinAtual, 2) & " " & .Cells(iLinAtual, 3) & " " & .Cells(iLinAtual, 4) & VBA.vbCrLf
    Next iLinAtual
  End With
  
  bot.Wait 5000
  bot.FindElementById("glebaInput").SendKeys sText
  
  bot.ExecuteScript ("showGlebaMap();")
      
'bot.FindElementByXPath("/html/body/header/div/div/div[1]/button").Click

 'On Error Resume Next

' executa ação para abrir uma nova janela
 bot.Wait 5000
 bot.ExecuteScript ("showArea();")
 On Error Resume Next

'mude para a nova janela.
On Error Resume Next

bot.WindowHandles.Last
bot.SwitchTo.Window
MsgBox bot.FindElementByClass("table").Text
        
' se você quiser voltar para sua primeira janela
'driver.SwitchToWindow (driver.WindowHandles.First)

 'Alterar o tipo de mapa para Satélite:
 
bot.FindElementByXPath("//*[@id=""map""]/div/div/div[9]/div[2]").Click



'Aumentar zoom


 'bot.FindElementByXPath("//*[@id=""map""]/div/div/div[8]/div[1]/div/button[1]").Click
 
'bot.FindElementByXPath("//*[@id=""map""]/div/div/div[1]/div[4]/div[4]").Click



'MsgBox bot.FindElementByClass("table").Text

 'importar
   
'Worksheets(2).Range("B2").Value = bot.FindElementByClass("table").Text

 
'bot.SwitchToFrame (0)

'Worksheets(2).Range("B2").Value = bot.FindElementByClassName("table").Text
 
   
 'Worksheets(2).Range("B2").Value = bot.FindElementByClass("gm-style-iw").Text
 
'MsgBox bot.FindElementByClass("table").Text

  'MsgBox bot.FindElementByClass("gm-style-iw").Text
  Debug.Print VBA.Format(VBA.Timer - iniTime, "0.00 segundos")

 End Sub


