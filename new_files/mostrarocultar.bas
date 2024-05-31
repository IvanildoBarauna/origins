Attribute VB_Name = "mostrarocultar"
'PROCEDIMENTO PARA OCULTAR_TUDO E EXIBIR_TUDO
Sub ocultar_tudo()
'Menu superior
Application.ExecuteExcel4Macro "show.toolbar(""ribbon"",false)"
'barra de Fórmulas
Application.DisplayFormulaBar = False
'barra de status
'Application.DisplayStatusBar = False


'Cabeçalhos
ActiveWindow.DisplayHeadings = False
'Guias da planilha
ActiveWindow.DisplayWorkbookTabs = False
'Linhas de grade
'ActiveWindow.DisplayGridlines = False
'Barras horozontais
'ActiveWindow.DisplayHorizontalScrollBar = False
'barras verticais
'ActiveWindow.DisplayVerticalScrollBar = False
Range("C7").Select

End Sub


Sub mostrar_tudo()
'Application.DisplayFullScreen = False
'Menu superior
Application.ExecuteExcel4Macro "show.toolbar(""ribbon"",true)"
'barra de Fórmulas
Application.DisplayFormulaBar = True
'barra de status
'Application.DisplayStatusBar = True


'Cabeçalhos
ActiveWindow.DisplayHeadings = True
'Guias da planilha
ActiveWindow.DisplayWorkbookTabs = True
'Linhas de grade
'ActiveWindow.DisplayGridlines = True
'Barras horozontais
ActiveWindow.DisplayHorizontalScrollBar = True
'barras verticais
ActiveWindow.DisplayVerticalScrollBar = True
End Sub

