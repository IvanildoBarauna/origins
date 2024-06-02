Attribute VB_Name = "MostarOcultar"
'PROCEDIMENTO PARA OCULTAR_TUDO E EXIBIR_TUDO
Sub ocultar_tudo()
Application.DisplayFullScreen = True
ActiveWindow.DisplayHeadings = False
Range("C7").Select

End Sub


Sub mostrar_tudo()
Application.DisplayFullScreen = False
'Cabeçalhos
ActiveWindow.DisplayHeadings = True
End Sub

