Attribute VB_Name = "CORES"
'RETORNO O ÍNDICE DA COR, BASEADO EM UM INTERVALO INFORMADO'

'ESSA DICA EU PEGUEI NO CANAL INTER EXCEL, CONFIRAM'

Function INDEXCOLOR(CellColor As Range)

INDEXCOLOR = CellColor.Interior.ColorIndex

End Function

'RETORNA O NOME DA COR BASEADO EM UM INTERVALO INFORMADO, SINTAXE =CELLCOLOR(INTERVALO;VERDADEIRO)'

'ESSE CÓDIGO EU PEGUEI DE UM SITE AONDE EU NÃO LEMBRO O LINK PESSOAL, PORTANTO NÃO TEM AUTORIA MINHA!!'

Function CellColor(rCell As Range, Optional ColorName As Boolean)
    Dim strColor As String, iIndexNum As Integer
 
    Select Case rCell.Interior.ColorIndex
    Case 1
        strColor = "Black"
        iIndexNum = 1
    Case 53
        strColor = "Brown"
        iIndexNum = 53
    Case 52
        strColor = "Olive Green"
        iIndexNum = 52
    Case 51
        strColor = "Dark Green"
        iIndexNum = 51
    Case 49
        strColor = "Dark Teal"
        iIndexNum = 49
    Case 11
        strColor = "Dark Blue"
        iIndexNum = 11
    Case 55
        strColor = "Indigo"
        iIndexNum = 55
    Case 56
        strColor = "Gray-80%"
        iIndexNum = 56
    Case 9
        strColor = "Dark Red"
        iIndexNum = 9
    Case 46
        strColor = "Orange"
        iIndexNum = 46
    Case 12
        strColor = "Dark Yellow"
        iIndexNum = 12
    Case 10
        strColor = "Green"
        iIndexNum = 10
    Case 14
        strColor = "Teal"
        iIndexNum = 14
    Case 5
        strColor = "Blue"
        iIndexNum = 5
    Case 47
        strColor = "Blue-Gray"
        iIndexNum = 47
    Case 16
        strColor = "Gray-50%"
        iIndexNum = 16
    Case 3
        strColor = "Red"
        iIndexNum = 3
    Case 45
        strColor = "Light Orange"
        iIndexNum = 45
    Case 43
        strColor = "Lime"
        iIndexNum = 43
    Case 50
        strColor = "Sea Green"
        iIndexNum = 50
    Case 42
        strColor = "Aqua"
        iIndexNum = 42
    Case 41
        strColor = "Light Blue"
        iIndexNum = 41
    Case 13
        strColor = "Violet"
        iIndexNum = 13
    Case 48
        strColor = "Gray-40%"
        iIndexNum = 48
    Case 7
        strColor = "Pink"
        iIndexNum = 7
    Case 44
        strColor = "Gold"
        iIndexNum = 44
    Case 6
        strColor = "Yellow"
        iIndexNum = 6
    Case 4
        strColor = "Bright Green"
        iIndexNum = 4
    Case 8
        strColor = "Turqoise"
        iIndexNum = 8
    Case 33
        strColor = "Sky Blue"
        iIndexNum = 33
    Case 54
        strColor = "Plum"
        iIndexNum = 54
    Case 15
        strColor = "Gray-25%"
        iIndexNum = 15
    Case 38
        strColor = "Rose"
        iIndexNum = 38
    Case 40
        strColor = "Tan"
        iIndexNum = 40
    Case 36
        strColor = "Light Yellow"
        iIndexNum = 36
    Case 35
        strColor = "Light Green"
        iIndexNum = 35
    Case 34
        strColor = "Light Turqoise"
        iIndexNum = 34
    Case 37
        strColor = "Pale Blue"
        iIndexNum = 37
    Case 39
        strColor = "Lavendar"
        iIndexNum = 39
    Case 2
        strColor = "White"
        iIndexNum = 2
    Case Else
        strColor = "Custom color or no fill"
    End Select
 
    If ColorName = True Or _
       strColor = "Custom color or no fill" Then
        CellColor = strColor
    Else
        CellColor = iIndexNum
    End If
 
End Function




