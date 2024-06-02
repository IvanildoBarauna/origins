Attribute VB_Name = "mdSolutions"
Option Explicit

Sub ProjectEulerProblem1()
'https://projecteuler.net/problem=1
'Encontre a soma de todos os múltiplos de 3 ou 5 abaixo de 1000

    Dim iCounter As Integer
    Dim vSum As Long
    
    For iCounter = 1 To 999
        If iCounter Mod 3 = 0 Or iCounter Mod 5 = 0 Then
            vSum = vSum + iCounter
        End If
    Next iCounter
    
    Debug.Print vSum
End Sub

Sub ProjectEulerProblem2()
'https://projecteuler.net/problem=2
'Na série de números de Fibonacci cujos valores não excedem quatro milhões, _
    encontre a soma dos termos de valor par.

    Dim FiboValues() As Long
    Dim iFibo        As Double
    Dim vSum         As Double
    
    ReDim FiboValues(1 To 2)
    
    FiboValues(1) = 1
    FiboValues(2) = 2
    iFibo = 2
    vSum = 2
    
    Do
        iFibo = iFibo + 1
        ReDim Preserve FiboValues(1 To iFibo)
        FiboValues(iFibo) = FiboValues(iFibo - 1) + FiboValues(iFibo - 2)
        
        If FiboValues(iFibo) > 4000000# Then Exit Do
        
        If FiboValues(iFibo) Mod 2 = 0 Then
            vSum = vSum + FiboValues(iFibo)
        End If
    Loop While FiboValues(iFibo) < 4000000#
    
    Debug.Print CStr(vSum)
End Sub

Sub ProjectEulerProblem17()
'https://projecteuler.net/problem=17
'Se todos os números de 1 a 1000 (mil) inclusive fossem escritos em palavras, quantas letras seriam usadas?
'NOTA: Não conte espaços ou hifens. _
       Por exemplo, 342 (trezentos e quarenta e dois) contém 23 letras e 115 (cento e quinze) contém 20 letras. _
       O uso de "e" ao redigir números está em conformidade com o uso britânico.
       
    Dim iCounter    As Double
    Dim vSum        As Double
    Dim vExtenso    As String
    
    For iCounter = 1 To 1000
        vExtenso = mdExtenso.Extenso_Valor(iCounter)
        vExtenso = VBA.Replace(VBA.IIf(iCounter < 2, VBA.Replace(vExtenso, "real", ""), _
                                         VBA.Replace(vExtenso, "reais", "")), " ", "")
        vSum = vSum + VBA.Len(vExtenso)
    Next iCounter
    
    Debug.Print vSum
End Sub
