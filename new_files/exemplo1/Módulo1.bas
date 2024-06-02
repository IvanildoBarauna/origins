Attribute VB_Name = "Módulo1"
Option Explicit
'====================================================================
'Nome.........: UnirTexto *(Criada e incluída no Excel 2016, mas ausente nas outras versões)
'Data.........: 29/03/2016 (dmy)
'Autor........: Fernando Fernandes
'Contato......: Fernando.Fernandes@outlook.com.br
'Descrição....: Concatena todos os textos informados no último parâmetro,
'               usando o delimitador para separá-los, ignorando ou não as células vazias
'Forum........: www.Planilhando.com.br
'====================================================================
Public Function UnirTexto(ByVal Delimitador As String, _
                          ByVal IgnorarVazios As Boolean, _
                          ParamArray Celulas() As Variant) As Variant
On Error GoTo TratarErro
Dim Intervalo   As Variant
Dim Resultado   As String
Dim i           As Long
Dim j           As Long
Dim k           As Long
   
    If UBound(Celulas, 1) < LBound(Celulas, 1) Then
        UnirTexto = VBA.Conversion.CVErr(xlErrValue)
        Exit Function
    End If
   
    For i = LBound(Celulas, 1) To UBound(Celulas, 1) Step 1

        If VBA.Information.IsArray(Celulas(i)) Then
       
            Intervalo = Celulas(i)
            For j = LBound(Intervalo, 1) To UBound(Intervalo, 1) Step 1
                For k = LBound(Intervalo, 2) To UBound(Intervalo, 2) Step 1
                    If Not VBA.Information.IsError(Intervalo(j, k)) Then
                        If Not VBA.Strings.Trim(Intervalo(j, k)) = vbNullString Then
                            Resultado = Resultado & Delimitador & Intervalo(j, k)
                        End If
                    End If
                Next k
               
            Next j
           
        Else
       
            If Not VBA.Information.IsError(Celulas(i)) Then
                If Not VBA.Strings.Trim(Celulas(i)) = vbNullString Then
                    Resultado = Resultado & Delimitador & Celulas(i)
                End If
            End If
           
        End If
       
    Next i

    If VBA.Strings.Len(Resultado) > VBA.Strings.Len(Delimitador) Then
        Resultado = VBA.Strings.Right(Resultado, VBA.Strings.Len(Resultado) - VBA.Strings.Len(Delimitador))
    End If
    UnirTexto = Resultado
    Exit Function
TratarErro:
    UnirTexto = VBA.Conversion.CVErr(xlErrValue)
End Function





