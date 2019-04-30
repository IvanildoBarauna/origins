Attribute VB_Name = "Módulo1"

Option Private Module
Public Enum Orientation
    Vertical = 1
    Horizontal = 2
End Enum

Public Enum FilterArrayAction
    Keep = 0
    Remove = 1
End Enum
'---------------------------------------------------------------------------------------
' Rotina....: FilterArray() / Function
' Contato...: fernando.fernandes@outlook.com.br
' Autor.....: Fernando Fernandes
' Empresa...: Planilhando
' Descricaoo.: This routine filters the content of any bidimensional array, with the given criterias
'---------------------------------------------------------------------------------------
Public Function FilterArray(ByVal arr As Variant, _
                            ByVal FilterAction As FilterArrayAction, _
                            ByVal Header As XlYesNoGuess, _
                            ByVal lColumn As Long, _
                            ByVal Criterias As String, _
                            Optional ByRef Registros As Long = 0) As Variant
On Error GoTo TreatError
Dim lCounterArray       As Long
Dim lCounterCriteria    As Long
'Dim lCounterAuxiliar    As Long
Dim lCounterItems       As Long
Dim lColumns            As Long
Dim arrCriterias        As Variant
Dim arrAux              As Variant

    If VBA.IsArray(arr) Then
        If lColumn <= UBound(arr, 2) Then
            lCounterItems = 0
            arrCriterias = VBA.Split(VBA.CStr(Criterias), ",", -1)
            
    'creating auxiliary array with same dimensions
            ReDim arrAux(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
            
            For lCounterArray = LBound(arr, 1) To UBound(arr, 1) Step 1
                If lCounterArray = 1 And Header = xlYes Then
                    lCounterItems = lCounterItems + 1
                    For lColumns = LBound(arrAux, 2) To UBound(arrAux, 2) Step 1
                        arrAux(lCounterItems, lColumns) = arr(lCounterArray, lColumns)
                    Next
                Else
                    For lCounterCriteria = LBound(arrCriterias, 1) To UBound(arrCriterias, 1) Step 1

                        If (FilterAction = Keep And VBA.UCase(arr(lCounterArray, lColumn)) Like VBA.UCase("*" & arrCriterias(lCounterCriteria)) & "*") Or _
                           (FilterAction = Remove And Not VBA.UCase(arr(lCounterArray, lColumn)) Like VBA.UCase("*" & arrCriterias(lCounterCriteria)) & "*") Then
                           
                            lCounterItems = lCounterItems + 1
                            For lColumns = LBound(arrAux, 2) To UBound(arrAux, 2) Step 1
                                arrAux(lCounterItems, lColumns) = arr(lCounterArray, lColumns)
                            Next
                        End If
                    
                    Next lCounterCriteria
                End If
                VBA.DoEvents
            Next lCounterArray
            
            Registros = lCounterItems
            If lCounterItems = 0 Then
                arrAux = Empty
            Else
                arrAux = TransposeArray(arrAux)
                ReDim Preserve arrAux(LBound(arrAux, 1) To UBound(arrAux, 1), LBound(arrAux, 2) To lCounterItems)
                arrAux = TransposeArray(arrAux)
            End If
        End If
    End If
    
    FilterArray = arrAux
    
On Error GoTo 0
Exit Function
TreatError:
'    Call xlExceptions.TreatError(VBA.Err.Description, VBA.Err.Number, "xlArrays.FilterArray()", Erl, True)
End Function

'---------------------------------------------------------------------------------------
' Rotina....: TransposeArray() / Function
' Contato...: fernando.fernandes@outlook.com.br
' Autor.....: Jefferson Dantas
' Revis?o...: Fernando Fernandes
' Empresa...: Planilhando
' Descri??o.: This routine transposes any bidimensional array
'---------------------------------------------------------------------------------------
Public Function TransposeArray(ByVal Matriz As Variant) As Variant
On Error GoTo TreatError
Dim lngContador     As Long
Dim lngContador1    As Long
Dim arrAux          As Variant

    If VBA.IsArray(Matriz) Then
        Select Case NumberOfDimensions(Matriz)
            Case 1
                Matriz = Matriz
            Case 2
'creating auxiliary array with inverted dimensions
                ReDim arrAux(LBound(Matriz, 2) To UBound(Matriz, 2), LBound(Matriz, 1) To UBound(Matriz, 1))
                
                For lngContador = LBound(Matriz, 2) To UBound(Matriz, 2) Step 1
                    For lngContador1 = LBound(Matriz, 1) To UBound(Matriz, 1) Step 1
                        arrAux(lngContador, lngContador1) = Matriz(lngContador1, lngContador)
                    Next lngContador1
                Next lngContador
        End Select
    End If
    TransposeArray = arrAux
On Error GoTo 0
Exit Function
TreatError:
'    Call xlExceptions.TreatError(VBA.Err.Description, VBA.Err.Number, "xlArrays.TransposeArray()", Erl, True)
End Function

Public Function NumberOfDimensions(ByVal arr As Variant) As Long
On Error GoTo TreatError
Dim cnt As Long

    cnt = 1
    Do Until Err.Number <> 0
        If LBound(arr, cnt) >= 0 Then NumberOfDimensions = cnt
        cnt = cnt + 1
    Loop
    
On Error GoTo 0
Exit Function
TreatError:
    Exit Function
    'Call xlExceptions.TreatError(VBA.Err.Description, VBA.Err.Number, "xlArrays.NumberOfDimensions()", Erl, True)
End Function

'---------------------------------------------------------------------------------------
' Rotina....: AppendArrays() / Function
' Contato...: Fernando.Fernandes@Outlook.com.br
' Autor.....: Fernando Fernandes
' Ad........: www.Planilhando.Com.Br
' Date......: 02/18/2014
' Descr.....: Routine that will append array 1 under array one. Both have to be the same dimensions
'---------------------------------------------------------------------------------------
Public Function AppendArrays(ByRef ArrayUp As Variant, _
                             ByRef ArrayDown As Variant, _
                             Optional ByVal Direction As Orientation = Orientation.Vertical)
On Error GoTo TreatError
Dim lRowStart   As Long
Dim lRowEnd     As Long
Dim lRow        As Long
Dim lColStart   As Long
Dim lColEnd     As Long
Dim lCol        As Long
Dim arrAux      As Variant
Dim ArrayLeft   As Variant
Dim ArrayRight  As Variant

    If VBA.IsArray(ArrayUp) And VBA.IsArray(ArrayDown) Then

        Select Case Direction
            Case Orientation.Vertical
                If UBound(ArrayUp, 2) = UBound(ArrayDown, 2) Then
                    lRowStart = UBound(ArrayUp, 1) + 1
                    lRowEnd = UBound(ArrayUp, 1) + UBound(ArrayDown, 1)
                    arrAux = ArrayUp
                    
                    Call ResizeArray(arrAux, lRowEnd)

                    For lRow = lRowStart To lRowEnd
                        For lCol = LBound(ArrayUp, 2) To UBound(ArrayUp, 2)
                            arrAux(lRow, lCol) = ArrayDown(lRow - UBound(ArrayUp, 1), lCol)
                        Next lCol
                    Next lRow
                End If
            
            Case Orientation.Horizontal
                ArrayLeft = ArrayUp
                ArrayRight = ArrayDown

                If UBound(ArrayLeft, 1) = UBound(ArrayRight, 1) Then

                    lColStart = UBound(ArrayLeft, 2) + 1
                    lColEnd = UBound(ArrayLeft, 2) + UBound(ArrayRight, 2)
                    arrAux = ArrayLeft

                    ReDim Preserve arrAux(LBound(ArrayUp, 1) To UBound(ArrayUp, 1), 1 To lColEnd)

                    For lCol = lColStart To lColEnd
                        For lRow = LBound(ArrayLeft, 1) To UBound(ArrayLeft, 1)
                            arrAux(lRow, lCol) = ArrayRight(lRow, lCol - UBound(ArrayLeft, 2))
                        Next lRow
                    Next lCol

                End If

        End Select
        If VBA.IsArray(arrAux) Then AppendArrays = arrAux
    End If

On Error GoTo 0
Exit Function
TreatError:
    '''Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "TrimArray()", Erl, True)
End Function

'---------------------------------------------------------------------------------------
' Rotina....: ResizeArray() / Sub
' Contato...: fernando.fernandes@outlook.com.br
' Autor.....: Jefferson Dantas
' Revis?o...: Fernando Fernandes
' Empresa...: Planilhando
' Descri??o.: This routine resizes any bidimensional array, to a new number of rows, keeping the contents
'             Redim Preserve
'---------------------------------------------------------------------------------------
Public Sub ResizeArray(ByRef mtz As Variant, ByVal NewSize As Long)
On Error GoTo TreatError

Dim FirstElementRow As Long, LastElementRow As Long
Dim FirstElementCol As Long, LastElementCol As Long

    FirstElementRow = LBound(mtz, 1): FirstElementCol = LBound(mtz, 2)

    LastElementRow = UBound(mtz, 1):  LastElementCol = UBound(mtz, 2)

    mtz = TransposeArray(mtz)
    ReDim Preserve mtz(FirstElementCol To LastElementCol, FirstElementRow To NewSize)
    mtz = TransposeArray(mtz)

On Error GoTo 0
Exit Sub
TreatError:
    '''Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ResizeArray()", Erl, True)
End Sub

'---------------------------------------------------------------------------------------
' Modulo    : MD_FF_Array / M?dulo
' Rotina    : RemoverDuplicados() / Function
' Autor     : Fernando Fernandes / Jeff Dantas
' Data      : 07/08/2015
' Proposta  : Fun??o para remover duplicados
'---------------------------------------------------------------------------------------
Public Function RemoverDuplicados(ByRef mtzCompleta As Variant) As Variant
On Error GoTo TRATARERRO
Dim dicChave        As New Scripting.Dictionary
Dim linCompleta     As Long
Dim col             As Long
Dim Chave           As String
    
    If VBA.IsArray(mtzCompleta) Then
        linCompleta = LBound(mtzCompleta, 1)
        Do While linCompleta <= UBound(mtzCompleta, 1)
            Chave = vbNullString
            For col = LBound(mtzCompleta, 2) To UBound(mtzCompleta, 2)
                Chave = Chave & mtzCompleta(linCompleta, col) & "|"
            Next col
            
            If dicChave.Exists(Chave) Then
                mtzCompleta = RemoveRowFromArray(mtzCompleta, linCompleta)
            Else
               dicChave.Add Chave, Chave
            End If
            linCompleta = linCompleta + 1
        Loop
        RemoverDuplicados = mtzCompleta
   End If
    
Set dicChave = Nothing
    
On Error GoTo 0
Exit Function
TRATARERRO:
    '''Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "CriarMatrizBidimensional()", Erl)
    Exit Function
End Function

'---------------------------------------------------------------------------------------
' Rotina....: RemoveRowFromArray() / Function
' Contato...: Fernando.Fernandes@Outlook.com.br
' Autor.....: Fernando Fernandes
' Ad........: www.Planilhando.Com.Br
' Date......: 08/14/2014
' Descr.....:
'---------------------------------------------------------------------------------------
Public Function RemoveRowFromArray(ByRef arr As Variant, _
                                   ByVal Row As Long) As Variant
On Error GoTo TreatError
Dim CutOffArray As Variant
Dim cntSource   As Long
Dim cntDestiny  As Long
Dim lCol        As Long

    If VBA.IsArray(arr) Then

        ReDim CutOffArray(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
        cntDestiny = LBound(arr, 1)

        For cntSource = LBound(arr, 1) To UBound(arr, 1)
            If cntSource <> Row Then
                For lCol = LBound(arr, 2) To UBound(arr, 2)

                    CutOffArray(cntDestiny, lCol) = arr(cntSource, lCol)

                Next lCol
                cntDestiny = cntDestiny + 1
            End If
        Next cntSource
        Call ResizeArray(CutOffArray, cntDestiny - 1)
        RemoveRowFromArray = CutOffArray

    Else
        RemoveRowFromArray = arr
    End If

    Call LIMPAR_MEMORIA(CutOffArray, cntSource, cntDestiny, lCol, arr)

On Error GoTo 0
Exit Function
TreatError:
    '''Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "RemoveRowFromArray()", Erl, True)

End Function

'---------------------------------------------------------------------------------------


