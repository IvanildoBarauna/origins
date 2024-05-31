Attribute VB_Name = "mdFormsAux"
Option Explicit

Public Sub FiltrarDados(AgenteName As String)
#If booAux Then
    Dim oDic As Scripting.Dictionary: Set oDic = New Scripting.Dictionary
#Else
    Dim oDic As Object: Set oDic = VBA.CreateObject("Scripting.Dictionary")
#End If
    Dim lo          As Excel.ListObject
    Dim counter     As Integer, lCounter    As Integer
    Dim fItem       As String
    Dim sKey        As Integer
    Dim loDB        As Excel.ListObject
    Dim Address     As String
    Dim iCol        As Integer
    
    Set loDB = shBD.ListObjects(1)
    Set lo = wsRuasAgents.ListObjects(1)
    
    For counter = 1 To lo.ListRows.Count
        If lo.DataBodyRange(counter, 2).Value2 = AgenteName Then
            For lCounter = 1 To loDB.ListRows.Count
                Address = loDB.DataBodyRange(lCounter, 6).Value2
                If Address Like "*" & lo.DataBodyRange(counter, 4).Value2 & "*" Then
                    sKey = sKey + 1
                    oDic.Add sKey, Address
                End If
            Next lCounter
        End If
    Next counter
    
    loDB.Range.AutoFilter 6
    lCounter = 0
    
    If oDic.Count > 0 Then loDB.Range.AutoFilter 6, Array(oDic.Items()), xlFilterValues
End Sub

Function GetFunctional(AgentName As String) As String
    Dim iCounter As Long
    Dim lo      As Excel.ListObject
    
    Set lo = wsListaAgents.ListObjects(1)
    
    For iCounter = 1 To lo.ListRows.Count
        If AgentName = "DESCOBERTA" Then
            GetFunctional = "DESCOBERTA"
            Exit For
        ElseIf lo.DataBodyRange(iCounter, lo.ListColumns("NOME").index).Value2 = AgentName Then
            GetFunctional = lo.DataBodyRange(iCounter, lo.ListColumns("FUNCIONAL").index).Value2
            Exit For
        End If
    Next iCounter
    
End Function

Public Sub ClearFields(frm As MSForms.UserForm, _
                       ParamArray ControlsType())
    Dim xCtrl   As MSForms.Control, iType
    
    For Each xCtrl In frm.Controls
        If xCtrl.Name = "txtData" Then
                xCtrl.Value = VBA.Date
            Else
                For Each iType In ControlsType
                    If iType = TypeName(xCtrl) Then
                        Select Case VBA.TypeName(xCtrl)
                            Case Is = "TextBox", "ComboBox"
                                If xCtrl.Value <> "" Then xCtrl.Value = ""
                            Case Is = "CheckBox", "OptionButton"
                                If xCtrl.Value Then xCtrl.Value = False
                            Case Is = "Label"
                                If xCtrl.Name = "lbctrl" Then xCtrl.Caption = ""
                        End Select
                        GoTo NextCtrl
                    End If
                Next iType
            End If
NextCtrl:
    Next xCtrl
End Sub

Public Function ValidateEmptyFields(ByVal Source As Object) As Boolean
    Dim Field As MSForms.Control

    For Each Field In Source.Controls
        Select Case VBA.TypeName(Field)
            Case "TextBox", "ComboBox"
                If Field.Value = vbNullString And Field.Enabled Then
                    Field.SetFocus
                    ValidateEmptyFields = True
                    Exit Function
                End If
        End Select
    Next Field
    
End Function


Function isSelected(ListBoxControl As MSForms.ListBox, HeaderValidate As XlYesNoGuess) As Boolean
    Dim iCounter  As Long
    Dim UpperFor  As Integer
    
    UpperFor = VBA.IIf(HeaderValidate = xlYes, ListBoxControl.ListCount - 1, ListBoxControl.ListCount)
    
    For iCounter = 1 To UpperFor
        If ListBoxControl.Selected(iCounter) Then
            isSelected = True
            Exit For
        End If
    Next iCounter
    
End Function

Function SelectedCounter(ListBoxControl As MSForms.ListBox, HeaderValidate As XlYesNoGuess) As Long
    Dim iCounter    As Long
    Dim UpperFor    As Integer
    Dim AuxCounter  As Long
    
    UpperFor = VBA.IIf(HeaderValidate = xlYes, ListBoxControl.ListCount - 1, ListBoxControl.ListCount)
    
    For iCounter = 1 To UpperFor
        If ListBoxControl.Selected(iCounter) Then
            AuxCounter = AuxCounter + 1
        End If
    Next iCounter
    
    SelectedCounter = AuxCounter
End Function

Public Function FilterArray(ByVal mtz, _
                            ByVal iCol As Integer, _
                            ByVal Criteria As String)
                                    
'------------------------------------------------------
'RotineType: Function / Variant - Array
'Criacao: Ivanildo Junior
'Criada em: 10/03/2018 / 19:41
'Objetivo: Filtrar uma coluna de data de uma matriz com os dados da data atual
'Aplicacaoo: FilterArrayWithDate(YourArray, 4)
'------------------------------------------------------
                          
    Dim mtzResult   As Variant
    Dim index       As Long
    Dim RowCounter  As Long
    Dim ColCounter  As Integer
    Dim mtzSize     As Long
    Dim mtzValue    As String
    
    For index = LBound(mtz, 1) To UBound(mtz, 1)
        On Error Resume Next
        mtzValue = mtz(index, iCol)
        On Error GoTo 0
        If mtzValue = Criteria Then mtzSize = mtzSize + 1
    Next index
    
    mtzValue = ""
    
    ReDim mtzResult(1 To mtzSize + 1, 1 To UBound(mtz, 2))
    
    mtzResult(1, 1) = "Funcional"
    mtzResult(1, 2) = "Nome Agente"
    mtzResult(1, 3) = "ÁREA.MICROÁREA"
    mtzResult(1, 4) = "Nome da Rua"
    mtzResult(1, 5) = "Bairro"
    mtzResult(1, 6) = "CEP"
    mtzResult(1, 7) = "Observação"
    
    For index = LBound(mtz, 1) To UBound(mtz, 1)
        On Error Resume Next
        mtzValue = mtz(index, iCol)
        On Error GoTo 0
        If mtzValue = Criteria Then
            RowCounter = RowCounter + 1
            For ColCounter = 1 To UBound(mtzResult, 2)
                mtzResult(RowCounter + 1, ColCounter) = mtz(index, ColCounter)
            Next ColCounter
        End If
    Next index
    
    FilterArray = mtzResult
End Function

