Attribute VB_Name = "frmMAin"
Attribute VB_Base = "0{E93F9ADC-4B13-45B6-8D0C-922AB30C8418}{65DDFAEC-18BF-4DEA-B87F-7857B9A80294}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub RotineFind(ByVal xCombo As MSForms.ComboBox)
    Dim ws       As Worksheet
    Dim lo       As ListObject
    Dim FoundRow As Long
    Dim MinValue As Long
    Dim MaxValue As Long
    
    Set ws = Planilha1
    Set lo = ws.ListObjects("tbProdutos")
    
    MinValue = Application.WorksheetFunction.Min(lo.ListColumns("ID").DataBodyRange.Value2)
    MaxValue = Application.WorksheetFunction.Max(lo.ListColumns("ID").DataBodyRange.Value2)
    
    If Me.txtName.Value <> vbNullString And VBA.IsNumeric(Me.txtName.Value) Then
        If Me.txtName.Value >= MinValue And Me.txtName.Value <= MaxValue Then
            FoundRow = lo.ListColumns("ID").Range.Find(Me.txtName.Value, lo.Range(1, 1), xlValues, xlWhole, _
                            xlByRows, xlNext, False, False).Row - 1
            xCombo.Value = lo.DataBodyRange(FoundRow, 2).Value2
        Else
            MsgBox "Digite um valor entre: " & MinValue & " e " & MaxValue, vbExclamation
        End If
    Else
        MsgBox "Digite um valor válido (ID númerico)", vbExclamation
    End If
End Sub

Public Sub RotineFunctionMatch(ByVal xCombo As MSForms.ComboBox)
    Dim ws       As Worksheet
    Dim lo       As ListObject
    Dim FoundRow As Long
    Dim MinValue As Long
    Dim MaxValue As Long
    
    Set ws = Planilha1
    Set lo = ws.ListObjects("tbProdutos")
    MinValue = Application.WorksheetFunction.Min(lo.ListColumns("ID").DataBodyRange.Value2)
    MaxValue = Application.WorksheetFunction.Max(lo.ListColumns("ID").DataBodyRange.Value2)
    
    With Me.txtName
        If .Value <> vbNullString And VBA.IsNumeric(.Value) Then
            If .Value >= MinValue And .Value <= MaxValue Then
                FoundRow = Application.WorksheetFunction.Match(VBA.Conversion.CInt(.Value), _
                            lo.ListColumns("ID").DataBodyRange, 0)
                xCombo.Value = lo.DataBodyRange(FoundRow, 2).Value2
            Else
                MsgBox "Digite um valor entre: " & MinValue & " e " & MaxValue, vbExclamation
            End If
        Else
            MsgBox "Digite um valor válido (ID númerico)", vbExclamation
        End If
    End With
End Sub

Public Sub RotineArray(ByVal xCombo As MSForms.ComboBox)
    Dim ws       As Worksheet
    Dim lo       As ListObject
    Dim FoundRow As Long
    Dim MinValue As Long
    Dim MaxValue As Long
    Dim MyArray  As Variant
    Dim lCtrl    As Long
    
    Set ws = Planilha1
    Set lo = ws.ListObjects("tbProdutos")
    MinValue = Application.WorksheetFunction.Min(lo.ListColumns("ID").DataBodyRange.Value2)
    MaxValue = Application.WorksheetFunction.Max(lo.ListColumns("ID").DataBodyRange.Value2)
    MyArray = lo.ListColumns("ID").DataBodyRange.Value
    
    With Me.txtName
        If .Value <> vbNullString And VBA.IsNumeric(.Value) Then
            If .Value >= MinValue And .Value <= MaxValue Then
                For lCtrl = LBound(MyArray) To UBound(MyArray)
                    If MyArray(lCtrl, 1) = VBA.Conversion.CInt(Me.txtName.Value) Then
                        FoundRow = lCtrl
                        Exit For
                    End If
                Next lCtrl
                xCombo.Value = lo.DataBodyRange(FoundRow, 2).Value2
            Else
                MsgBox "Digite um valor entre: " & MinValue & " e " & MaxValue, vbExclamation
            End If
        Else
            MsgBox "Digite um valor válido (ID númerico)", vbExclamation
        End If
    End With
End Sub

Public Sub RotineCollection(ByVal xCombo As MSForms.ComboBox)
    Dim ws       As Worksheet
    Dim lo       As ListObject
    Dim FoundRow As Long
    Dim MinValue As Long
    Dim MaxValue As Long
    Dim lCtrl    As Long
    Dim oDic     As Collection
    
    Set ws = Planilha1
    Set lo = ws.ListObjects("tbProdutos")
    MinValue = Application.WorksheetFunction.Min(lo.ListColumns("ID").DataBodyRange.Value2)
    MaxValue = Application.WorksheetFunction.Max(lo.ListColumns("ID").DataBodyRange.Value2)
    Set oDic = New Collection
    
    For lCtrl = 1 To MaxValue
        oDic.Add lo.DataBodyRange(lCtrl, 1).Value2, VBA.Conversion.CStr(lo.DataBodyRange(lCtrl, 1).Value2)
    Next lCtrl
    
    With Me.txtName
        If .Value <> vbNullString And VBA.IsNumeric(.Value) Then
            If .Value >= MinValue And .Value <= MaxValue Then
                For lCtrl = 1 To MaxValue
                    If oDic.Item(lCtrl) = VBA.Conversion.CInt(Me.txtName.Value) Then
                        FoundRow = oDic.Item(lCtrl)
                        Exit For
                    End If
                Next lCtrl
                xCombo.Value = lo.DataBodyRange(FoundRow, 2).Value2
            Else
                MsgBox "Digite um valor entre: " & MinValue & " e " & MaxValue, vbExclamation
            End If
        Else
            MsgBox "Digite um valor válido (ID númerico)", vbExclamation
        End If
    End With
End Sub

Public Sub RotineDictionary(ByVal xCombo As MSForms.ComboBox)
    Dim ws       As Worksheet
    Dim lo       As ListObject
    Dim FoundRow As Long
    Dim MinValue As Long
    Dim MaxValue As Long
    Dim lCtrl    As Long
    Dim oDic     As Scripting.Dictionary
    
    Set ws = Planilha1
    Set lo = ws.ListObjects("tbProdutos")
    MinValue = Application.WorksheetFunction.Min(lo.ListColumns("ID").DataBodyRange.Value2)
    MaxValue = Application.WorksheetFunction.Max(lo.ListColumns("ID").DataBodyRange.Value2)
    Set oDic = New Scripting.Dictionary
    
    For lCtrl = 1 To MaxValue
        oDic.Add VBA.Conversion.CStr(lo.DataBodyRange(lCtrl, 1).Value2), lo.DataBodyRange(lCtrl, 1).Value2
    Next lCtrl
    
    With Me.txtName
        If .Value <> vbNullString And VBA.IsNumeric(.Value) Then
            If .Value >= MinValue And .Value <= MaxValue Then
                For lCtrl = 1 To MaxValue
                    If oDic.Item(VBA.Conversion.CStr(lCtrl)) = VBA.Conversion.CInt(Me.txtName.Value) Then
                        FoundRow = oDic.Item(VBA.Conversion.CStr(lCtrl))
                        Exit For
                    End If
                Next lCtrl
        
        xCombo.Value = lo.DataBodyRange(FoundRow, 2).Value2
            Else
                MsgBox "Digite um valor entre: " & MinValue & " e " & MaxValue, vbExclamation
            End If
        Else
            MsgBox "Digite um valor válido (ID númerico)", vbExclamation
        End If
    End With
End Sub




Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call SearchRotine
End Sub

Private Sub UserForm_Initialize()
    Dim ws      As Worksheet
    Dim lo      As ListObject
    Dim xCtrl   As MSForms.Control
    
    Set ws = Planilha1
    Set lo = ws.ListObjects(1)
    
    For Each xCtrl In Me.Controls
        If TypeOf xCtrl Is MSForms.ComboBox Then
            xCtrl.RowSource = lo.ListColumns("ID").DataBodyRange.Address(, , , 1)
        End If
    Next xCtrl
    
End Sub

Private Sub SearchRotine()
    Dim iTime As Single
    
    iTime = VBA.Timer()
    Call RotineFind(Me.cboFind)
    Debug.Print "Método Find: " & VBA.Format((VBA.Timer - iTime) * 100, "0.0000 segundos")
    
    iTime = Empty
    iTime = VBA.Timer()
    Call RotineFunctionMatch(Me.cboCorresp)
    Debug.Print "Função Match: " & VBA.Format((VBA.Timer - iTime) * 1000, "0.0000 segundos")
    
    iTime = Empty
    iTime = VBA.Timer()
    Call RotineArray(Me.cboArray)
    Debug.Print "Objeto Array: " & VBA.Format((VBA.Timer - iTime) * 1000, "0.0000 segundos")
    
    iTime = Empty
    iTime = VBA.Timer()
    Call RotineCollection(Me.cboCol)
    Debug.Print "Objeto Collection: " & VBA.Format((VBA.Timer - iTime) * 1000, "0.0000 segundos")
    
    iTime = Empty
    iTime = VBA.Timer()
    Call RotineDictionary(Me.cboDic)
    Debug.Print "Objeto Dictionary: " & VBA.Format((VBA.Timer - iTime) * 1000, "0.0000 segundos")
End Sub



