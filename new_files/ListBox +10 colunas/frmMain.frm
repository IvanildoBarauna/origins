Attribute VB_Name = "frmMain"
Attribute VB_Base = "0{6FCCBEE2-0416-40D3-828E-ABACB882EC72}{A684967A-6890-4C3D-92EE-336D4B2B8622}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub PopulaListBoxComArray()
    Dim tmpArr As Variant
    
    tmpArr = ArrtoListBox
    
    With Me.lstDados
        .ColumnCount = UBound(tmpArr, 2)
        .List = tmpArr
    End With
    
    Erase tmpArr
End Sub

Private Sub PopulaListBoxComRange()
    Dim tmpRange As Excel.Range
    
    Set tmpRange = RangeToListBox
    
    With Me.lstDados
        .ColumnCount = tmpRange.Columns.Count
        .RowSource = tmpRange.Address(, , , -1)
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim InitialTime As Single
    
    InitialTime = VBA.Timer
    Call PopulaListBoxComArray
    Debug.Print "Carregando ListBox com .List: " & _
        VBA.Format(VBA.Timer - InitialTime, "0.00 segundos")
    
    InitialTime = Empty
    InitialTime = VBA.Timer
    Call PopulaListBoxComRange
    Debug.Print "Carregando ListBox com .RowSource: " & _
        VBA.Format(VBA.Timer - InitialTime, "0.00 segundos")
    
End Sub
