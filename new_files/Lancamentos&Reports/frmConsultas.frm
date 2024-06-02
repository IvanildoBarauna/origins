Attribute VB_Name = "frmConsultas"
Attribute VB_Base = "0{6F3FB6CE-59F9-4E0E-9377-578EEAB5C50B}{F76C567B-4B50-47D0-AE3D-64176EFBF89A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub btnExcluir_Click()
#If booAux Then
    Dim conn        As ADODB.Connection
    Dim rs          As ADODB.Recordset
#Else
    Dim conn        As Object
    Dim rs          As Object
#End If
    Dim selected    As Long
    Dim deleteitem  As Long
    
    selected = Me.lstConsultas.ListIndex
        
    If selected >= 0 Then
        deleteitem = Me.lstConsultas.List(selected, 0)
        If MsgBox("Tem certeza que deseja excluir o item selecionado do banco de dados?", vbQuestion + vbYesNo) = vbYes Then
            Set conn = DataBaseConnection()
            Set rs = myRecordSet()
            rs.Open "DELETE * from tbConsultas where ID =" & deleteitem, conn
            Call UpdateListBox
            conn.Close
            MsgBox "Item deletado com sucesso do banco de dados.", vbInformation, Me.Caption
        End If
    End If
End Sub

Private Sub txt_databpa_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Me.txt_databpa.MaxLength = 10
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
            Select Case VBA.Len(Me.txt_databpa.Value)
                Case 2, 5
                    Me.txt_databpa.SelText = "/"
            End Select
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txt_nascto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Me.txt_nascto.MaxLength = 10
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
            Select Case VBA.Len(Me.txt_nascto.Value)
                Case 2, 5
                    Me.txt_nascto.SelText = "/"
            End Select
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub UserForm_Initialize()
    Dim ws      As Excel.Worksheet
    Dim lo      As Excel.ListObject
    
    Set ws = wsCadastros
    Set lo = ws.ListObjects("tbCadastroConsultas")
    wsView.Activate
    With Me
        .cbo_prof.RowSource = lo.Application.Range(lo.DataBodyRange(1, 2), _
                                lo.Range(lo.ListRows.Count, 2)).Address(external:=1)
        .txt_databpa = StartDateBPA()
        .cbo_prof.SetFocus
        Call UpdateListBox
    End With
End Sub

Private Sub btnLan_Click()
    Dim oConsulta As cFichaConsulta
    
    Set oConsulta = New cFichaConsulta
    
    'Validações de campos antes de entrar na classe
    If ValidateEmptyControls(Me) Or _
       Not Me.cbo_prof.Value <> "" Or _
       Not VBA.IsDate(Me.txt_nascto.Value) Or _
       Not VBA.IsDate(Me.txt_databpa.Value) Then Exit Sub
    
    With oConsulta
        .NomeProfissional = Me.cbo_prof.Value
        .DataNascimento = Me.txt_nascto.Value
        .DataInicial = Me.txt_databpa.Value
        .InsertData
        Call UpdateListBox
    End With
    
    Call ClearFields(Me)
    Me.lbValida.Caption = "Registro Lançado!"
    Me.cbo_prof.SetFocus
End Sub

Public Function ValidateEmptyControls(ByRef FRM As UserForm) As Boolean
    Dim xControl As MSForms.Control
    Dim sList    As String
    
    For Each xControl In FRM.Controls
        Select Case TypeName(xControl)
            Case "TextBox", "ComboBox"
                If xControl.Value = vbNullString Then
                    If Not ValidateEmptyControls Then _
                        ValidateEmptyControls = True
                    sList = sList & vbNewLine & xControl.Tag
                End If
        End Select
    Next xControl
    
    If ValidateEmptyControls Then MsgBox "Preencha os campos abaixo:" _
        & vbNewLine & sList
End Function

Private Sub UpdateListBox()
#If booAux Then
    Dim conn        As ADODB.Connection
    Dim rs          As ADODB.Recordset
#Else
    Dim conn        As Object
    Dim rs          As Object
#End If
    Dim arrAux  As Variant
    
    Set conn = DataBaseConnection()
    Set rs = myRecordSet()
    
    rs.Open "SELECT ID, PROFESSIONAL, BORN_DATE, IDADE, INITIAL_DATE FROM tbConsultas", conn
    
    If Not rs.EOF Then
        arrAux = rs.GetRows
        conn.Close
        arrAux = Array2DTranspose(arrAux)
        Me.lstConsultas.ColumnCount = UBound(arrAux, 2) + 1
        Me.lstConsultas.List = arrAux
    Else
        Me.lstConsultas.Clear
    End If
    
End Sub
