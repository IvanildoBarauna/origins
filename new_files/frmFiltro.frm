Attribute VB_Name = "frmFiltro"
Attribute VB_Base = "0{0CB4EF09-5DFC-47CB-B59F-805E4F1BDC74}{7656C3EF-E51D-48EC-8CAD-A6A1C16395DA}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub UserForm_Initialize()
    Dim loRuas  As ListObject
    Dim ws      As Worksheet
    Dim lo      As ListObject
    Dim loBanco As ListObject
    Dim lc      As ListColumn
    
    Set ws = wsListaAgents
    Set lo = ws.ListObjects(1)
    Set loBanco = shBD.ListObjects(1)
    Set loRuas = wsRuasAgents.ListObjects(1)
    
    
    For Each lc In loBanco.ListColumns
        If lc.Range.EntireColumn.Hidden Then lc.Range.EntireColumn.Hidden = False
    Next lc
    
    With Me
        .ComboAgent.RowSource = lo.ListColumns(3).DataBodyRange.Address(external:=1)
        
        With .lstFiltro
            .ColumnCount = loBanco.ListColumns.Count
            .List = loBanco.Range.Value
        End With
        
        With .lstRuas
            .ColumnCount = loRuas.ListColumns.Count
            .List = loRuas.Range.Value
        End With
        
        .lbInforma.Caption = "Selecione um agente na lista..."
        .lbRuas.Caption = lbInforma.Caption
    End With
End Sub

Private Sub btnCancel_Click()
    Unload Me
    shBD.ListObjects(1).Range.AutoFilter 6
End Sub

Private Sub btnPrint_Click()
    Dim lo          As ListObject
    Dim ws          As Worksheet
    Dim counter     As Integer
    Dim fDialog     As FileDialog
    Dim Selected    As String
    Dim sPath       As String
    
    Set ws = shBD
    Set lo = ws.ListObjects(1)
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    If Me.lstFiltro.ListCount - 1 = lo.ListRows.Count Then
        MsgBox "Não é possível imprimir todos os dados, selecione um ACS.", vbExclamation, Me.Caption
        Exit Sub
    ElseIf fDialog.Show Then
        Selected = fDialog.SelectedItems(1)
        sPath = Selected & "\RelaçãoRuas_" & Me.ComboAgent.Value & "_" & VBA.Replace(VBA.Date, "/", ".") & ".pdf"
        lo.ListColumns(5).Range.EntireColumn.Hidden = True
        shBD.PageSetup.PrintTitleRows = "$1:$1"
        ws.ExportAsFixedFormat xlTypePDF, sPath, xlQualityStandard, False, False, OpenAfterPublish:=True
    Else
        MsgBox "O caminho para salvar o arquivo não foi selecionado corretamente.", vbExclamation
    End If
End Sub

Private Sub ComboAgent_Change()
    Call FiltrarDados(Me.ComboAgent.Value)
    Call UpdateListBoxReport
    Call UpdateListBoxRuas(Me.ComboAgent.Value)
    Me.lbInforma.Caption = Me.lbInforma.Caption & " Localizados" & " | " & Me.lstFiltro.ListCount - 1 & " sendo exibidos."
    Me.lbRuas.Caption = Me.lstRuas.ListCount - 1 & " registros localizados."
End Sub

Private Sub btnClear_Click()
    Me.ComboAgent.Value = ""
End Sub

Private Sub UpdateListBoxReport()
    Dim lo          As ListObject
    Dim RowCounter  As Long, ColCounter As Integer, counter As Long
    Dim mtzResult   As Variant
    Dim mtzSize     As Long
    Dim indexCol    As Integer
    
    Set lo = shBD.ListObjects(1)
    Me.lstFiltro.ColumnCount = lo.ListColumns.Count
    
    If Me.ComboAgent.Value = "" Then
        Me.lstFiltro.List = lo.Range.Value
        Exit Sub
    End If
    
    For counter = 1 To lo.ListRows.Count
        If Not lo.ListRows(counter).Range.Rows.Hidden Then
            mtzSize = mtzSize + 1
        End If
    Next counter
    
    ReDim mtzResult(1 To mtzSize + 1, 1 To lo.ListColumns.Count)
    
    For indexCol = 1 To lo.ListColumns.Count
        mtzResult(1, indexCol) = lo.ListColumns(indexCol).Name
    Next indexCol
    
    For counter = 1 To lo.ListRows.Count
        If Not lo.ListRows(counter).Range.Rows.Hidden Then
            RowCounter = RowCounter + 1
            For ColCounter = 1 To lo.ListColumns.Count
                mtzResult(RowCounter + 1, ColCounter) = lo.DataBodyRange(counter, ColCounter).Value
            Next ColCounter
        End If
    Next counter
    
    Me.lstFiltro.List = mtzResult
End Sub

Sub UpdateListBoxRuas(AgentName As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim RowCounter  As Long, ColCounter As Integer, counter As Long
    Dim mtzResult   As Variant
    Dim mtzSize     As Long
    Dim indexCol    As Integer
    
    Set ws = wsRuasAgents
    Set lo = ws.ListObjects(1)
    
    Me.lstRuas.ColumnCount = lo.ListColumns.Count
    
    If Me.ComboAgent.Value = "" Then
        Me.lstRuas.List = lo.Range.Value
        Exit Sub
    End If
    
    For counter = 1 To lo.ListRows.Count
        If lo.DataBodyRange(counter, 2).Value2 = AgentName Then mtzSize = mtzSize + 1
    Next counter
    
    ReDim mtzResult(1 To mtzSize + 1, 1 To lo.ListColumns.Count)
    
    For indexCol = 1 To lo.ListColumns.Count
        mtzResult(1, indexCol) = lo.ListColumns(indexCol).Name
    Next indexCol
    
    For counter = 1 To lo.ListRows.Count
        If lo.DataBodyRange(counter, lo.ListColumns("Nome Agente").index).Value2 = AgentName Then
            RowCounter = RowCounter + 1
            For ColCounter = 1 To lo.ListColumns.Count
                mtzResult(RowCounter + 1, ColCounter) = lo.DataBodyRange(counter, ColCounter).Value
            Next ColCounter
        End If
    Next counter
    
    Me.lstRuas.List = mtzResult
End Sub

Private Sub UserForm_Terminate()
    shBD.ListObjects(1).Range.AutoFilter 6
    shBD.ListObjects(1).ListColumns(5).Range.EntireColumn.Hidden = False
End Sub
