Attribute VB_Name = "frmExportReport"
Attribute VB_Base = "0{EDFF51F9-2143-4052-A9B1-55305177F608}{32CC8666-3141-4990-B1A2-06B8AB6E56E8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
#Const booAux = False

Private Sub CarregaComboAno(Source As String)
#If booAux Then
    Dim oDic As Scripting.Dictionary: Set oDic = New Scripting.Dictionary
#Else
    Dim oDic As Object: Set oDic = VBA.CreateObject("Scripting.Dictionary")
#End If
    Dim loConsultas     As ListObject
    Dim loProcedimentos As ListObject
    Dim lo              As ListObject
    Dim counter         As Long
    Dim item            As String
    
    'Set loConsultas = wsConsultas.ListObjects("tbConsultas")
    Set loProcedimentos = wsProcedimentos.ListObjects("tbProcedimentos")
    Set lo = VBA.IIf(Source = "Consultas", loConsultas, loProcedimentos)
    
    For counter = 1 To lo.ListRows.Count
        item = lo.DataBodyRange(counter, lo.ListColumns("ANO").Index).Value2
        If Not oDic.Exists(item) Then oDic.Add item, item
    Next counter
    
    Me.ComboAno.List = oDic.Items
End Sub

Private Sub CarregaComboMês(Consultas As Boolean)
#If booAux Then
    Dim oDic As Scripting.Dictionary: Set oDic = New Scripting.Dictionary
#Else
    Dim oDic As Object: Set oDic = VBA.CreateObject("Scripting.Dictionary")
#End If
    Dim loConsultas     As ListObject
    Dim loProcedimentos As ListObject
    Dim lo              As ListObject
    Dim counter         As Long
    Dim itemAno         As String
    Dim itemMês         As String
    
    'Set loConsultas = wsConsultas.ListObjects("tbConsultas")
    Set loProcedimentos = wsProcedimentos.ListObjects("tbProcedimentos")
    Set lo = VBA.IIf(Consultas, loConsultas, loProcedimentos)
    
    For counter = 1 To lo.ListRows.Count
        itemAno = lo.DataBodyRange(counter, lo.ListColumns("ANO").Index).Value2
        itemMês = lo.DataBodyRange(counter, lo.ListColumns("MÊS").Index).Value2
        If Not oDic.Exists(itemMês) And itemAno = Me.ComboAno.Value Then
            oDic.Add itemMês, itemMês
        End If
    Next counter
    
    Me.ComboMês.List = oDic.Items
End Sub

Private Sub btnExport_Click()
    If Not ValidateEmptyControls(Me) Then
        If Me.obtnConsultas.Value And Me.chkPDF.Value Then
            Call ExportReport(Consultas, Me.ComboAno.Value, Me.ComboMês.Value, xlYes)
        ElseIf Me.obtnConsultas.Value And Not Me.chkPDF.Value Then
            Call ExportReport(Consultas, Me.ComboAno.Value, Me.ComboMês.Value, xlNo)
        ElseIf Me.obtnProcedimentos.Value And Me.chkPDF.Value Then
            Call ExportReport(Procedimentos, Me.ComboAno.Value, Me.ComboMês.Value, xlYes)
        Else
            Call ExportReport(Procedimentos, Me.ComboAno.Value, Me.ComboMês.Value, xlNo)
        End If
    End If
    If Not MsgBox("Deseja realizar mais alguma operação?", vbQuestion + vbYesNo) = vbYes Then
        VBA.Unload Me
    End If
End Sub

Private Sub chkPDF_Click(): Me.chkPDF.Caption = VBA.IIf(Me.chkPDF.Value, "Gerar PDF!", "Gerar PDF?"): End Sub
Private Sub ComboAno_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call CarregaComboMês(Me.obtnConsultas): End Sub
Private Sub fmReports_Enter(): Me.ComboAno.Value = "": Me.ComboMês.Value = "": End Sub

Private Sub fmReports_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not Me.obtnConsultas.Value And Not Me.obtnProcedimentos.Value Then
        Cancel = True
    Else
        If Me.obtnConsultas Then Call CarregaComboAno("Consultas") Else Call CarregaComboAno("Procedimentos")
    End If
End Sub

Private Sub UserForm_Initialize(): wsView.Activate: End Sub
