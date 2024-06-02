Attribute VB_Name = "mdImportExport"
Option Explicit
Public Sub FileImport()
    Dim fDlg        As Office.FileDialog
    Dim strFile     As String
    Dim wbk         As Excel.Workbook
    Dim wsDestino   As Excel.Worksheet, wsFonte As Excel.Worksheet
    Dim lRowDestino As Long, lRowFonte    As Long
    Dim lo          As Excel.ListObject
    Dim lCol        As Integer
    Dim counter     As Integer
    Dim sItem       As String
    Dim sBairro     As String
    Dim arrDados, tmparr
    
    Set fDlg = Application.FileDialog(FileDialogType:=msoFileDialogOpen)
    Set wsDestino = shBD
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If wsDestino.ListObjects.Count > 0 Then Set lo = wsDestino.ListObjects(1)
    
    With fDlg
        .ButtonName = "Importar"
        .Title = "Importação Relatório Estatístico Bolsa Família"
        .InitialFileName = ThisWorkbook.Path
        .Filters.Add "Arquivo CSV Bolsa Família", "*.csv", 1
    End With
    
    With wsDestino
        If fDlg.Show Then
            On Error Resume Next
            lo.Delete
            On Error GoTo 0
            .Cells.Delete
            strFile = fDlg.SelectedItems(1)
            Set wbk = Workbooks.Open(FileName:=strFile)
            Set wsFonte = wbk.Sheets(1)
            wsFonte.Rows("1:5").Delete
            lRowFonte = wsFonte.Cells(wsFonte.Rows.Count, 1).End(xlUp).Row
            arrDados = wsFonte.Range("A1:A" & lRowFonte).Value2
            wbk.Close False
            .Range("A1").Resize(UBound(arrDados, 1), 1).Value2 = arrDados
            .Range("A1:A" & UBound(arrDados, 1)).TextToColumns Destination:=.Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, Semicolon:=True
            lCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            Set lo = .ListObjects.Add(xlSrcRange, .Range("A1").Resize(UBound(arrDados, 1), lCol), , xlYes, .Range("A1"))
            
            With lo
                .Name = "tbReport"
                .ListColumns(1).Name = "Nome"
                .ListColumns(2).Name = "NIS"
                .ListColumns(3).Name = "Perfil"
                .ListColumns(4).Name = "Data Nascto"
                .ListColumns(5).Name = "Situação"
                .ListColumns(6).Name = "Endereço"
                .ListColumns(7).Name = "Bairro"
                .ListColumns("EAS ").Delete
                .ListColumns("Profissional ").Delete
                
                For counter = 1 To .ListRows.Count
                    sItem = .DataBodyRange(counter, .ListColumns("Endereço").index).Value2
                    .DataBodyRange(counter, .ListColumns("Endereço").index).Value2 = AbreviaLogradouro(RemoveAcentos(sItem))
                    sBairro = .DataBodyRange(counter, .ListColumns("Bairro").index).Value2
                    .DataBodyRange(counter, .ListColumns("Bairro").index).Value2 = AbreviaBairros(sBairro)
                Next counter
                
                .Range.EntireColumn.AutoFit
                .ListColumns(6).Range.EntireColumn.ColumnWidth = 66.57
                .DataBodyRange.Value2 = RemoveEspaços(.DataBodyRange.Value)
                .ListColumns(4).DataBodyRange.NumberFormat = "DD/MM/YYYY"
                Call ConsolidaReportPertencentes
                MsgBox "Importação e Consolidação de dados realizada com sucesso.", vbInformation
            End With
        Else
            MsgBox "Não foi selecionado nenhum arquivo, operação cancelada.", vbExclamation, "Report Import"
        End If
    End With
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub ExportDataAllAgents()
    Dim FolderPath  As String
    Dim FSO         As Object
    Dim srcFolder   As Object
    Dim iFile       As Object
    Dim loAgents    As Excel.ListObject
    Dim loReport    As Excel.ListObject
    Dim iCounter    As Long
    Dim iAgent      As String
    Dim sPath       As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
    Set loAgents = wsListaAgents.ListObjects(1)
    Set loReport = wsRuasAgents.ListObjects(1)
    FolderPath = ThisWorkbook.Path & "\RELATORIOSPDF"
    
    If Not FSO.FolderExists(FolderPath) Then
        FSO.CreateFolder (FolderPath)
    Else
        Set srcFolder = FSO.GetFolder(FolderPath)
        For Each iFile In srcFolder.Files
            iFile.Delete
        Next iFile
    End If
    
    For iCounter = 1 To loAgents.ListRows.Count
        iAgent = loAgents.DataBodyRange(iCounter, loAgents.ListColumns("NOME").index).Value2
        FiltrarDados (iAgent)
        sPath = FolderPath & "\" & iAgent & ".pdf"
        shBD.ExportAsFixedFormat xlTypePDF, sPath
        loAgents.DataBodyRange(iCounter, loAgents.ListColumns("ULT EXPORTACAO").index).Value = VBA.Now
    Next iCounter
    
    shBD.ListObjects(1).Range.AutoFilter 6
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Relatórios foram exportados com sucesso, verifique a pasta." & FolderPath, vbInformation
End Sub

Public Sub ExportaImpresso(AgentName As String, ImpressosQuantidade As Integer)
#If booAux Then
    Dim appword As Word.Application
    Dim docword As Word.Document
    Set appword = New Word.Application
#Else
    Dim appword As Object
    Dim docword As Object
    Set appword = VBA.CreateObject("Word.Application")
#End If
    Dim iCounter    As Integer
    Dim sPath       As String
    Set docword = appword.Documents.Open("D:\Junior & Cris\Desktop\BolsaFamiliaNew\Formulário de Coleta de Dados Bolsa Família.doc")
    
    sPath = "D:\Junior & Cris\Desktop\BolsaFamiliaNew\RELATORIOSPDF\IMPRESSO_" & AgentName & ".pdf"
    
    For iCounter = 1 To ImpressosQuantidade
        With docword
            .Shapes(1).TextFrame.TextRange.Text = AgentName
            .ExportAsFixedFormat sPath, wdExportFormatPDF
        End With
    Next iCounter
    
    docword.Close False
    appword.Quit
    Set docword = Nothing
    Set appword = Nothing
End Sub
Public Sub ConsolidaReportPertencentes()
    Dim ws              As Excel.Worksheet
    Dim lo              As Excel.ListObject
    Dim mtz             As Variant, ArrCriterias
    Dim rng             As Excel.Range
    Dim iCounter        As Long
    Dim AddressItem     As String
    Dim iCriteria       As Variant
    Dim AddressAgent    As String
    Dim iCounterAgent   As Long, aux As Long
    Dim loRuas          As Excel.ListObject
    Dim MyCustomArray   As Variant
    
    Set loRuas = wsRuasAgents.ListObjects(1)
    Set ws = wsPertencentes
    mtz = shBD.ListObjects(1).Range.Value
            
    If ws.ListObjects.Count > 0 Then
        ws.ListObjects(1).Delete
        ws.Cells.Delete
    End If
    
    Set rng = ws.Range("A1").Resize(UBound(mtz, 1), UBound(mtz, 2))
    rng.Value = mtz
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes, ws.Range("A1"))
    lo.Name = "tbPertencentes"
    lo.Range.Columns.EntireColumn.AutoFit
    MyCustomArray = ArrExceptions()
    
    For iCounter = 1 To lo.ListRows.Count
        AddressItem = lo.DataBodyRange(iCounter, lo.ListColumns("Endereço").index).Value2
            For Each iCriteria In MyCustomArray
                If AddressItem Like "*" & iCriteria & "*" Then
                    lo.ListRows(iCounter).Range.ClearContents
                End If
            Next iCriteria
    Next iCounter
    lo.ListColumns(2).DataBodyRange.SpecialCells(xlCellTypeBlanks).Rows.Delete
    Exit Sub
End Sub

Private Function ArrExceptions() As Variant: ArrExceptions = wsCriterias.ListObjects("tbCriteria").DataBodyRange.Value2: End Function
