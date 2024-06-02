Attribute VB_Name = "Módulo1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("tbConsultas.Connection"), Version:=6). _
        CreatePivotTable TableDestination:="RelatórioConsultas!R16C6", TableName:= _
        "Tabela dinâmica2", DefaultVersion:=6
    Cells(16, 6).Select
    With ActiveSheet.PivotTables("Tabela dinâmica2")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = True
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("Tabela dinâmica2").RepeatAllLabels xlRepeatLabels
    ActiveWorkbook.ShowPivotTableFieldList = True
    Application.CutCopyMode = False
    ActiveSheet.PivotTables("Tabela dinâmica2").Location = _
        "RelatórioConsultas!$A$1"
    Range("F18").Select
    Sheets("RelatórioProcedimentos").Select
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("tbProcedimentos.Connection"), Version:=6). _
        CreatePivotTable TableDestination:="RelatórioProcedimentos!R1C1", TableName _
        :="Tabela dinâmica3", DefaultVersion:=6
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("Tabela dinâmica3")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = True
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica3").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("Tabela dinâmica3").RepeatAllLabels xlRepeatLabels
    Sheets("RelatórioProcedimentos").Select
    Application.WindowState = xlNormal
    Sheets("RelatórioProcedimentos").Select
    Range("A2").Select
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    ActiveWorkbook.Connections("tbProcedimentos.Connection").Refresh
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    Range("G12").Select
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    ActiveWorkbook.Connections("tbProcedimentos.Connection").Refresh
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    ActiveWorkbook.Connections("tbProcedimentos.Connection").Refresh
    Sheets("RelatórioConsultas").Select
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    Range("C5").Select
    Application.WindowState = xlNormal
    ChDir "D:\Junior & Cris\Desktop"
    Workbooks.Open Filename:="D:\Junior & Cris\Desktop\BPA.xlsm"
    Sheets("ReportConsultas").Select
    Range("A6").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables("dynConsultas").PivotSelect "PROFISSIONAL[All]", _
        xlLabelOnly, True
    Range("A6").Select
    ActiveSheet.PivotTables("dynConsultas").PivotSelect "CÓDIGO[All]", xlLabelOnly _
        , True
    Cells.Select
    ActiveSheet.PivotTables("dynConsultas").PivotSelect "", xlDataAndLabel, True
    Range("A4").Select
    ActiveWindow.ActivateNext
    Windows("BPA.xlsm").Activate
    Windows("Lancamentos&Reports.xlsm").Activate
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("YEAR_NUM")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("MONTH_NAME")
        .Orientation = xlPageField
        .Position = 1
    End With
    Windows("BPA.xlsm").Activate
    Range("B5").Select
    Windows("Lancamentos&Reports.xlsm").Activate
    ActiveSheet.PivotTables("Tabela dinâmica2").RowAxisLayout xlTabularRow
    Range("A2").Select
    Windows("BPA.xlsm").Activate
    Range("A7").Select
    Windows("Lancamentos&Reports.xlsm").Activate
    Windows("BPA.xlsm").Activate
    Windows("Lancamentos&Reports.xlsm").Activate
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("PROFESSIONAL")
        .Orientation = xlRowField
        .Position = 1
    End With
    Windows("BPA.xlsm").Activate
    Range("B5").Select
    ActiveSheet.PivotTables("dynConsultas").PivotSelect "CÓDIGO[All]", xlLabelOnly _
        , True
    Windows("Lancamentos&Reports.xlsm").Activate
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("PROC_CODE")
        .Orientation = xlRowField
        .Position = 2
    End With
    Windows("BPA.xlsm").Activate
    Range("B7").Select
    ActiveSheet.PivotTables("dynConsultas").PivotSelect "CBO[All]", xlLabelOnly, _
        True
    Range("C6").Select
    Windows("Lancamentos&Reports.xlsm").Activate
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("CBO_CODE")
        .Orientation = xlRowField
        .Position = 3
    End With
    Windows("BPA.xlsm").Activate
    Windows("Lancamentos&Reports.xlsm").Activate
    ActiveSheet.PivotTables("Tabela dinâmica2").AddDataField ActiveSheet. _
        PivotTables("Tabela dinâmica2").PivotFields("IDADE"), "Contagem de IDADE", _
        xlCount
    Windows("BPA.xlsm").Activate
    Windows("Lancamentos&Reports.xlsm").Activate
    Range("A4").Select
    Sheets("RelatórioProcedimentos").Select
    Range("A1").Select
    Windows("BPA.xlsm").Activate
    Sheets("ReportProcedimentos").Select
    Range("A4").Select
    Windows("Lancamentos&Reports.xlsm").Activate
    Windows("BPA.xlsm").Activate
    Windows("Lancamentos&Reports.xlsm").Activate
    With ActiveSheet.PivotTables("Tabela dinâmica3").PivotFields( _
        "NOMEPROCED_PROFISSIONAL")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica3").PivotFields("YEAR_NUM")
        .Orientation = xlPageField
        .Position = 1
    End With
    Range("A3").Select
    ActiveSheet.PivotTables("Tabela dinâmica3").RowAxisLayout xlTabularRow
    Range("A1").Select
    With ActiveSheet.PivotTables("Tabela dinâmica3").PivotFields("MONTH_NAME")
        .Orientation = xlPageField
        .Position = 1
    End With
    Windows("BPA.xlsm").Activate
    Range("B4").Select
    Windows("Lancamentos&Reports.xlsm").Activate
    With ActiveSheet.PivotTables("Tabela dinâmica3").PivotFields("CODPROC_CODCBO")
        .Orientation = xlRowField
        .Position = 2
    End With
    Windows("BPA.xlsm").Activate
    Windows("Lancamentos&Reports.xlsm").Activate
    Windows("BPA.xlsm").Activate
    Range("C8").Select
    Windows("Lancamentos&Reports.xlsm").Activate
    Windows("BPA.xlsm").Activate
    Windows("Lancamentos&Reports.xlsm").Activate
    ActiveSheet.PivotTables("Tabela dinâmica3").AddDataField ActiveSheet. _
        PivotTables("Tabela dinâmica3").PivotFields("QUANTIDADE"), _
        "Contagem de QUANTIDADE", xlCount
    With ActiveSheet.PivotTables("Tabela dinâmica3").PivotFields( _
        "Contagem de QUANTIDADE")
        .Caption = "Soma de QUANTIDADE"
        .Function = xlSum
    End With
    Range("C4").Select
    Windows("BPA.xlsm").Activate
    Range("C15").Select
    ActiveWindow.SmallScroll Down:=24
    Windows("Lancamentos&Reports.xlsm").Activate
    ActiveWorkbook.ShowPivotTableFieldList = False
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Application.CommandBars("Queries and Connections").Visible = False
    Sheets("RelatórioConsultas").Select
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.RefreshAll
    Range("B4").Select
    ActiveWorkbook.Save
    Range("I9").Select
    ActiveWindow.ActivateNext
    ActiveWorkbook.Save
    ActiveWindow.Close
    Range("E10").Select
    Sheets("RelatórioConsultas").Select
    Range("B4").Select
    ActiveWindow.SmallScroll Down:=-30
    Range("A4").Select
    ActiveWorkbook.RefreshAll
    Sheets("RelatórioConsultas").Select
    With ActiveWorkbook.Connections("tbConsultas.Connection").OLEDBConnection
        .BackgroundQuery = False
        .CommandText = Array("tbConsultas")
        .CommandType = xlCmdTable
        .Connection = Array( _
        "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=D:\Junior & Cris\Desktop\NovoBPA\DIRETÓRIO TESTE\dbBPA.mdb;Mode=Read;" _
        , _
        "Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking M" _
        , _
        "ode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Creat" _
        , _
        "e System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without " _
        , "Replica Repair=False;Jet OLEDB:SFP=False")
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SourceConnectionFile = ""
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
        .AlwaysUseConnectionFile = False
    End With
    With ActiveWorkbook.Connections("tbConsultas.Connection")
        .Name = "tbConsultas.Connection"
        .Description = _
        "Conexão realizada para trazer os dados de tbConsultas do Banco de Dados."
    End With
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    Range("A4").Select
    Sheets("RelatórioProcedimentos").Select
    Range("A5").Select
    Sheets("RelatórioConsultas").Select
    Range("A4").Select
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    ActiveWorkbook.Connections("tbProcedimentos.Connection").Refresh
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    Range("D7").Select
    Application.CutCopyMode = False
    With ActiveWorkbook.Connections("tbConsultas.Connection").OLEDBConnection
        .BackgroundQuery = False
        .CommandText = Array("SELECT * FROM tbConsultas")
        .CommandType = xlCmdSql
        .Connection = Array( _
        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=D:\Junior & Cris\Desktop\NovoBPA\DIRETÓRIO TESTE\dbBPA.mdb;Mode=Shar" _
        , _
        "e Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Da" _
        , _
        "tabase Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password=""" _
        , _
        """;Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB" _
        , _
        ":Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validat" _
        , _
        "ion=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
        )
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SourceConnectionFile = ""
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
        .AlwaysUseConnectionFile = False
    End With
    With ActiveWorkbook.Connections("tbConsultas.Connection")
        .Name = "tbConsultas.Connection"
        .Description = _
        "Conexão realizada para trazer os dados de tbConsultas do Banco de Dados."
    End With
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    ActiveWorkbook.Connections("tbConsultas.Connection").Refresh
    ActiveWorkbook.Connections("tbProcedimentos.Connection").Refresh
    ActiveWorkbook.Connections("tbProcedimentos.Connection").Refresh
End Sub
