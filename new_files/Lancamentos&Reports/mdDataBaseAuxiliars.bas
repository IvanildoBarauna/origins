Attribute VB_Name = "mdDataBaseAuxiliars"
Option Explicit
#If booAux Then
    Public Function DataBaseConnection() As ADODB.Connection
        Const sProvider As String = "Microsoft.Jet.OLEDB.4.0"
        Dim sPath       As String
        
        sPath = ThisWorkbook.Path & Application.PathSeparator & "dbBPA.mdb"
        
        Set DataBaseConnection = New ADODB.Connection
    
        With DataBaseConnection
            .ConnectionString = sPath
            .Provider = sProvider
            .Open
        End With
        
    End Function
    Public Function myRecordSet() As ADODB.Recordset: Set myRecordSet = New ADODB.Recordset: End Function
#Else
    Public Function DataBaseConnection() As Object
        Const sProvider As String = "Microsoft.Jet.OLEDB.4.0"
        Dim sPath       As String
        
        sPath = ThisWorkbook.Path & Application.PathSeparator & "dbBPA.mdb"
        
        Set DataBaseConnection = VBA.CreateObject("ADODB.Connection")
        
        With DataBaseConnection
            .ConnectionString = sPath
            .Provider = sProvider
            .Open
        End With
        
    End Function
    Public Function myRecordSet() As Object: Set myRecordSet = VBA.CreateObject("ADODB.Recordset"): End Function
#End If

Public Function isExistsData(UniqueID As String) As String
#If booAux Then
    Dim rs          As ADODB.Recordset:  Set rs = New ADODB.Recordset
    Dim conn        As ADODB.Connection
#Else
    Dim rs          As Object
    Dim conn        As Object
#End If
    Dim mtz()       As Variant
    Dim iCounter    As Long
    Dim item        As String
    
    Set conn = DataBaseConnection()
    Set rs = myRecordSet()
    
    rs.Open "SELECT * FROM tbProcedimentos", conn
    
    If Not rs.EOF Then
        mtz = rs.GetRows
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
        
        For iCounter = LBound(mtz, 2) To UBound(mtz, 2)
            item = mtz(1, iCounter) & mtz(2, iCounter) & mtz(4, iCounter)
            If item = UniqueID Then
                isExistsData = mtz(3, iCounter) & ";" & mtz(0, iCounter)
                Exit For
            End If
        Next iCounter
    End If
End Function

Public Function ConsultaCodes(Professional As String) As String
    Dim lo          As Excel.ListObject
    Dim iCounter    As Integer
    
    Set lo = wsCadastros.ListObjects("tbCadastroConsultas")
    
    For iCounter = 1 To lo.ListRows.Count
        If lo.DataBodyRange(iCounter, lo.ListColumns("PROFISSIONAL").Index).Value2 = Professional Then
            ConsultaCodes = lo.DataBodyRange(iCounter, lo.ListColumns("CÓD. DO PROCED.").Index).Value2 & _
                            ";" & lo.DataBodyRange(iCounter, lo.ListColumns("Nº DE CBO").Index).Value2
        End If
    Next iCounter
End Function

Public Function ProcedimentosCodes(Procedimento As String, Professional As String) As String
    Dim iCounter        As Integer
    Dim lo              As Excel.ListObject
    Dim ProcCode        As String
    Dim ProfCode        As String
    
    Set lo = wsCadastros.ListObjects("tbCadastroProcedimento")
    
    For iCounter = 1 To lo.ListRows.Count
        If lo.DataBodyRange(iCounter, lo.ListColumns("PROCEDIMENTO").Index).Value2 = Procedimento Then
            ProcCode = lo.DataBodyRange(iCounter, lo.ListColumns("CÓD. PROCED.").Index).Value2
            Exit For
            Set lo = Nothing
            iCounter = 0
        End If
    Next iCounter
    
    Set lo = wsCadastros.ListObjects("tbCadastroProfissional")
    
    For iCounter = 1 To lo.ListRows.Count
        If lo.DataBodyRange(iCounter, lo.ListColumns("PROFISSIONAL").Index).Value2 = Professional Then
            ProfCode = lo.DataBodyRange(iCounter, lo.ListColumns("Nº: CBO").Index).Value2
            Exit For
        End If
    Next iCounter
    
    ProcedimentosCodes = ProcCode & "-" & ProfCode
End Function

Function Array2DTranspose(avValues)
    Dim lThisCol As Long, lThisRow As Long
    Dim lUb2 As Long, lLb2 As Long
    Dim lUb1 As Long, lLb1 As Long
    Dim avTransposed As Variant
    
    If VBA.Information.IsArray(avValues) Then
        On Error GoTo ErrFailed
        lUb2 = UBound(avValues, 2)
        lLb2 = LBound(avValues, 2)
        lUb1 = UBound(avValues, 1)
        lLb1 = LBound(avValues, 1)
        ReDim avTransposed(lLb2 To lUb2, lLb1 To lUb1)
        
        For lThisCol = lLb1 To lUb1
            For lThisRow = lLb2 To lUb2
                avTransposed(lThisRow, lThisCol) = avValues(lThisCol, lThisRow)
            Next lThisRow
        Next lThisCol
        
    End If
    Array2DTranspose = avTransposed
    Exit Function
ErrFailed:
    Debug.Print err.Description
    Debug.Assert False
    Array2DTranspose = Empty
    Exit Function
    Resume
End Function

Sub main()
    Dim conn            As WorkbookConnection
    Dim sConnection     As String
    Dim bdPath          As String
    
    bdPath = ThisWorkbook.Path & "\" & "dbBPA.mdb"
    sConnection = "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & _
                    bdPath & _
                        ";Mode=Read;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False"
    
    For Each conn In ThisWorkbook.Connections
        With conn.OLEDBConnection
            .Connection = sConnection
            .SourceConnectionFile = bdPath
            .AlwaysUseConnectionFile = True
        End With
    Next conn
    
End Sub
