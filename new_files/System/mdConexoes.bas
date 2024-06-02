Attribute VB_Name = "mdConexoes"
Option Explicit
Function mRecordSet(SQLCommand As String) As ADODB.Recordset
    On Error GoTo errhandler
    Const sProvider As String = "Microsoft.Jet.OLEDB.4.0"
    Dim sPath       As String
    Dim conn        As ADODB.Connection
    Dim rs          As ADODB.Recordset
    Dim sConn       As String
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    sConn = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & ThisWorkbook.FullName & ";HDR=Yes';"
    
    With conn
        .ConnectionString = sConn
        .Open
    End With
    
    rs.Open SQLCommand, conn
    Set mRecordSet = rs
    GoTo ExitPoint
errhandler:
    Debug.Print Err.Description & " ERRO Nº: " & Err.Number
    Exit Function
ExitPoint:
    conn.Close
End Function

Function Array2DTranspose(avValues As Variant) As Variant
    Dim lThisCol As Long, lThisRow As Long
    Dim lUb2 As Long, lLb2 As Long
    Dim lUb1 As Long, lLb1 As Long
    Dim avTransposed As Variant
    If IsArray(avValues) Then
        On Error GoTo ErrFailed
        lUb2 = UBound(avValues, 2)
        lLb2 = LBound(avValues, 2)
        lUb1 = UBound(avValues, 1)
        lLb1 = LBound(avValues, 1)
        ReDim avTransposed(lLb2 To lUb2, lLb1 To lUb1)
        For lThisCol = lLb1 To lUb1
            For lThisRow = lLb2 To lUb2
                avTransposed(lThisRow, lThisCol) = avValues(lThisCol, lThisRow)
            Next
        Next
    End If
    Array2DTranspose = avTransposed
    Exit Function
ErrFailed:
    Debug.Print Err.Description
    Debug.Assert False
    Array2DTranspose = Empty
    Exit Function
    Resume
End Function


