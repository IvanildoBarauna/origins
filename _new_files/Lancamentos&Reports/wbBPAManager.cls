Attribute VB_Name = "wbBPAManager"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()
    Dim conn            As WorkbookConnection
    Dim sConnection     As String
    Dim bdPath          As String
    Dim NomeTabela, TablesName
    
    TablesName = Array("tbProcedimentos", "tbConsultas")
    
    bdPath = ThisWorkbook.Path & "\" & "dbBPA.mdb"
    sConnection = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" & bdPath & ";Mode=Read;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
                            
    For Each conn In ThisWorkbook.Connections
        With conn.OLEDBConnection
            .Connection = sConnection
        End With
    Next conn
End Sub
