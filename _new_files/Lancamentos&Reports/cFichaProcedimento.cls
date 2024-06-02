Attribute VB_Name = "cFichaProcedimento"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Const TableName As String = "tbProcedimentos"
Private Type ClassType
    sProfissional As String
    sProcedimento As String
    iQuantidade   As Integer
    sDataInicial  As Date
End Type

Private This As ClassType

Public Property Let ProfissionalNome(Value As String): This.sProfissional = Value: End Property
Public Property Let ProcedimentoNome(Value As String): This.sProcedimento = Value: End Property
Public Property Let Quantidade(Value As Integer): This.iQuantidade = Value: End Property
Public Property Let DataInicial(Value As Date): This.sDataInicial = Value: End Property

Sub InsertOrSumReg(UniqueID As String)
    #If booAux Then
    Dim conn        As ADODB.Connection
    Dim rs          As ADODB.Recordset
#Else
    Dim conn        As Object
    Dim rs          As Object
#End If
    Dim SQL         As String
    Dim FXReturn    As String
    Dim Quant       As Integer
    Dim IDNum       As Long
    
    FXReturn = isExistsData(UniqueID)
    Set conn = DataBaseConnection()
    Set rs = myRecordSet()
    
    If Not FXReturn = "" Then
        Quant = VBA.Split(FXReturn, ";")(0)
        IDNum = VBA.Split(FXReturn, ";")(1)
    End If
    
    If Quant = 0 Then
        rs.Open SQL_INSERT, conn
    Else
        rs.Open SQL_UPDATE(Quant + This.iQuantidade, IDNum), conn
    End If
    
    conn.Close
End Sub

Private Function SQL_INSERT() As String
    Dim FinalDate       As Date
    Dim AnoNum          As Integer
    Dim MesNome         As String
    Dim ProcCBOCodes    As String
    Dim ProcCBONames    As String
    Dim myCMD           As String
    
    With This
        FinalDate = VBA.DateSerial(VBA.Year(.sDataInicial), VBA.Month(.sDataInicial) + 1, 20)
        AnoNum = VBA.Year(FinalDate)
        MesNome = VBA.UCase(VBA.Format(FinalDate, "mmmm"))
        ProcCBONames = .sProcedimento & "-" & .sProfissional
        ProcCBOCodes = ProcedimentosCodes(.sProcedimento, .sProfissional)
    End With
    
    myCMD = "INSERT INTO " & TableName
    myCMD = myCMD & " (PROFESSIONAL, PROCEDIMENTO, QUANTIDADE, INITIAL_DATE, FINAL_DATE, YEAR_NUM, MONTH_NAME, CODPROC_CODCBO, NOMEPROCED_PROFISSIONAL)"
    myCMD = myCMD & " VALUES "
    myCMD = myCMD & "('" & VBA.UCase(This.sProfissional) & "', "
    myCMD = myCMD & " '" & This.sProcedimento & "', "
    myCMD = myCMD & This.iQuantidade & ", "
    myCMD = myCMD & " '" & This.sDataInicial & "', "
    myCMD = myCMD & " '" & FinalDate & "', "
    myCMD = myCMD & AnoNum & ", "
    myCMD = myCMD & "'" & MesNome & "', "
    myCMD = myCMD & "'" & ProcCBOCodes & "', "
    myCMD = myCMD & "'" & ProcCBONames & "')"
    
    SQL_INSERT = myCMD
End Function

Private Function SQL_UPDATE(QuantSum As Integer, IDNum As Long)
    Dim myCMD As String
    
    myCMD = "UPDATE " & TableName
    myCMD = myCMD & " Set QUANTIDADE = " & QuantSum
    myCMD = myCMD & " WHERE ID = " & IDNum
    
    SQL_UPDATE = myCMD
End Function
