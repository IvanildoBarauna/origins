Attribute VB_Name = "cFichaConsulta"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Const TableName         As String = "tbConsultas"

Private Type ClassType
    ProfName            As String
    DataNasto           As Date
    InitialDate         As Date
End Type

Private This As ClassType

Public Property Let NomeProfissional(Value As String): This.ProfName = Value: End Property
Public Property Let DataNascimento(Value As Date): This.DataNasto = Value: End Property
Public Property Let DataInicial(Value As Date): This.InitialDate = Value: End Property

Sub InsertData()
#If booAux Then
    Dim conn        As ADODB.Connection
    Dim rs          As ADODB.Recordset
#Else
    Dim conn        As Object
    Dim rs          As Object
#End If

    Set conn = DataBaseConnection()
    Set rs = myRecordSet()
    
    rs.Open SQL_INSERTINTO, conn
    
    conn.Close
    Set conn = Nothing
    Set rs = Nothing
End Sub

Private Function SQL_INSERTINTO() As String
    Dim myCMD       As String
    Dim FinalDate   As Date
    Dim Idade       As Integer
    Dim YNum        As Integer
    Dim MesNome     As String
    Dim ProcCode    As Double
    Dim CBOCode     As Double
    Dim Codes       As String
    
    With This
        Codes = ConsultaCodes(.ProfName)
        FinalDate = VBA.DateSerial(VBA.Year(.InitialDate), VBA.Month(.InitialDate) + 1, 20)
        Idade = VBA.Int((VBA.Date - .DataNasto) / 365.25)
        YNum = VBA.Year(FinalDate)
        MesNome = VBA.UCase(VBA.Format(FinalDate, "mmmm"))
        ProcCode = VBA.Split(Codes, ";")(0)
        CBOCode = VBA.Split(Codes, ";")(1)
    End With
    
    myCMD = "INSERT INTO " & TableName
    myCMD = myCMD & " (REG_TIME, PROFESSIONAL, BORN_DATE, IDADE, INITIAL_DATE, FINAL_DATE, YEAR_NUM, MONTH_NAME, PROC_CODE, CBO_CODE)"
    myCMD = myCMD & " VALUES "
    myCMD = myCMD & "('" & VBA.Now & "', "
    myCMD = myCMD & "'" & VBA.UCase(This.ProfName) & "', "
    myCMD = myCMD & " '" & This.DataNasto & "', "
    myCMD = myCMD & Idade & ", "
    myCMD = myCMD & " '" & This.InitialDate & "', "
    myCMD = myCMD & " '" & FinalDate & "', "
    myCMD = myCMD & YNum & ", "
    myCMD = myCMD & "'" & MesNome & "', "
    myCMD = myCMD & ProcCode & ", "
    myCMD = myCMD & CBOCode & ")"
    
    SQL_INSERTINTO = myCMD
End Function

