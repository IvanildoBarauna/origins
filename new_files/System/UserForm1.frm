Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{35523A08-EB27-49B4-8058-E8209655508E}{55EF2668-43CE-4242-A294-74AA4981319E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim arr As Variant
    
    arr = Array2DTranspose(mdConexoes.mRecordSet("SELECT * FROM [Planilha1$]").GetRows())
    
    Me.ListBox1.ColumnCount = UBound(arr, 2)
    Me.ListBox1.List = arr
End Sub
