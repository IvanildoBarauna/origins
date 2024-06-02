Attribute VB_Name = "wsSimulados"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo errformat
    If Target.Column = 1 Then
        With Me.ListObjects("tbSimulados")
            Excel.Application.EnableEvents = False
            .Range.ClearFormats
            .TableStyle = "TableStyleMedium9"
            Excel.Application.EnableEvents = True
        End With
        GoTo UpdateTable
errformat:
        MsgBox "Não foi possível formatar a tabela e o relatório não foi atualizado.", vbCritical
        Exit Sub
UpdateTable:
        On Error GoTo errUpdate
        wsReports.PivotTables(1).RefreshTable
        MsgBox "A tabela de dados foi alterada e os relatórios foram atualizados!", vbInformation
        Exit Sub
    End If
errUpdate:
    MsgBox "A tabela foi formatada mas não foi possível atualizar o relatório.", vbExclamation
End Sub
