Attribute VB_Name = "shReportSaidas"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
Dim uRow       As Long
Dim strFilter   As String

     If Target.Address = "$B$8" Then
          With Me
               strFilter = .Range("rngVal").Value2
               uRow = .Range("K1048576").End(xlUp).Row
               On Error Resume Next
               If .AutoFilterMode Then .Range("A10:K" & uRow).AutoFilter.ShowAllData
               If strFilter = vbNullString Then
                    .Range("A10:K" & uRow).AutoFilter.ShowAllData
               Else
                    .Range("A10:K" & uRow).AutoFilter Field:=2, Criteria1:=strFilter
                    .Cells.Columns.AutoFit
               End If
          End With
     End If
End Sub
