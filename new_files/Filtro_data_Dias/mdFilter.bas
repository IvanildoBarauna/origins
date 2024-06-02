Attribute VB_Name = "mdFilter"
Option Explicit

Sub DateFilter()
Attribute DateFilter.VB_ProcData.VB_Invoke_Func = " \n14"
      Dim rst As Date
      Dim uLine As Long
      rst = InputBox("Digite o número de dias anteriores desejado:", "Filtro de Data")
           
      shtData.ListObjects("tbDatas").Range.AutoFilter Field:=1, Criteria1:= _
            "<" & Date - rst
            
      uLine = shtData.Range("A1").End(xlDown).Row - 1
      
      MsgBox "Total de: " & uLine & " registros encontrados.", vbInformation
End Sub
