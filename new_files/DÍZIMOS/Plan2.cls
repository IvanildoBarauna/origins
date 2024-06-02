Attribute VB_Name = "Plan2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Sub CalculaPlan()
'Calcula todas as planilhas
For Each Sheet In ThisWorkbook.Sheets
With Sheet
.Activate
.Calculate
End With
Next Sheet
End Sub

Sub CalculaWorkbook()
'Atualiza todos os vínculos da pasta de trabalho
ActiveWorkbook.RefreshAll
End Sub



Sub AbrirComVBA()
'grava em uma variável dados sobre o arquivo
retval = Dir("caminho e nome do arquivo com a extensão")
'verifica se o endereço informado coincide com o que foi gravado na variável
If retval = "informe aqui nome do arquivo com extensão" Then
ChDir "nome do caminho"
Workbooks.Open Filename:= _
"informe aqui caminho e nome do arquivo com a extensão" _
, ReadOnly:=False, Notify:=False, IgnoreReadOnlyRecommended:=True
End If
End Sub


