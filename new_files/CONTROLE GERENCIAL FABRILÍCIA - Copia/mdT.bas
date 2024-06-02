Attribute VB_Name = "mdT"
Option Explicit
Function ValidateEmptyControls() As Boolean
    Dim iCtrl As OLEObject
    
    For Each iCtrl In shtIN.OLEObjects
        If iCtrl.Object.Value = "..." Or iCtrl.Object.Value = "" Or iCtrl.Value = 0 Then
            ValidateEmptyControls = True
            Exit For
        End If
    Next iCtrl
End Function
Sub AbastecerEstoque()
    Dim Quantidade          As Integer
    Dim QuantidadeEstoque   As Integer
    Dim iRow                As Long
    Dim vSum                As Long
    Dim lo                  As ListObject
    
    Set lo = shtESTOQUE.ListObjects("tbESTOQUE")
    
    With lo
        If ValidateEmptyControls() Then: VBA.MsgBox "Todos os campos são obrigatórios.", vbCritical: Exit Sub
        
        iRow = shtIN.txtcod + 1
        Quantidade = shtIN.txtqtd.Value * 1
        QuantidadeEstoque = .Range(iRow, lo.ListColumns("QUANTIDADE").Index)
        vSum = QuantidadeEstoque + Quantidade
        .Range(iRow, lo.ListColumns("QUANTIDADE").Index) = vSum
        Registrar "ENTRADA"
        MsgBox "Entrada registrada com sucesso!", vbInformation, "ENTRADA DE ESTOQUE"
        End If
    End With
End Sub

Sub SaídaEstoque()
Dim Quantidade, QuantidadeEstoque, iRow, vSum As Long
Dim lo As ListObject
Set lo = shtESTOQUE.ListObjects("tbESTOQUE")

With lo

If shtOUT.txtcod = Empty Or shtOUT.txtdesc = Empty Then
MsgBox "O campo código do produto está vazio, ou digite um valor válido" _
, vbExclamation, "Dados inválidos"
Exit Sub
Else
iRow = shtOUT.txtcod + 1
Quantidade = shtOUT.txtqtd.Value * 1
QuantidadeEstoque = .Range(iRow, lo.ListColumns("QUANTIDADE").Index)
vSum = QuantidadeEstoque - Quantidade

.Range(iRow, lo.ListColumns("QUANTIDADE").Index) = vSum

Registrar "SAÍDA"

MsgBox "Saída Registrada com Sucesso!", vbInformation, "SAÍDA DE ESTOQUE"

End If
End With
End Sub

Sub AcrescentarProduto()
Dim iRow As Integer

iRow = shtESTOQUE.Range("A6").End(xlDown)

shtADD.txtcod = iRow

End Sub
