Attribute VB_Name = "EXTRAIR_BCD"
Option Explicit
Option Private Module

Sub BCD_EXTRACAO()

ACTIVATE_

Dim Inicio As Date
Dim Final As Date
Dim linha As Integer
Dim linha2 As Integer
Dim report As String

linha = shtBCD.Range("A1048576").End(xlUp).Row + 1
linha2 = shtBCDUP.Range("A1048576").End(xlUp).Row + 1

On Error GoTo Erro

BAR "Coletando informações ..."

Inicio = InputBox("Digite uma data de início para a extração do relatório no padrão 'DD/MM/AAAA'", AppName)

Erro:

    If Inicio = Empty Then

    ErroCritico "Data inicial para extração foi inserida incorretamente, informe-a no padrão 'dd/mm/aaaa"

    DEACTIVATE_

    Exit Sub

    Else

On Error GoTo Erro2

Final = InputBox("Digite uma data final para a extração do relatório no padrão 'DD/MM/AAAA'", AppName)

End If

Erro2:

    If Final = Empty Then

    ErroCritico "Data final para extração foi inserida incorretamente, informe-a no padrão 'dd/mm/aaaa"

    DEACTIVATE_

    Exit Sub

    Else

BAR "Abrindo fonte de dados ..."

report = "Interação"

ChDir ("\\10.166.2.17\shareportal\HP-CONSUMER\Relatórios\Extração relatório\BCD")
Workbooks.Open Filename:="\\10.166.2.17\shareportal\HP-CONSUMER\Relatórios\Extração relatório\BCD\extracao_bcd_NEW.xlsm"

Windows("extracao_bcd_NEW").Activate

Range("D3").Value = report

Range("D4").Value = Inicio

Range("D5").Value = Final

BAR "Realizando Extração ... Aguarde"

Application.Run "'extracao_bcd_NEW.xlsm'!Botão3_Clique"

Range("A9:AD9").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

BAR "Consolidando dados ..."

Windows("FN_REPORT").Activate

shtBCD.Select

If shtBCD.FilterMode = True Then

shtBCD.ShowAllData

End If


shtBCD.Range("A" & linha).PasteSpecial xlPasteFormulasAndNumberFormats

   shtBCD.Range("tblBCD").RemoveDuplicates Columns:=1, Header:= _
        xlYes

Application.CutCopyMode = False

Windows("extracao_bcd_NEW.xlsm").Activate

report = "Upload"

Range("D3").Value = report

Range("D4").Value = Inicio

Range("D5").Value = Final

Application.Run "'extracao_bcd_NEW.xlsm'!Botão3_Clique"

Range("A9:O9").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows("FN_REPORT").Activate

shtBCDUP.Select

If shtBCDUP.FilterMode = True Then

shtBCDUP.ShowAllData

End If

shtBCDUP.Range("A" & linha2).PasteSpecial xlPasteFormulasAndNumberFormats

   shtBCDUP.Range("tblBCD_UP").RemoveDuplicates Columns:=3, Header:= _
        xlYes

Application.CutCopyMode = False

Workbooks("extracao_bcd_NEW").Close False

ThisWorkbook.Save

Informar "Extração realizada com sucesso!"

BAR "Concluído"

End If

CAPA

DEACTIVATE_

End Sub




