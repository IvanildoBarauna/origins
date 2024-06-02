Attribute VB_Name = "Módulo1"
Option Explicit

Private Declare PtrSafe Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

Public Sub OpenUrl(link)

    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", link)

End Sub

Sub BtOnActionCall_01(control As IRibbonControl)

    MsgBox "Adm. Fernando Garcia" + Chr(10) + Chr(10) + "Administrador de empresas, pós graduado em Desenvolvimento de Sistemas, atua na área de Planejamento Integrado de Operações (ramo de Óleo e Gás) criando Dashboards para suportar o processo decisório, presta consultoria empresarial e ministra palestras e treinamentos diversos (Business Intlligence)." + Chr(10) + Chr(10) + "e-mail: Garcia@planilheiros.com.br", vbInformation + vbOKOnly, "Quem sou:"
    
End Sub

Sub BtOnActionCall_02(control As IRibbonControl)

    OpenUrl ("http://br.linkedin.com/in/garciamartins")

End Sub

Sub BtOnActionCall_03(control As IRibbonControl)

    MsgBox "Adm. Ruy Lacerda" + Chr(10) + Chr(10) + "Administrador de empresas, graduado pela Universidade Federal do ES, atua como Analista de BI e programador VBA (ramo de Óleo e Gás) desenvolvendo ferramentas de Gestão de Pessoas e Gestão de Processos" + Chr(10) + Chr(10) + "e-mail: ruy@planilheiros.com.br", vbInformation + vbOKOnly, "Quem sou:"
    
End Sub

Sub BtOnActionCall_04(control As IRibbonControl)

    OpenUrl ("http://br.linkedin.com/in/jandersonruy")

End Sub

Sub BtOnActionCall_05(control As IRibbonControl)

    OpenUrl ("https://plus.google.com/u/1/b/114136136517250022909/+PlanilheirosBrasil/posts?hl=pt-BR")

End Sub

Sub BtOnActionCall_06(control As IRibbonControl)

    OpenUrl ("https://facebook.com/planilheiros")

End Sub

Sub BtOnActionCall_07(control As IRibbonControl)

    OpenUrl ("http://youtube.com/c/PlanilheirosBrasil")

End Sub

Sub BtOnActionCall_08(control As IRibbonControl)

    OpenUrl ("http://planilheiros.com.br/download-pastas-de-trabalho/")

End Sub

Sub BtOnActionCall_09(control As IRibbonControl)

    OpenUrl ("http://planilheiros.com.br/produto_dashboard/")

End Sub

Sub BtOnActionCall_10(control As IRibbonControl)

    OpenUrl ("http://planilheiros.com.br/cursos_presenciais/")

End Sub

Sub BtOnActionCall_11(control As IRibbonControl)

    OpenUrl ("http://planilheiros.com.br")

End Sub

Sub BtOnActionCall_12(control As IRibbonControl)

    OpenUrl ("https://www.youtube.com/playlist?list=PLWfPHxJoa7zvwYDr3O8KnH4i_K8yzAHwY")

End Sub

Sub BtOnActionCall_13(control As IRibbonControl)

    OpenUrl ("https://www.youtube.com/playlist?list=PLWfPHxJoa7ztbpYtjwyfpZKGShItfDpbZ")

End Sub

Sub BtOnActionCall_14(control As IRibbonControl)

    OpenUrl ("https://www.youtube.com/playlist?list=PLWfPHxJoa7zv_j14Lj5B8BWW3cht0qewK")

End Sub

Sub BtOnActionCall_15(control As IRibbonControl)

    OpenUrl ("https://www.youtube.com/playlist?list=PLWfPHxJoa7zvI-Ra0ftnr_-KHs_cupVS8")

End Sub

Sub BtOnActionCall_16(control As IRibbonControl)

    OpenUrl ("https://www.youtube.com/playlist?list=PLWfPHxJoa7zuuWLmZ5HhYLZPOHf-aQQ3c")

End Sub

Sub BtOnActionCall_17(control As IRibbonControl)

    OpenUrl ("https://www.youtube.com/playlist?list=PLWfPHxJoa7zvZNGQTMPHTvAnk1eBTvjBU")

End Sub

Sub BtOnActionCall_18(control As IRibbonControl)

    OpenUrl ("https://conecte.petrobras.com.br/communities/service/html/communityview?communityUuid=0d3c4107-bc21-4078-ad23-f3d748a2b15f")

End Sub

Sub BtOnActionCall_19(control As IRibbonControl)

    OpenUrl ("https://www.youtube.com/watch?v=so56hqD5OMA&list=PLWfPHxJoa7zthSaAMlt0JkJpFeVtdHzq6")

End Sub

Sub BtOnActionCall_20(control As IRibbonControl)

    OpenUrl ("https://www.youtube.com/watch?v=VlkqcX9-oGE&list=PLWfPHxJoa7zs4cGJpQQI7a1T_h6ehrb9T")

End Sub

