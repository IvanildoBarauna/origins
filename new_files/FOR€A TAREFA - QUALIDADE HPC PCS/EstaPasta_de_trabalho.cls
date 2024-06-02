Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
LOG "ARQUIVO SALVO"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

LOG "ARQUIVO FECHADO"

shtLOG.Visible = False

ActiveWorkbook.Save

End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)

LOG "NOVA PLANILHA INSERIDA"

End Sub

Private Sub Workbook_Open()

ActiveWindow.DisplayWorkbookTabs = False
Plan6.Visible = False
Plan4.Visible = True
Plan4.Select
Range("C3").Select

Dim USER As String
Dim Caminho As String

    Caminho = "\\10.166.108.20\Equipe\Ivan Jr - FUPO\Relatórios"
    
    USER = Environ("USERNAME")

If ThisWorkbook.Path = Caminho Then

Select Case USER

'Autorizados na Abertura

'Junior&Cris = PC Pessoal
'ijuni002 = Ivanildo
'Vgome001 = Vinicius Gomes
'jteno002 = Jonatas Tenório
'sfran060 = Stefany Aparecida
'wlara004 = Wendrey Lara
'carai001 = Caio Arai
'calfr001 = Clarina Alfredo
'campa003 = Caroline Nunes
'vsous001 = Super Vinicius
'dsilv060 = Coordinator Daniela
'rsouz023 = Super Reginaldo
'mmelo002 = Super Melo
'jsilv061 = Super Jefte
'rtaga001 = Renan Tegani
'aoliv061 = Alvani Oliveira
'bsant009 = Bruno Romualdo
'jmart003 = Jessé Martins
'lpint007 = Lucas Pinto


Case Is = "Junior&Cris", "ijuni002", "vgome001", "rtaga001", "jteno002", "sfran060", "wlara004", "carai001", "calfr001", "campa003", "vsous001", "dsilv060", "rsouz023", "mmelo002", "jsilv061", "aoliv061", "jmart003"

      'AUTORIZADOS A VER A BARRA DE SHEETS
      
      Select Case USER
      
      'Autorizados Acesso Total
      
      'Junior&Cris = PC Pessoal
      'ijuni002 = Ivanildo
      'vsous001 = Super Vinicius
      'dsilv060 = Coordinator Daniela
      'rsouz023 = Super Reginaldo
      'mmelo002 = Super Melo = Retirado | Excessão
      'jsilv061 = Super Jefte
      
      
      Case "Junior&Cris", "ijuni002", "jsilv061", "vsous001", "dsilv060", "rsouz023"
      
            Plan4.Visible = False
      
            Plan5.Visible = True
      
            Plan6.Visible = True

             Plan5.Unprotect ["ftqualidadehpi222"]
            
            Sheets("HOME").Unprotect ["ftqualidadehpi222"]
            
         
            ActiveWindow.DisplayWorkbookTabs = True
    
      
      Case Else
      
                ActiveWindow.DisplayWorkbookTabs = False
     
            Plan5.Protect ["ftqualidadehpi222"]
            
            Sheets("HOME").Protect ["ftqualidadehpi222"]

                Plan5.Visible = False
      
      End Select

      LOG "ARQUIVO ABERTO"

            Home.Activate
            Range("A15").Select

        If Hour(Now()) > 0 And Hour(Now()) <= 12 Then
            MsgBox "Bom dia " & Environ("USERNAME") & "! Seja bem vindo(a)!", vbInformation
        End If
        
        If Hour(Now()) > 12 And Hour(Now()) < 18 Then
            MsgBox "Boa tarde " & Environ("USERNAME") & "! Seja bem vindo(a)!", vbInformation
        End If
        
        If Hour(Now()) >= 18 Then
            MsgBox "Boa noite " & Environ("USERNAME") & "! Seja bem vindo(a)!", vbInformation
        End If
        
        
    Case Else
        
            LOG "TENTATIVA DE ACESSO NÃO AUTORIZADO"
    
            ActiveWorkbook.Save

            MsgBox "Acesso não autorizado, procure o propietário do arquivo ou o seu supervisor!", vbCritical, AppName
    
            ActiveWorkbook.Close False
    
    End Select
    
    Else

        LOG "TENTATIVA DE ACESSO NO LOCAL: " & ThisWorkbook.Path
        
        ActiveWorkbook.Save
    
        MsgBox "Acesso não autorizado, procure o propietário do arquivo ou o seu supervisor!", vbCritical, AppName
    
        ThisWorkbook.Close False

    End If

    
End Sub


