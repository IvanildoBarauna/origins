Attribute VB_Name = "Módulo7"
Sub Selecionar_Imprimir01()



 Application.Dialogs(xlDialogPrinterSetup).Show
   
   Sheets("RELATÓRIO_DÍZIMO").Range("A2:L" & Range("L1048576").End(xlUp).Row).PrintOut Copies:=1, Collate:=True



'On Error GoTo TratarErro
    
 '  Sheets("RELATÓRIO").Select
  'Range("AF1048576").End(xlUp).Select
   '
    'Dim uRow     As Integer

     ' uRow = Range("AF1048576").End(xlUp).Row

      '  Range("V2:AF" & uRow).Select
    
       ' Application.Dialogs(xlDialogPrinterSetup).Show
    'Selection.PrintOut Copies:=1, Collate:=True
    
         

'TratarErro:

                
 '        Range("V9").Select


    'Sheets("RELATÓRIO").Range("V2:AF" & Range("AF1048576").End(xlUp).Row).Application.Dialogs(xlDialogPrinterSetup).Show.PrintOut Copies:=1, Collate:=True
                   
End Sub

Sub Selecionar_Imprimir02()



 Application.Dialogs(xlDialogPrinterSetup).Show
   
   Sheets("RELATÓRIO_SAÍDAS").Range("A2:K" & Range("K1048576").End(xlUp).Row).PrintOut Copies:=1, Collate:=True


                   
End Sub

