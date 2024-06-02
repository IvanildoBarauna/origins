Attribute VB_Name = "formImportador"
Attribute VB_Base = "0{90C1611E-E251-4D1A-93B5-FF6962F722C7}{9EF40DA5-A0CC-4B0E-B599-1E36F4F3BD74}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Function ActivePrinterName() As String
    Application.Dialogs(xlDialogPrinterSetup).Show
    ActivePrinterName = VBA.Left(ActivePrinter, WorksheetFunction.Find(" em", _
                                                    ActivePrinter) - 1)
End Function
Private Sub cmdMudar_Click()
    Me.txtPrint.Value = ActivePrinterName
End Sub

Sub cmdSelectTXT_Click()
    Dim wbk         As Workbook
    Dim PathFile    As String
    Dim fDialog     As fileDialog
    
    Application.ScreenUpdating = False
    
    Set fDialog = Application.fileDialog(msoFileDialogFilePicker)
    If fDialog.Show Then
        Set wbk = Application.Workbooks.Open(fDialog.SelectedItems(1))
        PathFile = wbk.FullName
        Me.txtCaminhoTXT.Value = PathFile
        Me.comboSheet.List = SheetsName(wbk)
        Me.comboSheet.SetFocus
        wbk.Close savechanges:=False
    Else
        MsgBox "Nenhum arquivo selecionado", vbInformation, Me.Caption
    End If
    
    Application.ScreenUpdating = True
End Sub

Private Sub cmdImportar_Click()

End Sub

