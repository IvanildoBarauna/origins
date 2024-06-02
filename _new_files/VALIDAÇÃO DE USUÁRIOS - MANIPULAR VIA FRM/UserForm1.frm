Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{58A96E09-9F68-419C-AA3D-099A01E236AA}{81594CC2-D7CD-4DFF-939F-F1A01777D16E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Dim ws                      As Worksheet
Dim textoDigitado           As String
Dim i                       As Integer
Dim linha               As Integer

Private Sub CommandButton1_Click()
    
    Dim lin                 As Long
    Dim aut                 As Boolean
    
        Set ws = ThisWorkbook.Worksheets(2)
        aut = CheckBox1.Value
        
    With ws
        lin = .UsedRange.Rows.Count + 1
        .Cells(lin, 1).Value = TextBox1.Text
        .Cells(lin, 2).Value = TextBox2.Text
        .Cells(lin, 3).Value = TextBox3.Text
        .Cells(lin, 4).Value = TextBox4.Text
        .Cells(lin, 5).Value = TextBox5.Text
        .Cells(lin, 6).Value = aut
    End With
        ListBox1.RowSource = "A2:E" & Rows.Count
End Sub

Private Sub CommandButton2_Click()
        
        Dim lin             As Long
        Dim aut             As Boolean
        
        Set ws = ThisWorkbook.Worksheets(2)
        aut = CheckBox1.Value
        
    With ws
        lin = linha
        .Cells(lin, 1).Value = TextBox1.Text
        .Cells(lin, 2).Value = TextBox2.Text
        .Cells(lin, 3).Value = TextBox3.Text
        .Cells(lin, 4).Value = TextBox4.Text
        .Cells(lin, 5).Value = TextBox5.Text
        .Cells(lin, 6).Value = aut
    End With
        ListBox1.RowSource = "A2:E" & Rows.Count
    
End Sub

Private Sub CommandButton3_Click()
    
    Dim linhalistbox        As Integer
    Dim textoCelula         As String
    Dim resposta            As Integer
        
        Set ws = ThisWorkbook.Worksheets(2)
    
    If ListBox1.ListIndex = -1 Then
        MsgBox "Você deve selecionar o item primeiro!", vbInformation
            Exit Sub
    End If
    
    linha = 2
    linhalistbox = 0
    textoDigitado = ListBox1.List(ListBox1.ListIndex, 0)
    
    With ws
       Do While .Cells(linha, 1).Value <> Empty
            textoCelula = .Cells(linha, 1).Value
            
            If textoCelula = textoDigitado Then
                resposta = MsgBox("Você tem certeza que gostaria de [EXCLUIR] esse User?", vbQuestion + vbYesNo, "Exluir User")
                    If resposta = vbYes Then
                        Range("A" & linha).EntireRow.Select
                        Selection.Delete Shift:=xlUp
                            Exit Do
                    Else
                        Exit Sub
                    End If
            End If
            linha = linha + 1
        Loop
    End With
    
    ListBox1.RowSource = "A2:E" & Rows.Count
    
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    i = ListBox1.ListIndex
        linha = i + 2
        
    TextBox1.Text = ListBox1.List(i, 0)
    TextBox2.Text = ListBox1.List(i, 1)
    TextBox3.Text = ListBox1.List(i, 2)
    TextBox4.Text = ListBox1.List(i, 3)
    TextBox5.Text = ListBox1.List(i, 4)
    
    If Cells(linha, 6) = "Falso" Or Cells(linha, 6) = "" Then
        CheckBox1.Value = False
    Else
        CheckBox1.Value = True
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    Set ws = ThisWorkbook.Worksheets(2)
    
    ListBox1.RowSource = "A2:F" & Rows.Count
    Label2.Caption = ws.Cells(1, 1)
    Label3.Caption = ws.Cells(1, 2)
    Label4.Caption = ws.Cells(1, 3)
    Label5.Caption = ws.Cells(1, 4)
    Label6.Caption = ws.Cells(1, 5)
    CheckBox1.Caption = ws.Cells(1, 6)
End Sub
