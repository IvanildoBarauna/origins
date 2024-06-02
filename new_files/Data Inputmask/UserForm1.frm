Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{2C0E2A9B-FAAA-47C8-88DA-9A3DF81D4D44}{CC159471-42AE-4B21-8768-39276E606C5D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Dim NewString, MyString, mask As String
Dim position As Variant

Private Sub txtData_Change()
If IsNumeric(Right(txtData.Text, 2)) And Len(txtData.Text) >= 11 Then
    txtData.Text = Left(txtData.Text, Len(txtData.Text) - 1)
Else
position = txtData.SelStart
MyString = txtData.Text
pos = InStr(1, MyString, "_")
If pos > 0 Then
NewString = Left(MyString, pos - 1)
Else
NewString = MyString
End If
If Len(NewString) < 11 Then
    txtData.Text = NewString & Right(mask, Len(mask) - Len(NewString))
    txtData.SelStart = Len(NewString)
End If
End If
If Len(txtData.Text) >= 11 Then
    txtData.Text = Left(txtData.Text, 10)
End If
End Sub

Private Sub txtData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
position = txtData.SelStart
If KeyCode = 8 Then
    txtData.Text = mask
End If
End Sub

Private Sub UserForm_Initialize()
txtData.SelStart = 0
mask = "__/__/____"
txtData.Text = mask
End Sub
