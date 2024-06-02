Attribute VB_Name = "cLOG"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private aStream As ADODB.Stream
Private sfPath As String

Private Sub Class_Initialize()
      Set aStream = New ADODB.Stream
      aStream.Open
End Sub

Private Sub Class_Terminate()
      aStream.Close
      Set aStream = Nothing
End Sub

Public Property Get fPath() As String: fPath = sfPath: End Property
Public Property Let fPath(pfPath As String)
      sfPath = pfPath
      If Dir(sfPath) <> "" Then
            With aStream
                  .LoadFromFile sfPath
                  .Position = aStream.Size
            End With
      End If
End Property

Public Sub Registrar(Ação As String)
      With aStream
            .WriteText Now & vbTab
            .WriteText Environ("USERNAME") & vbTab
            .WriteText Environ("COMPUTERNAME") & vbTab
            .WriteText Ação & vbCrLf
      End With
End Sub

Public Sub Salvar()
      aStream.SaveToFile sfPath, adSaveCreateOverWrite
End Sub
