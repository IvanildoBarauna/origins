Attribute VB_Name = "mdAux"
Option Explicit
Public Enum HeaderOption
    Manter = 1
    Remover = 0
End Enum

Private Declare Function GetWindow Lib "user32" ( _
                         ByVal hWnd As Long, _
                         ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                         ByVal lpClassName As String, _
                         ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
                         ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
                         ByVal lpsz1 As String, _
                         ByVal lpsz2 As String) As Long
Private Declare Function GetKeyboardState Lib "user32" ( _
                         pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" ( _
                         lppbKeyState As Byte) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
                         ByVal hWnd As Long, ByVal wMsg As Long, _
                         ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GW_HWNDNEXT = 2
Private Const WM_KEYDOWN As Long = &H100
Private Const KEYSTATE_KEYDOWN As Long = &H80

Private InitialKS(0 To 255) As Byte

Sub CLS()
    Dim ImmediateWindowHandle As Long
    Dim TempKS(0 To 255) As Byte

    ImmediateWindowHandle = GetImmediateWindowHandle
    If ImmediateWindowHandle = 0 Then MsgBox "Immediate Window not found."
    If ImmediateWindowHandle < 1 Then Exit Sub
    
    GetKeyboardState InitialKS(0)
    
    'Segura Ctrl
    TempKS(vbKeyControl) = KEYSTATE_KEYDOWN
    SetKeyboardState TempKS(0)
    'Envia (Ctrl)+End
    PostMessage ImmediateWindowHandle, WM_KEYDOWN, vbKeyEnd, 0&
    'Segura o Shift
    TempKS(vbKeyShift) = KEYSTATE_KEYDOWN
    SetKeyboardState TempKS(0)
    'Envia o (Ctrl)+(Shift)+Home
    PostMessage ImmediateWindowHandle, WM_KEYDOWN, vbKeyHome, 0&
    'Envia o (Ctrl)+(Shift)+Backspace
    PostMessage ImmediateWindowHandle, WM_KEYDOWN, vbKeyBack, 0&

    'Schedule cleanup code to run
    Application.OnTime Now + TimeSerial(0, 0, 0), "RetoreInitialKS"
End Sub

Sub RetoreInitialKS()
    SetKeyboardState InitialKS(0)
End Sub

Private Function GetImmediateWindowHandle() As Long
    Dim iWindow As Object
    Dim IsDocked As Boolean
    Dim VisibleState As Boolean
    Dim WindowCaption As String
    Dim PaneHandle As Long
    Dim MainHandle As Long
    Dim DockHandle As Long
    Dim DockCaption As String
    Dim ErrorNumber As Long
    
    On Error Resume Next
    WindowCaption = Application.VBE.MainWindow.Caption
    ErrorNumber = Err.Number
    On Error GoTo 0
    
    If ErrorNumber <> 0 Then
        MsgBox "Não foi possível acessar a janela de Verificação Imediata. Ela está aberta?", vbExclamation
        GetImmediateWindowHandle = -1
        Exit Function
    End If
    
    For Each iWindow In Application.VBE.Windows
        If iWindow.Type = 5 Then
            VisibleState = iWindow.Visible
            WindowCaption = iWindow.Caption
            If Not iWindow.LinkedWindowFrame Is Nothing Then
                IsDocked = True
                DockCaption = iWindow.LinkedWindowFrame.Caption
            End If
            
            Exit For
        End If
    Next iWindow
    
    MainHandle = FindWindow("wndclass_desked_gsk", WindowCaption)
    
    If IsDocked Then
        PaneHandle = FindWindowEx(MainHandle, 0&, "VbaWindow", WindowCaption)
        If PaneHandle = 0 Then
            'Painel flutuante
            DockHandle = FindWindow("VbFloatingPalette", vbNullString)
            PaneHandle = FindWindowEx(DockHandle, 0&, "VbaWindow", WindowCaption)
            While DockHandle > 0 And PaneHandle = 0
                DockHandle = GetWindow(DockHandle, GW_HWNDNEXT)
                PaneHandle = FindWindowEx(DockHandle, 0&, "VbaWindow", WindowCaption)
            Wend
        End If
    ElseIf VisibleState Then
        DockHandle = FindWindowEx(MainHandle, 0&, "MDIClient", vbNullString)
        DockHandle = FindWindowEx(DockHandle, 0&, "DockingView", vbNullString)
        PaneHandle = FindWindowEx(DockHandle, 0&, "VbaWindow", WindowCaption)
    Else
        PaneHandle = FindWindowEx(MainHandle, 0&, "VbaWindow", WindowCaption)
    End If
    
    GetImmediateWindowHandle = PaneHandle
End Function

Public Sub DeleteUsedRange(ByVal ws As Worksheet, _
                           ByVal InitialRow As Long, _
                           ByVal InitialColumn As Long, _
                           ByVal DeslocDirection As XlDirection, _
                           ByVal HeaderOP As HeaderOption)
    Dim LastRow     As Long
    Dim LastColumn  As Long
    With ws
        LastRow = .Cells(.Rows.Count, InitialColumn).End(xlUp).Row
        LastColumn = .Cells(InitialRow, .Columns.Count).End(xlToLeft).Column
        If HeaderOP = Remover And InitialRow > 1 Then InitialRow = InitialRow - 1
        .Range(.Cells(InitialRow, InitialColumn), .Cells(LastRow, LastColumn)).Delete Shift:=DeslocDirection
    End With
End Sub

Public Sub DeleteUsedRangeAllSheets()
    Dim sCount As Integer
    For sCount = 1 To ThisWorkbook.Sheets.Count
        DeleteUsedRange ThisWorkbook.Sheets(sCount), 1, 1, xlUp, Remover
    Next sCount
End Sub

' Modulo    : xlApplication / Módulo
' Rotina    : LIMPAR_MEMORIA() / Sub
' Autor     : Jefferson Dantas (jefferson@tecnun.com.br)
' Data      : 07/11/2012 - 16:42
' Revisão   : Fernando Fernandes (fernando@tecnun.com.br)
' Data      : 07/01/2013 (mdy)
' Proposta  : Remove Objetcs from memory
'---------------------------------------------------------------------------------------
Public Sub LIMPAR_MEMORIA(ParamArray Objects() As Variant)
    On Error GoTo TRATAR_ERRO
    Dim Counter As Integer
    
    For Counter = 0 To UBound(Objects) Step 1
        Select Case TypeName(Objects(Counter))
            Case "Boolean"
                Objects(Counter) = False
            Case "Variant"
                If VBA.IsArray(Objects(Counter)) Then Erase Objects(Counter)
                    Objects(Counter) = Empty
            Case "String"
                Objects(Counter) = vbNullString
            Case "Worksheet"
                Set Objects(Counter) = Nothing
            Case "Workbook"
                Objects(Counter).Close SaveChanges:=False
                Set Objects(Counter) = Nothing
            Case "Connection", "Database", "Recordset2", "Recordset"
                Objects(Counter).Close
                Set Objects(Counter) = Nothing
            Case Else
                Set Objects(Counter) = Nothing
            If VBA.IsObject(Objects(Counter)) Then
                Set Objects(Counter) = Nothing
            Else
                Objects(Counter) = Empty
            End If
        End Select
    Next Counter
    Exit Sub
    On Error GoTo 0
TRATAR_ERRO:
    Resume Next
    Debug.Print "DEU ERRO NA LIMPEZA DE MEMÓRIA: " & VBA.TypeName(Objects(Counter))
End Sub

