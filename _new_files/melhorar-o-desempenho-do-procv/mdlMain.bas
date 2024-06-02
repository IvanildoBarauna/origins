Attribute VB_Name = "mdlMain"
Option Explicit

Private Sub ExecutarTestes()
    Dim tblTestes As ListObject

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    Set tblTestes = wsInício.ListObjects("tblTestes")
    
    CalcularVelocidade tblTestes.ListColumns("Um PROCV Linear").DataBodyRange, [PROCVLinear]
    CalcularVelocidade tblTestes.ListColumns("Dois PROCV Binários").DataBodyRange, [PROCVBinário]
    CalcularVelocidade tblTestes.ListColumns("Dois ÍNDICE+CORRESP").DataBodyRange, [ÍNDICE_CORRESP]
    CalcularVelocidade tblTestes.ListColumns("Dois PROC").DataBodyRange, [PROC]
    
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
    MsgBox "Testes concluídos!", vbInformation
End Sub

Private Sub CalcularVelocidade(intervaloDeTeste As Range, _
                               célulaDeSaída As Range)
    Dim tempoInicial As Single
    Dim tempoFinal As Single
    Dim tempoTotal As Single
    
    tempoInicial = Timer
    intervaloDeTeste.Calculate
    tempoFinal = Timer
    tempoTotal = tempoFinal - tempoInicial
    
    célulaDeSaída = tempoTotal
    DoEvents
End Sub
