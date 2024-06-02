Attribute VB_Name = "LINKS"
Option Explicit

Sub BAR(msg As String)

Application.StatusBar = msg

End Sub

Sub Informar(msg As String)

MsgBox msg, vbInformation, AppName

End Sub

Sub ErroCritico(msg As String)

MsgBox msg, vbCritical, AppName

End Sub
Sub CAPA()

shtCons.Activate

End Sub

Sub BASE_BCD()

shtBCD.Activate

End Sub

Sub FiltrarFN()

With shtBCD.Range("A5:AH5")

    .AutoFilter Field:=12, Criteria1:="Jefte Soares da Silva"
    .AutoFilter Field:=33, Criteria1:="Feedback"

End With

End Sub

Sub Abrir_Consolidado()

ChDir "\\10.166.2.17\shareportal\HP-CONSUMER\Supervisores\Jefte Soares\Recusas"
Workbooks.Open Filename:= _
        "\\10.166.2.17\shareportal\HP-CONSUMER\Supervisores\Jefte Soares\Recusas\RECUSAS_2017.xlsm"


End Sub

Sub Abrir_BD()

BASE_ACESS.Activate

End Sub
