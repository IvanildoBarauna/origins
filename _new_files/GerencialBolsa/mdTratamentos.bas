Attribute VB_Name = "mdTratamentos"
Option Explicit

Function RemoveAcentos(sString As String) As String
    Dim sAcentos    As String
    Dim sSemAcentos As String
    Dim sTemp       As String
    Dim counter     As Long
    
    sAcentos = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
    sSemAcentos = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
    sTemp = sString
         
    For counter = 1 To VBA.Strings.Len(sAcentos)
        sTemp = VBA.Replace(sTemp, VBA.Strings.Mid(sAcentos, counter, 1), VBA.Strings.Mid(sSemAcentos, counter, 1))
    Next counter
    
    RemoveAcentos = sTemp
End Function

Function AbreviaBairros(item As String)
    Dim arrBairros(1 To 4) As String
    Dim sBairro            As Variant
    Dim bairroitem         As String
    Dim newbairro          As String

    arrBairros(1) = "JARDIM"
    arrBairros(2) = "JDM"
    arrBairros(3) = "PARQUE"
    arrBairros(4) = "VILA"
    
    bairroitem = VBA.Strings.Trim(item)
    
    For Each sBairro In arrBairros
        If VBA.Strings.UCase(bairroitem) Like "*" & sBairro & "*" Then
            Select Case sBairro
                Case Is = "JARDIM", "JDM": newbairro = "JD."
                Case Is = "PARQUE": newbairro = "PQ."
                Case Is = "VILA": newbairro = "VL."
            End Select
            bairroitem = VBA.Strings.Replace(VBA.UCase(bairroitem), sBairro, newbairro)
        End If
    Next sBairro
    
    AbreviaBairros = VBA.UCase(bairroitem)
End Function

Function AbreviaLogradouro(item As String) As String
    Dim arrLOGS(1 To 6) As String
    Dim log             As Variant
    Dim newlog          As String
    Dim Address         As String
    
    arrLOGS(1) = "RUA"
    arrLOGS(2) = "AVENIDA"
    arrLOGS(3) = "ALAMEDA"
    arrLOGS(4) = "VIELA"
    arrLOGS(5) = "MARECHAL"
    arrLOGS(6) = "ESTRADA"
    
    Address = VBA.Trim(item)
    
    For Each log In arrLOGS
        If VBA.UCase(Address) Like "*" & log & "*" Then
            Select Case log
                Case Is = "RUA": newlog = "R."
                Case Is = "AVENIDA": newlog = "AV."
                Case Is = "ALAMEDA": newlog = "AL."
                Case Is = "VIELA": newlog = "VL."
                Case Is = "MARECHAL": newlog = "MAL."
                Case Is = "ESTRADA": newlog = "ESTR."
            End Select
            Address = VBA.Replace(VBA.UCase(Address), log, newlog)
        End If
    Next log
    
    AbreviaLogradouro = VBA.UCase(Address)
End Function

Function RemoveEspaços(ByVal arr) As Variant
    Dim iRow As Long
    Dim iCol As Integer
    Dim tmparr As Variant: tmparr = arr
    
    For iRow = 1 To UBound(arr, 1)
        For iCol = 1 To UBound(arr, 2)
            If iCol = 4 Then
                tmparr(iRow, iCol) = VBA.Conversion.CDate(tmparr(iRow, iCol))
            Else
                tmparr(iRow, iCol) = VBA.Trim(tmparr(iRow, iCol))
            End If
        Next iCol
    Next iRow
    
    RemoveEspaços = tmparr
End Function

