Attribute VB_Name = "Módulo1"
Option Explicit
Sub arrayTo500CSv()
    Dim w           As Excel.Worksheet
    Dim wd          As Excel.Worksheet
    Dim LastRow     As Long
    Dim xTimes      As Double
    Dim xRange      As Range
    Dim count       As Long
    Dim arr         As Variant
    Dim i           As Long

    Set w = Planilha1
    Set wd = Planilha2
    Let LastRow = w.Range("A" & w.Rows.count).End(xlUp).Row
    Let xTimes = IIf((LastRow Mod 500) > 0, Round(LastRow / 500, 0) + 1, Round(LastRow / 500))
    
    Set xRange = w.Range("A2:A" & LastRow)
    arr = xRange.Value
    
    For i = 1 To xTimes Step 1
        wd.Range("A2").Resize(500, 1) = xlTrasposseArray(arrValue(LastRow, i, arr))
        ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & i * 500, FileFormat:=xlCSV, CreateBackup:=False
        wd.Range("A2").Resize(500, 1).ClearContents
    Next i

End Sub

Function arrValue(totalRows As Long, nCount As Long, arr As Variant) As Variant
    
    Dim arrTemp As Variant
    Dim i       As Long
    
    ReDim arrTemp(nCount To nCount + 499)
    
    For i = nCount To nCount + 499 Step 1
        arrTemp(i) = arr(i, 1)
    Next i
    
    arrValue = arrTemp
    Erase arrTemp
End Function

Function xlTrasposseArray(arr As Variant) As Variant
    
    Dim x               As Long
    Dim tempArray       As Variant
    
    ReDim tempArray(1 To UBound(arr, 1), 1 To 1)
    
    For x = 1 To UBound(arr, 1) Step 1
        tempArray(x, 1) = arr(x)
    Next x

    xlTrasposseArray = tempArray
End Function
