Attribute VB_Name = "mdMain"
Public Function LastRefresh(sPath As String) As Date
    Dim FSO As Scripting.FileSystemObject: Set FSO = New Scripting.FileSystemObject
    LastRefresh = FSO.GetFile(ThisWorkbook.FullName).DateLastModified
End Function
