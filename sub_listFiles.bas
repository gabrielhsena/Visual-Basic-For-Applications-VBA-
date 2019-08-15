Sub ListFiles()
    
    Dim MyFolder As String
    Dim MyFile As String
    Dim j As Integer
    
    MyFolder = "G:\"
    MyFile = Dir(MyFolder & "\*.*")
    a = 0
    
    Do While MyFile <> ""
        a = a + 1
        Cells(a, 1).Value = MyFile
        MyFile = Dir
    Loop
    
End Sub


Sub RenameFiles()
    
    MyFolder = "G:\"
    
    For R = 1 To Range("A1").End(xlDown).Row
        OldFileName = MyFolder & Cells(R, 1).Value
        NewFileName = MyFolder & Cells(R, 2).Value
        On Error Resume Next
            If Not Dir(OldFileName) = "" Then Name OldFileName As NewFileName
        On Error GoTo 0
    Next
End Sub
