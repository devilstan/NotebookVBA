'以不開啟 excel 檔案的方式，直接讀取資料
Private Function GetData(Path, File, Sheet, Address) 
    Dim Data$ 
    Data = "'" & Path & "[" & File & "]" & Sheet & "'!" & _ 
    Range(Address).Range("A1").Address(, , xlR1C1) 
    GetData = ExecuteExcel4Macro(Data) 
End Function 

'以物件方法開啟 excel 檔案，完整取得 excel 的物件資料結構
Function GetFileList() As Variant
    '列出與這個EXCEL同階層的檔案列表
    Dim myfilelist() As String  '索引值從0起算
    Dim filecount As Integer: filecount = 0
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.GetFolder(Application.ActiveWorkbook.Path)
    'loops through each file in the directory and prints their names and path
    For Each objFile In objFolder.Files
        'filter for filename
        If UCase(Left(objFile.Name, 3)) = "IDM" Then
            ReDim Preserve myfilelist(filecount)
            myfilelist(filecount) = objFile.Path
            filecount = filecount + 1
        End If
    Next
    GetFileList = myfilelist
End Function
