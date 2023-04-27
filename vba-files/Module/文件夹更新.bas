Attribute VB_Name = "文件夹更新"
Sub Copy()
    '创建对计算机文件系统访问的对象
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 定义SourceFolder的路径并在窗口输入
    Dim SourceFolder As String
    SourceFolder = InputBox("请输入源文件所在文件夹的路径：")
    
    '如果没有输入，则停止过程
    If SourceFolder = "" Then
        Exit Sub
    End If
    
    ' 定义新的DestinationFolderName并在窗口输入
    Dim DestinationFolderName As String
    DestinationFolderName = InputBox("请输入目标文件夹的文件名")
    
    '如果没有输入，则停止过程
    If DestinationFolderName = "" Then
        Exit Sub
    End If
    
    ' 定义新的NewMonth并把DestinationFolderName的月份情况赋值给它
    Dim NewMonth As String
    NewMonth = Right(DestinationFolderName, 2)
    
    ' 把SourceFolder文件夹的上一级路径赋给ParentDestinationFolder
    Dim ParentDestinationFolder As String
    ParentDestinationFolder = Left(SourceFolder, InStrRev(SourceFolder, "\"))
    
    '定义新的DestinationFolder并对其赋值
    Dim DestinationFolder As String
    DestinationFolder = ParentDestinationFolder & DestinationFolderName
    
    '使用Fso.CopyFolder，将SourceFolder文件夹复制到DestinationFolder
    fso.CopyFolder SourceFolder, DestinationFolder
    MsgBox "成功复制", vbInformation
    
    '调用Delete函数进行删除操作，参数为DestinationFolder
    Call Delete(DestinationFolder)
    MsgBox "成功删除", vbInformation
        
    '调用Rename函数进行文件重命名
    Call Rename(DestinationFolder, fso, NewMonth)
    MsgBox "成功命名", vbInformation
    
    '退出所有Wd程序
    On Error Resume Next
    For Each WdApp In GetObject(, "Word.Application")
        WdApp.Quit
    Next
    On Error GoTo 0
    
End Sub

Sub Delete(DestinationFolder As String)

    '创建对计算机文件系统访问的对象
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '遍历目标文件夹的子文件夹和文件，删除它们
    Dim Folder As Object
    Set Folder = fso.GetFolder(DestinationFolder)
    For Each Subfolder In Folder.SubFolders
        For Each File In Subfolder.Files
            fso.DeleteFile File.Path, True
        Next
        '通过将SubFolder文件的路径赋值给Delete过程的变量进行递归
        Call Delete(Subfolder.Path)
    Next
    
End Sub

Sub Rename(DestinationFolder As String, fso As Object, NewMonth As String)
        
    Dim OldName As String
    Dim NewName As String
    
    '获得DestinationFolder选项的文件对象
    Dim Folder As Object
    Set Folder = fso.GetFolder(DestinationFolder)
    
    '使用了 VBScript.RegExp 类来创建一个正则表达式对象
    Dim RegExp As Object
    Set RegExp = CreateObject("VBScript.RegExp")
    
    '设置正则表达式的模式为匹配字符串中的"数字无限次"
    RegExp.Pattern = "\d+"
    
    '利用正则表达式替换OldName，同时不用再构建一个新的数组存放NewName
    For Each File In Folder.Files
        OldName = File.Name
        NewName = RegExp.Replace(OldName, NewMonth)
        File.Name = NewName
    Next
    
End Sub

