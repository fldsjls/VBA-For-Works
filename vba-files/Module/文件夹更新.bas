Attribute VB_Name = "�ļ��и���"
Sub Copy()
    '�����Լ�����ļ�ϵͳ���ʵĶ���
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ����SourceFolder��·�����ڴ�������
    Dim SourceFolder As String
    SourceFolder = InputBox("������Դ�ļ������ļ��е�·����")
    
    '���û�����룬��ֹͣ����
    If SourceFolder = "" Then
        Exit Sub
    End If
    
    ' �����µ�DestinationFolderName���ڴ�������
    Dim DestinationFolderName As String
    DestinationFolderName = InputBox("������Ŀ���ļ��е��ļ���")
    
    '���û�����룬��ֹͣ����
    If DestinationFolderName = "" Then
        Exit Sub
    End If
    
    ' �����µ�NewMonth����DestinationFolderName���·������ֵ����
    Dim NewMonth As String
    NewMonth = Right(DestinationFolderName, 2)
    
    ' ��SourceFolder�ļ��е���һ��·������ParentDestinationFolder
    Dim ParentDestinationFolder As String
    ParentDestinationFolder = Left(SourceFolder, InStrRev(SourceFolder, "\"))
    
    '�����µ�DestinationFolder�����丳ֵ
    Dim DestinationFolder As String
    DestinationFolder = ParentDestinationFolder & DestinationFolderName
    
    'ʹ��Fso.CopyFolder����SourceFolder�ļ��и��Ƶ�DestinationFolder
    fso.CopyFolder SourceFolder, DestinationFolder
    MsgBox "�ɹ�����", vbInformation
    
    '����Delete��������ɾ������������ΪDestinationFolder
    Call Delete(DestinationFolder)
    MsgBox "�ɹ�ɾ��", vbInformation
        
    '����Rename���������ļ�������
    Call Rename(DestinationFolder, fso, NewMonth)
    MsgBox "�ɹ�����", vbInformation
    
    '�˳�����Wd����
    On Error Resume Next
    For Each WdApp In GetObject(, "Word.Application")
        WdApp.Quit
    Next
    On Error GoTo 0
    
End Sub

Sub Delete(DestinationFolder As String)

    '�����Լ�����ļ�ϵͳ���ʵĶ���
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '����Ŀ���ļ��е����ļ��к��ļ���ɾ������
    Dim Folder As Object
    Set Folder = fso.GetFolder(DestinationFolder)
    For Each Subfolder In Folder.SubFolders
        For Each File In Subfolder.Files
            fso.DeleteFile File.Path, True
        Next
        'ͨ����SubFolder�ļ���·����ֵ��Delete���̵ı������еݹ�
        Call Delete(Subfolder.Path)
    Next
    
End Sub

Sub Rename(DestinationFolder As String, fso As Object, NewMonth As String)
        
    Dim OldName As String
    Dim NewName As String
    
    '���DestinationFolderѡ����ļ�����
    Dim Folder As Object
    Set Folder = fso.GetFolder(DestinationFolder)
    
    'ʹ���� VBScript.RegExp ��������һ��������ʽ����
    Dim RegExp As Object
    Set RegExp = CreateObject("VBScript.RegExp")
    
    '����������ʽ��ģʽΪƥ���ַ����е�"�������޴�"
    RegExp.Pattern = "\d+"
    
    '����������ʽ�滻OldName��ͬʱ�����ٹ���һ���µ�������NewName
    For Each File In Folder.Files
        OldName = File.Name
        NewName = RegExp.Replace(OldName, NewMonth)
        File.Name = NewName
    Next
    
End Sub

