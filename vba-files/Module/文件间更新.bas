Attribute VB_Name = "�ļ������"
Sub Copy()

    '�����µ�DestinationFolder�����丳ֵ
    Dim DestinationFolder As String
    DestinationFolder = InputBox("������Ŀ���ļ��е�·��") & "\"
    
    '���û�����룬��ֹͣ����
    If DestinationFolder = "\" Then
        Exit Sub
    End If
    
    '����UpDate�����������ݸ���
    Call ExcelUpdate���ʱ�(DestinationFolder)
    Call ExcelUpdate�����(DestinationFolder)
    Call WordUpdateί����(DestinationFolder)
    Call WordUpdate�տ�ȷ����(DestinationFolder)
    MsgBox "���³ɹ���", vbInformation
    
End Sub

Sub ExcelUpdate���ʱ�(DestinationFolder As String)

    '����Dir��&�õ�WbSourceFile�ļ�·��
    Dim WbSourceFolder As String
    Dim WbSourceFile As String
    WbSourceFolder = DestinationFolder & "��������\"
    WbSourceFile = WbSourceFolder & Dir(WbSourceFolder & "*���ʱ�*������*")
    
    '����Dir��&�õ�WbTargetFile�ļ�·��
    Dim WbTargetFile As String
    WbTargetFile = DestinationFolder & Dir(DestinationFolder & "*�����ܱ�*")
    
    '����Դ�������Ŀ�깤����Ĺ������͵�Ԫ�����������ֵ��������
    Dim WbTarget As Workbook
    Dim WbSource As Workbook
    Dim WsSource As Worksheet
    Dim WsTarget As Worksheet
    Dim RngSource As Range
    Dim RngTarget As Range
    
    Set WbTarget = Workbooks.Open(WbTargetFile, 3, False)
    Set WsTarget = WbTarget.Worksheets("���ʱ�")
    Set WbSource = Workbooks.Open(WbSourceFile, 3, False)
    Set WsSource = WbSource.Worksheets("���ʱ�")
    
    '����Ŀ����������ΪĿ�깤�������ͬ��С������A1��Ԫ��Ϊ���
    Set RngSource = WsSource.Range("A1:I20")
    Set RngTarget = WsTarget.Range("A1:I20")
    
    '����Դ��������Ŀ���������򣬲�ʹ��PasteSpecial xlPasteFormulas��������
    RngSource.Copy
    RngTarget.PasteSpecial xlPasteFormulas
    
    '��ռ��а��ϵ����ݷ�ֹ������ͬʱ�رջ������
    Application.CutCopyMode = False
    WbTarget.Close Savechanges:=True
    WbSource.Close Savechanges:=True
    
End Sub

Sub ExcelUpdate�����(DestinationFolder As String)

    '����Dir��&�õ�WbSourceFile�ļ�·��
    Dim WbSourceFolder As String
    Dim WbSourceFile As String
    WbSourceFolder = DestinationFolder & "ͳ������\"
    WbSourceFile = WbSourceFolder & Dir(WbSourceFolder & "*�����*")

    '����Dir��&�õ�WbTargetFile�ļ�·��
    Dim WbTargetFile As String
    WbTargetFile = DestinationFolder & Dir(DestinationFolder & "*�����ܱ�*")
    
    Dim WbTarget As Workbook
    Dim WbSource As Workbook
    Dim WsSource As Worksheet
    Dim WsTarget As Worksheet
    Dim RngSource As Range
    Dim RngTarget As Range
    
    Set WbTarget = Workbooks.Open(WbTargetFile, 3, False)
    Set WsTarget = WbTarget.Worksheets("���������ܱ�")
    Set WbSource = Workbooks.Open(WbSourceFile, 3, False)
    Set WsSource = WbSource.Worksheets("���������ܱ�")
    
    '����Ŀ����������ΪĿ�깤�������ͬ��С������A1��Ԫ��Ϊ���
    Set RngSource = WsSource.Range("A5:J7")
    Set RngTarget = WsTarget.Range("A5:J7")
    
    '����Դ��������Ŀ���������򣬲�ʹ��PasteSpecial xlPasteFormulas��������
    RngSource.Copy
    RngTarget.PasteSpecial xlPasteFormulas
    
    '��ռ��а��ϵ����ݷ�ֹ������ͬʱ�رջ������
    Application.CutCopyMode = False
    WbTarget.Close Savechanges:=True
    WbSource.Close Savechanges:=True
    
End Sub

Sub WordUpdateί����(DestinationFolder As String)

    '����WdAppӦ�ó������
    Dim WdApp As Word.Application
    Set WdApp = CreateObject("Word.Application")
    
    '����Dir��&�õ�WdTemplateFile�ļ�·����������һ���ļ�ģ�����
    Dim WdTemplateFile As String
    WdTemplateFile = "E:\1Data Management\2Work Material\3��˾����\3������¼\�Զ���ģ��\ί����ģ��.docx"
    
    '��Open�ķ�����WdTemplateFile�ļ�����WdDoc����
    Dim WdDoc As Word.Document
    Set WdDoc = Documents.Open(WdTemplateFile)
    
    '����XlAppӦ�ó������
    Dim XlApp As Application
    Set XlApp = CreateObject("Excel.Application")

    '����Dir��&�õ�ExcelFile�ļ�·��,�趨Excel�ɼ���δFlase
    Dim ExcelFile As String
    ExcelFile = DestinationFolder & Dir(DestinationFolder & "*�����ܱ�*")

    '��Open�ķ�����ExcelFile�ļ�����ExcelWb����
    Dim ExcelWb As Workbook
    Set ExcelWb = Workbooks.Open(ExcelFile, 3)
    
    '����Ԫ������ݸ�ֵ��LinkAddress����
    Dim LinkAddress(1 To 6) As Excel.Range
    
    Set LinkAddress(1) = ExcelWb.Sheets("���ʱ�").Range("L12")
    Set LinkAddress(2) = ExcelWb.Sheets("���ʱ�").Range("L14")
    Set LinkAddress(3) = ExcelWb.Sheets("���ʱ�").Range("B20")
    Set LinkAddress(4) = ExcelWb.Sheets("���ʱ�").Range("I20")
    Set LinkAddress(5) = ExcelWb.Sheets("���ʱ�").Range("L15")
    Set LinkAddress(6) = ExcelWb.Sheets("���ʱ�").Range("L13")
    
    '��Find��ѯ�ƶ���꣬��Selectionѡ��Ҫճ�������ݣ���While...Wend���ѭ��Find��ʵ��ȫ����ѯ�Ĺ���
    Dim LinkText As String
    Dim i As Integer
    Dim WdRanges(1 To 6) As Word.Range
    
    For i = 1 To 6
        Set WdRanges(i) = WdDoc.Content
        LinkText = "[Link" & i & "]"
        LinkAddress(i).Copy
        
        '��Findȷ��ê��Range��������PasteSpecial��ճ��Ϊһ�������(����)
        With WdRanges(i).Find
            .ClearFormatting
            .Text = LinkText
            .ClearFormatting
            While .Execute '��While...wend�����ĵ��е�Find����ֵ
                WdRanges(i).PasteSpecial link:=True, Placement:=wdInLine, DisplayAsIcon:=False, DataType:=wdPasteText
            Wend
        End With
    Next
    
    'ʹ��For Eachѭ��ΪWdDoc�е�ÿ��Fieldִ���滻����ʹ��������Ҫ��
    For Each Field In WdDoc.Fields
        With Field.Code.Find
            .ClearFormatting
            .Execute FindText:="\t", ReplaceWith:="\t \f2"
        End With
    Next
    
    '����ģ���ļ���·��
    Dim WdFile As String
    WdFile = DestinationFolder & Dir(DestinationFolder & "*ί����*")
    
    '��ռ��а��ϵ����ݷ�ֹ�������ر�Excel�ļ������ΪWord�ļ�
    Application.CutCopyMode = False
    ExcelWb.Close Savechanges:=True
    WdDoc.SaveAs2 Filename:=WdFile, FileFormat:=wdFormatXMLDocument
    WdDoc.Close
    
    '�ͷ�Wd����������ڴ�
    Set WdDoc = Nothing
    Set WdApp = Nothing

End Sub

Sub WordUpdate�տ�ȷ����(DestinationFolder As String)
    
    '����WdAppӦ�ó������
    Dim WdApp As Word.Application
    Set WdApp = CreateObject("Word.Application")
    
    '����Dir��&�õ�WdTemplateFile�ļ�·����������һ���ļ�ģ�����
    Dim WdTemplateFile As String
    WdTemplateFile = "E:\1Data Management\2Work Material\3��˾����\3������¼\�Զ���ģ��\�տ�ȷ����ģ��.docx"
    
    '��Open�ķ�����WdTemplateFile�ļ�����WdDoc����
    Dim WdDoc As Word.Document
    Set WdDoc = WdApp.Documents.Open(WdTemplateFile)
    
    '����XlAppӦ�ó������
    Dim XlApp As Application
    Set XlApp = CreateObject("Excel.Application")

    '����Dir��&�õ�ExcelFile�ļ�·��,�趨Excel�ɼ���δFlase
    Dim ExcelFile As String
    ExcelFile = DestinationFolder & Dir(DestinationFolder & "*�����ܱ�*")

    '��Open�ķ�����ExcelFile�ļ�����ExcelWb����
    Dim ExcelWb As Workbook
    Set ExcelWb = XlApp.Workbooks.Open(ExcelFile, 3)
    
    '����Ԫ������ݸ�ֵ��LinkAddress����
    Dim LinkAddress(1 To 6) As Excel.Range
    
    Set LinkAddress(1) = ExcelWb.Sheets("���ʱ�").Range("L12")
    Set LinkAddress(2) = ExcelWb.Sheets("���ʱ�").Range("L13")
    Set LinkAddress(3) = ExcelWb.Sheets("���˺�֧��").Range("J5")
    Set LinkAddress(4) = ExcelWb.Sheets("���˺�֧��").Range("K5")
    Set LinkAddress(5) = ExcelWb.Sheets("���˺�֧��").Range("J6")
    Set LinkAddress(6) = ExcelWb.Sheets("���˺�֧��").Range("K6")
    
    '��Copy�﷽����LinkAddress�е�ֵ���Ƶ����а��У�����Find������в�ѯ���滻
    Dim LinkText As String
    Dim i As Integer
    Dim WdRanges(1 To 6) As Word.Range
    
    For i = 1 To 6
        Set WdRanges(i) = WdDoc.Content
        LinkText = "[Link" & i & "]"
        LinkAddress(i).Copy
        
        '��Findȷ��ê��Range��������PasteSpecial��ճ��Ϊһ�������(����)
        With WdRanges(i).Find
            .ClearFormatting
            .Text = LinkText
            .ClearFormatting
            While .Execute '��While...wend�����ĵ��е�Find����ֵ
                WdRanges(i).PasteSpecial link:=True, Placement:=wdInLine, DisplayAsIcon:=False, DataType:=wdPasteText
            Wend
        End With
    Next
    
    'ʹ��For Eachѭ��ΪWdDoc�е�ÿ��Fieldִ���滻����ʹ��������Ҫ��
    For Each Field In WdDoc.Fields
        With Field.Code.Find
            .ClearFormatting
            .Execute FindText:="\t", ReplaceWith:="\t \f2"
        End With
    Next
    
    '����ģ���ļ���·��
    Dim WdFile As String
    WdFile = DestinationFolder & Dir(DestinationFolder & "*�տ�ȷ����*")
    
    '��ռ��а��ϵ����ݷ�ֹ�������ر�Excel�ļ������ΪWord�ļ�
    Application.CutCopyMode = False
    ExcelWb.Close Savechanges:=True
    WdDoc.SaveAs2 Filename:=WdFile, FileFormat:=wdFormatXMLDocument
    WdDoc.Close
    
    '�ͷ�Wd����������ڴ�
    Set WdDoc = Nothing
    Set WdApp = Nothing
    
End Sub


