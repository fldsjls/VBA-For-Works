Attribute VB_Name = "�ļ��ڸ���"
Sub CallOption()
  
    '�����µ�DestinationFile�����丳ֵ
    Dim DestinationFile As String
    DestinationFile = InputBox("������Ŀ���ļ���·��")
    
    '����WbFile·����������򿪵�WbΪ����������
    Dim WbFile As String
    WbFile = DestinationFile
    Dim Wb As Workbook
    Set Wb = Workbooks.Open(WbFile)
    
    '���û�����룬��ֹͣ����
    If DestinationFile = "" Then
        Exit Sub
    End If
    
    '����UpDate�����������ݸ���
    Call Update����(Wb)
    Call Update��Ŀ(Wb)
    Call Update����(Wb)
    
    Wb.Close Savechanges:=True
    MsgBox "���³ɹ���", vbInformation
    
End Sub


Sub Update����(Wb As Workbook)

    '����ճ��ԴPasteSource
    Dim PasteSource As Excel.Range
    Set PasteSource = Wb.Worksheets("���ʱ�").Range("L13")
    
    '����ճ������PasteTarget
    Dim PasteTarget(1 To 3) As Excel.Range
    Set PasteTarget(1) = Wb.Worksheets("���������ܱ�").Range("G2��J2")
    Set PasteTarget(2) = Wb.Worksheets("�˹��Ѻ�˰�ܷ�").Range("G2��I2,G28:I28")
    Set PasteTarget(3) = Wb.Worksheets("���˺�֧��").Range("D2,D19,D36,D53")
        
    '����Forѭ��ΪÿPasteTarget�������ճ��
    Dim i As Integer
    For i = 1 To 3
        PasteTarget(i).Value = PasteSource.Value
        i = i + 1
    Next
    
    '��ռ��а��ϵ����ݷ�ֹ�������ر�Excel�ļ�
    Application.CutCopyMode = False
    '
End Sub

Sub Update��Ŀ(Wb As Workbook)

    '����ճ��ԴPasteSource
    Dim PasteSource As Excel.Range
    Set PasteSource = Wb.Worksheets("���ʱ�").Range("L12")
    
    '����ճ������PasteTarget
    Dim PasteTarget(1 To 3) As Excel.Range
    Set PasteTarget(1) = Wb.Worksheets("���������ܱ�").Range("C3��G3")
    Set PasteTarget(2) = Wb.Worksheets("�˹��Ѻ�˰�ܷ�").Range("B3��G3,B29:G29")
    Set PasteTarget(3) = Wb.Worksheets("���˺�֧��").Range("A3:E3,A20:E20,A37:E37,A54:E54")
        
    '����Forѭ��ΪÿPasteTarget�������ճ��
    Dim i As Integer
    For i = 1 To 3
        PasteTarget(i).Value = PasteSource.Value
        i = i + 1
    Next
    
    '��ռ��а��ϵ����ݷ�ֹ�������ر�Excel�ļ�
    Application.CutCopyMode = False

End Sub

Sub Update����(Wb As Workbook)

    '����ճ��ԴPasteSource
    Dim PasteSource As Excel.Range
    Set PasteSource = Wb.Worksheets("���ʱ�").Range("L15")
    
    '����ճ������PasteTarget
    Dim PasteTarget(1 To 3) As Excel.Range
    Set PasteTarget(1) = Wb.Worksheets("���������ܱ�").Range("I3��M3")
    Set PasteTarget(2) = Wb.Worksheets("�˹��Ѻ�˰�ܷ�").Range("B4:G4,B30:G30")
    Set PasteTarget(3) = Wb.Worksheets("���˺�֧��").Range("B4:F4,B21:F21,B38:F38,B55:F55")
        
    '����Forѭ��ΪÿPasteTarget�������ճ��
    Dim i As Integer
    For i = 1 To 3
        PasteTarget(i).Value = PasteSource.Value
        i = i + 1
    Next
    
    '��ռ��а��ϵ����ݷ�ֹ�������ر�Excel�ļ�
    Application.CutCopyMode = False

End Sub
