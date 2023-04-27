Attribute VB_Name = "文件内更新"
Sub CallOption()
  
    '定义新的DestinationFile并对其赋值
    Dim DestinationFile As String
    DestinationFile = InputBox("请输入目标文件的路径")
    
    '定义WbFile路径，并定义打开的Wb为工作簿对象
    Dim WbFile As String
    WbFile = DestinationFile
    Dim Wb As Workbook
    Set Wb = Workbooks.Open(WbFile)
    
    '如果没有输入，则停止过程
    If DestinationFile = "" Then
        Exit Sub
    End If
    
    '调用UpDate函数进行数据更新
    Call Update日期(Wb)
    Call Update项目(Wb)
    Call Update劳务(Wb)
    
    Wb.Close Savechanges:=True
    MsgBox "更新成功！", vbInformation
    
End Sub


Sub Update日期(Wb As Workbook)

    '设置粘贴源PasteSource
    Dim PasteSource As Excel.Range
    Set PasteSource = Wb.Worksheets("工资表").Range("L13")
    
    '设置粘贴对象PasteTarget
    Dim PasteTarget(1 To 3) As Excel.Range
    Set PasteTarget(1) = Wb.Worksheets("班组结算汇总表").Range("G2：J2")
    Set PasteTarget(2) = Wb.Worksheets("人工费和税管费").Range("G2：I2,G28:I28")
    Set PasteTarget(3) = Wb.Worksheets("挂账和支付").Range("D2,D19,D36,D53")
        
    '利用For循环为每PasteTarget数组进行粘贴
    Dim i As Integer
    For i = 1 To 3
        PasteTarget(i).Value = PasteSource.Value
        i = i + 1
    Next
    
    '清空剪切板上的内容防止弹窗，关闭Excel文件
    Application.CutCopyMode = False
    '
End Sub

Sub Update项目(Wb As Workbook)

    '设置粘贴源PasteSource
    Dim PasteSource As Excel.Range
    Set PasteSource = Wb.Worksheets("工资表").Range("L12")
    
    '设置粘贴对象PasteTarget
    Dim PasteTarget(1 To 3) As Excel.Range
    Set PasteTarget(1) = Wb.Worksheets("班组结算汇总表").Range("C3：G3")
    Set PasteTarget(2) = Wb.Worksheets("人工费和税管费").Range("B3：G3,B29:G29")
    Set PasteTarget(3) = Wb.Worksheets("挂账和支付").Range("A3:E3,A20:E20,A37:E37,A54:E54")
        
    '利用For循环为每PasteTarget数组进行粘贴
    Dim i As Integer
    For i = 1 To 3
        PasteTarget(i).Value = PasteSource.Value
        i = i + 1
    Next
    
    '清空剪切板上的内容防止弹窗，关闭Excel文件
    Application.CutCopyMode = False

End Sub

Sub Update劳务(Wb As Workbook)

    '设置粘贴源PasteSource
    Dim PasteSource As Excel.Range
    Set PasteSource = Wb.Worksheets("工资表").Range("L15")
    
    '设置粘贴对象PasteTarget
    Dim PasteTarget(1 To 3) As Excel.Range
    Set PasteTarget(1) = Wb.Worksheets("班组结算汇总表").Range("I3：M3")
    Set PasteTarget(2) = Wb.Worksheets("人工费和税管费").Range("B4:G4,B30:G30")
    Set PasteTarget(3) = Wb.Worksheets("挂账和支付").Range("B4:F4,B21:F21,B38:F38,B55:F55")
        
    '利用For循环为每PasteTarget数组进行粘贴
    Dim i As Integer
    For i = 1 To 3
        PasteTarget(i).Value = PasteSource.Value
        i = i + 1
    Next
    
    '清空剪切板上的内容防止弹窗，关闭Excel文件
    Application.CutCopyMode = False

End Sub
