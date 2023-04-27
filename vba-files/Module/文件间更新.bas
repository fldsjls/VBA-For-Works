Attribute VB_Name = "文件间更新"
Sub Copy()

    '定义新的DestinationFolder并对其赋值
    Dim DestinationFolder As String
    DestinationFolder = InputBox("请输入目标文件夹的路径") & "\"
    
    '如果没有输入，则停止过程
    If DestinationFolder = "\" Then
        Exit Sub
    End If
    
    '调用UpDate函数进行数据更新
    Call ExcelUpdate工资表(DestinationFolder)
    Call ExcelUpdate出入库(DestinationFolder)
    Call WordUpdate委托书(DestinationFolder)
    Call WordUpdate收款确认书(DestinationFolder)
    MsgBox "更新成功！", vbInformation
    
End Sub

Sub ExcelUpdate工资表(DestinationFolder As String)

    '利用Dir和&得到WbSourceFile文件路径
    Dim WbSourceFolder As String
    Dim WbSourceFile As String
    WbSourceFolder = DestinationFolder & "工资资料\"
    WbSourceFile = WbSourceFolder & Dir(WbSourceFolder & "*工资表*纯净版*")
    
    '利用Dir和&得到WbTargetFile文件路径
    Dim WbTargetFile As String
    WbTargetFile = DestinationFolder & Dir(DestinationFolder & "*计算总表*")
    
    '定义源工作表和目标工作表的工作簿和单元格变量，并将值赋给它们
    Dim WbTarget As Workbook
    Dim WbSource As Workbook
    Dim WsSource As Worksheet
    Dim WsTarget As Worksheet
    Dim RngSource As Range
    Dim RngTarget As Range
    
    Set WbTarget = Workbooks.Open(WbTargetFile, 3, False)
    Set WsTarget = WbTarget.Worksheets("工资表")
    Set WbSource = Workbooks.Open(WbSourceFile, 3, False)
    Set WsSource = WbSource.Worksheets("工资表")
    
    '设置目标数据区域为目标工作表的相同大小区域，以A1单元格为起点
    Set RngSource = WsSource.Range("A1:I20")
    Set RngTarget = WsTarget.Range("A1:I20")
    
    '复制源数据区域到目标数据区域，并使用PasteSpecial xlPasteFormulas进行链接
    RngSource.Copy
    RngTarget.PasteSpecial xlPasteFormulas
    
    '清空剪切板上的内容防止弹窗，同时关闭活动工作簿
    Application.CutCopyMode = False
    WbTarget.Close Savechanges:=True
    WbSource.Close Savechanges:=True
    
End Sub

Sub ExcelUpdate出入库(DestinationFolder As String)

    '利用Dir和&得到WbSourceFile文件路径
    Dim WbSourceFolder As String
    Dim WbSourceFile As String
    WbSourceFolder = DestinationFolder & "统计资料\"
    WbSourceFile = WbSourceFolder & Dir(WbSourceFolder & "*出入库*")

    '利用Dir和&得到WbTargetFile文件路径
    Dim WbTargetFile As String
    WbTargetFile = DestinationFolder & Dir(DestinationFolder & "*计算总表*")
    
    Dim WbTarget As Workbook
    Dim WbSource As Workbook
    Dim WsSource As Worksheet
    Dim WsTarget As Worksheet
    Dim RngSource As Range
    Dim RngTarget As Range
    
    Set WbTarget = Workbooks.Open(WbTargetFile, 3, False)
    Set WsTarget = WbTarget.Worksheets("班组结算汇总表")
    Set WbSource = Workbooks.Open(WbSourceFile, 3, False)
    Set WsSource = WbSource.Worksheets("班组结算汇总表")
    
    '设置目标数据区域为目标工作表的相同大小区域，以A1单元格为起点
    Set RngSource = WsSource.Range("A5:J7")
    Set RngTarget = WsTarget.Range("A5:J7")
    
    '复制源数据区域到目标数据区域，并使用PasteSpecial xlPasteFormulas进行链接
    RngSource.Copy
    RngTarget.PasteSpecial xlPasteFormulas
    
    '清空剪切板上的内容防止弹窗，同时关闭活动工作簿
    Application.CutCopyMode = False
    WbTarget.Close Savechanges:=True
    WbSource.Close Savechanges:=True
    
End Sub

Sub WordUpdate委托书(DestinationFolder As String)

    '创建WdApp应用程序对象
    Dim WdApp As Word.Application
    Set WdApp = CreateObject("Word.Application")
    
    '利用Dir和&得到WdTemplateFile文件路径，并定义一个文件模板变量
    Dim WdTemplateFile As String
    WdTemplateFile = "E:\1Data Management\2Work Material\3公司工作\3工作记录\自动化模板\委托书模板.docx"
    
    '用Open的方法打开WdTemplateFile文件返回WdDoc对象
    Dim WdDoc As Word.Document
    Set WdDoc = Documents.Open(WdTemplateFile)
    
    '创建XlApp应用程序对象
    Dim XlApp As Application
    Set XlApp = CreateObject("Excel.Application")

    '利用Dir和&得到ExcelFile文件路径,设定Excel可见性未Flase
    Dim ExcelFile As String
    ExcelFile = DestinationFolder & Dir(DestinationFolder & "*计算总表*")

    '用Open的方法打开ExcelFile文件返回ExcelWb对象
    Dim ExcelWb As Workbook
    Set ExcelWb = Workbooks.Open(ExcelFile, 3)
    
    '将单元格的内容赋值给LinkAddress数组
    Dim LinkAddress(1 To 6) As Excel.Range
    
    Set LinkAddress(1) = ExcelWb.Sheets("工资表").Range("L12")
    Set LinkAddress(2) = ExcelWb.Sheets("工资表").Range("L14")
    Set LinkAddress(3) = ExcelWb.Sheets("工资表").Range("B20")
    Set LinkAddress(4) = ExcelWb.Sheets("工资表").Range("I20")
    Set LinkAddress(5) = ExcelWb.Sheets("工资表").Range("L15")
    Set LinkAddress(6) = ExcelWb.Sheets("工资表").Range("L13")
    
    '用Find查询移动光标，用Selection选择要粘贴的内容，用While...Wend语句循环Find，实现全部查询的功能
    Dim LinkText As String
    Dim i As Integer
    Dim WdRanges(1 To 6) As Word.Range
    
    For i = 1 To 6
        Set WdRanges(i) = WdDoc.Content
        LinkText = "[Link" & i & "]"
        LinkAddress(i).Copy
        
        '用Find确定锚的Range，并进行PasteSpecial，粘贴为一个域代码(链接)
        With WdRanges(i).Find
            .ClearFormatting
            .Text = LinkText
            .ClearFormatting
            While .Execute '用While...wend语句穷尽文档中的Find查找值
                WdRanges(i).PasteSpecial link:=True, Placement:=wdInLine, DisplayAsIcon:=False, DataType:=wdPasteText
            Wend
        End With
    Next
    
    '使用For Each循环为WdDoc中的每个Field执行替换程序，使域代码符合要求
    For Each Field In WdDoc.Fields
        With Field.Code.Find
            .ClearFormatting
            .Execute FindText:="\t", ReplaceWith:="\t \f2"
        End With
    Next
    
    '定义模板文件的路径
    Dim WdFile As String
    WdFile = DestinationFolder & Dir(DestinationFolder & "*委托书*")
    
    '清空剪切板上的内容防止弹窗，关闭Excel文件，另存为Word文件
    Application.CutCopyMode = False
    ExcelWb.Close Savechanges:=True
    WdDoc.SaveAs2 Filename:=WdFile, FileFormat:=wdFormatXMLDocument
    WdDoc.Close
    
    '释放Wd程序变量的内存
    Set WdDoc = Nothing
    Set WdApp = Nothing

End Sub

Sub WordUpdate收款确认书(DestinationFolder As String)
    
    '创建WdApp应用程序对象
    Dim WdApp As Word.Application
    Set WdApp = CreateObject("Word.Application")
    
    '利用Dir和&得到WdTemplateFile文件路径，并定义一个文件模板变量
    Dim WdTemplateFile As String
    WdTemplateFile = "E:\1Data Management\2Work Material\3公司工作\3工作记录\自动化模板\收款确认书模板.docx"
    
    '用Open的方法打开WdTemplateFile文件返回WdDoc对象
    Dim WdDoc As Word.Document
    Set WdDoc = WdApp.Documents.Open(WdTemplateFile)
    
    '创建XlApp应用程序对象
    Dim XlApp As Application
    Set XlApp = CreateObject("Excel.Application")

    '利用Dir和&得到ExcelFile文件路径,设定Excel可见性未Flase
    Dim ExcelFile As String
    ExcelFile = DestinationFolder & Dir(DestinationFolder & "*计算总表*")

    '用Open的方法打开ExcelFile文件返回ExcelWb对象
    Dim ExcelWb As Workbook
    Set ExcelWb = XlApp.Workbooks.Open(ExcelFile, 3)
    
    '将单元格的内容赋值给LinkAddress数组
    Dim LinkAddress(1 To 6) As Excel.Range
    
    Set LinkAddress(1) = ExcelWb.Sheets("工资表").Range("L12")
    Set LinkAddress(2) = ExcelWb.Sheets("工资表").Range("L13")
    Set LinkAddress(3) = ExcelWb.Sheets("挂账和支付").Range("J5")
    Set LinkAddress(4) = ExcelWb.Sheets("挂账和支付").Range("K5")
    Set LinkAddress(5) = ExcelWb.Sheets("挂账和支付").Range("J6")
    Set LinkAddress(6) = ExcelWb.Sheets("挂账和支付").Range("K6")
    
    '用Copy语方法将LinkAddress中的值复制到剪切板中，再用Find对象进行查询和替换
    Dim LinkText As String
    Dim i As Integer
    Dim WdRanges(1 To 6) As Word.Range
    
    For i = 1 To 6
        Set WdRanges(i) = WdDoc.Content
        LinkText = "[Link" & i & "]"
        LinkAddress(i).Copy
        
        '用Find确定锚的Range，并进行PasteSpecial，粘贴为一个域代码(链接)
        With WdRanges(i).Find
            .ClearFormatting
            .Text = LinkText
            .ClearFormatting
            While .Execute '用While...wend语句穷尽文档中的Find查找值
                WdRanges(i).PasteSpecial link:=True, Placement:=wdInLine, DisplayAsIcon:=False, DataType:=wdPasteText
            Wend
        End With
    Next
    
    '使用For Each循环为WdDoc中的每个Field执行替换程序，使域代码符合要求
    For Each Field In WdDoc.Fields
        With Field.Code.Find
            .ClearFormatting
            .Execute FindText:="\t", ReplaceWith:="\t \f2"
        End With
    Next
    
    '定义模板文件的路径
    Dim WdFile As String
    WdFile = DestinationFolder & Dir(DestinationFolder & "*收款确认书*")
    
    '清空剪切板上的内容防止弹窗，关闭Excel文件，另存为Word文件
    Application.CutCopyMode = False
    ExcelWb.Close Savechanges:=True
    WdDoc.SaveAs2 Filename:=WdFile, FileFormat:=wdFormatXMLDocument
    WdDoc.Close
    
    '释放Wd程序变量的内存
    Set WdDoc = Nothing
    Set WdApp = Nothing
    
End Sub


