# game
Function IsValidPhone(phone As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "^(13[0-9]|14[5-9]|15[0-3,5-9]|16[2567]|17[0-8]|18[0-9]|19[0-3,5-9])\d{8}$"
    IsValidPhone = regEx.Test(phone)
End Function

' 使用：=IsValidPhone(A1)  → 返回TRUE/FALSE





Sub Extract_Money()
    Dim regEx As Object, matches As Object, m As Object
    Dim text As String, result As String
    text = "苹果单价5.8元，运费20元，总计￥125.60"
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "\d+\.?\d*"   ' 匹配数字（含小数）
    regEx.Global = True
    
    Set matches = regEx.Execute(text)
    For Each m In matches
        result = result & m.Value & vbCrLf
    Next
    
    MsgBox "提取金额：" & result
End Sub





' 将 "张三-销售部|李四-财务部" 拆分为二维数组
Sub Split_ComplexText()
    Dim regEx As Object, matches As Object
    Dim arr(), i As Long, text As String
    text = "张三-销售部|李四-财务部"
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "([^-|]+)-([^-|]+)"  ' 捕获姓名和部门
    regEx.Global = True
    
    Set matches = regEx.Execute(text)
    ReDim arr(1 To matches.Count, 1 To 2)
    
    For i = 0 To matches.Count - 1
        arr(i + 1, 1) = matches(i).SubMatches(0) ' 姓名
        arr(i + 1, 2) = matches(i).SubMatches(1) ' 部门
    Next
    
    Range("A1:B" & matches.Count) = arr
End Sub


Function RemoveHTML(text As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "<[^>]+>"   ' 匹配所有<...>标签
    regEx.Global = True
    RemoveHTML = regEx.Replace(text, "")
End Function



Sub DownloadHTMLToRows()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    ' 设置目标网址
    Const URL As String = "https://example.com"  ' 替换为实际网址
    
    ' 创建HTTP请求对象
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")
    
    ' 配置并发送请求
    With httpRequest
        .Open "GET", URL, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36"
        .setRequestHeader "Accept", "text/html"
        .send
    End With
    
    ' 检查HTTP状态码
    If httpRequest.Status <> 200 Then
        MsgBox "HTTP错误 " & httpRequest.Status & ": " & httpRequest.statusText, vbCritical
        Exit Sub
    End If
    
    ' 获取HTML内容并分割为行数组
    Dim htmlContent As String
    htmlContent = httpRequest.responseText
    
    ' 处理不同换行符格式 (Windows/Unix/Mac)
    htmlContent = Replace(htmlContent, vbLf, vbCrLf)   ' 将LF转为CRLF
    htmlContent = Replace(htmlContent, vbCr, vbCrLf)    ' 处理旧版Mac换行符
    Dim htmlLines() As String
    htmlLines = Split(htmlContent, vbCrLf)
    
    ' 准备写入工作表
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("HTML Output")
    ws.Cells.Delete  ' 清除旧数据
    
    ' 设置标题行格式
    With ws.Range("A1:B1")
        .Value = Array("行号", "HTML内容")
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 255)
    End With
    
    ' 计算需要写入的行数（避免超过Excel限制）
    Dim lineCount As Long
    lineCount = UBound(htmlLines) + 1
    If lineCount > 1048575 Then
        lineCount = 1048575  ' Excel最大行数限制
        MsgBox "警告：HTML超过Excel最大行数限制，已截断", vbExclamation
    End If
    
    ' 批量写入数据（提高性能）
    Dim outputData() As Variant
    ReDim outputData(1 To lineCount, 1 To 2)
    
    Dim i As Long
    For i = 1 To lineCount
        outputData(i, 1) = i  ' 行号
        outputData(i, 2) = htmlLines(i - 1)  ' HTML内容
    Next i
    
    ' 写入工作表并调整格式
    With ws.Range("A2").Resize(lineCount, 2)
        .Value = outputData
        .Columns(1).AutoFit  ' 行号列自动宽度
        .Columns(2).ColumnWidth = 100  ' HTML内容列宽度
        .Rows.RowHeight = 15  ' 行高
        .WrapText = False     ' 不自动换行（提高性能）
    End With
    
    ' 添加自动筛选
    ws.Range("A1:B1").AutoFilter
    
    ' 添加状态栏提示
    Application.StatusBar = "成功下载 " & lineCount & " 行HTML代码"
    
    ' 可选：保存原始HTML到文本文件
    ' SaveRawHTML htmlContent
    
    Exit Sub
    
ErrorHandler:
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

' 可选：保存原始HTML到文本文件
Sub SaveRawHTML(content As String)
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\Webpage_" & Format(Now, "yyyymmdd_hhmmss") & ".html"
    
    Open filePath For Output As #1
    Print #1, content
    Close #1
    
    MsgBox "原始HTML已保存至：" & vbCrLf & filePath, vbInformation
End Sub


