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
