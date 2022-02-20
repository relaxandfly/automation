DIR *.* /B > Library.txt


rename *.* *.rar



Public Sub tt()
s = "用正则1方法A分离B字符"
With CreateObject("VBSCRIPT.REGEXP")
  .Global = True
  .Pattern = "[^一-龥]+"
  MsgBox .Replace(s, " ")
End With
End Sub

Function zldccmx(Rng As Range, Ms As Integer)
    Dim ys(1 To 12): Dim RegEx
    ys(1) = "[^A-Za-z0-9]"    '只保留字母和数字
    ys(2) = "[^!-~]"  '去除中文
    ys(3) = "[!-~]" '"\w"    '留中文
    ys(4) = "\d"    '  去掉数字
    ys(5) = "[^\d]"   ' 留数字
    ys(6) = "\D"    '去除非数字(留数字)
    ys(7) = "[a-zA_Z]"    '去除英文大小写字符
    ys(8) = "3*a*"    '去除所有指定字符，这里指去除3和1
    ys(9) = "36*"    '去除所有指定字符，这里指去除"36"
    ys(10) = "[^3]"    '去除所有非特定字符，这里指去除不是3的字符
    ys(11) = "[^0-9.]"    '只保留数字和小数点
    ys(12) = "[^0-9/.+-^\*^]"    '保留数字和运算符号+-*/^
    Set RegEx = CreateObject("VBSCRIPT.REGEXP")    'RegEx为建立正则表达式
    RegEx.Global = True    '设置全局可用
    RegEx.Pattern = ys(Ms)    '样式
    zldccmx = Replace(RegEx.Replace(Rng, Chr(9)), Chr(9), "")
    Set RegEx = Nothing



'classic
Sub test()
    Dim objRegExp As Object
    Dim i As Long, arr
    arr = Range("a1:b" & Cells(Rows.Count, 1).End(xlUp).Row).Value
    Set objRegExp = CreateObject("VBScript.regExp")
    With objRegExp
        .Global = True
		
		'4e00-9fa5 is chinese regular expression
        'origin .Pattern = "[\u4e00-\u9fa5]{1,}"
		.Pattern = "[a-z_A-Z]{1,}"
        For i = 1 To UBound(arr)
            'If .test(arr(i, 1)) Then
            arr(i, 1) = .Replace(arr(i, 1), "")
            'End If
        Next
    End With
    Set objRegExp = Nothing
    Range("b1").Resize(UBound(arr)) = arr
    MsgBox "删除成功"
End Sub
