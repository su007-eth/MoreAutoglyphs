Attribute VB_Name = "Module1"
Option Explicit

Const ONE As Long = 10939    'seed最大76933，(滚动条是32767)
Const SIZE As Long = 64
Const HALF_SIZE = SIZE / 2
Const prefix As String = "data:text/plain;charset=utf-8,"
Const endmark As String = "%0A"
Private numTokens As Long
Private allSeed() As Long

'    // 0x2E = .
'    // 0x4F = O
'    // 0x2B = +
'    // 0x58 = X
'    // 0x7C = |
'    // 0x2D = -
'    // 0x5C = \
'    // 0x2F = /
'    // 0x23 = #
'//getScheme函数按照设定的概率返回1-10。index范围是0-82，返回值是1-10的频率依次
'//是20，15，13，11，9，5，4，3，2，1。
Private Function getScheme(ByVal a As Long) As Integer
    Dim index As Long
    Dim scheme As Integer
    index = a Mod 83
    If index < 20 Then
        scheme = 1
    ElseIf index < 35 Then
        scheme = 2
    ElseIf index < 48 Then
        scheme = 3
    ElseIf index < 59 Then
        scheme = 4
    ElseIf index < 68 Then
        scheme = 5
    ElseIf index < 73 Then
        scheme = 6
    ElseIf index < 77 Then
        scheme = 7
    ElseIf index < 80 Then
        scheme = 8
    ElseIf index < 82 Then
        scheme = 9
    Else
        scheme = 10
        '正好=82时，返回10
    End If
    getScheme = scheme
End Function

Public Function draw(ByVal seed As Long) As String
    If seed > 76933 Or seed < 1 Then
        MsgBox "Out of range, 1-76933.", vbCritical, ""
        Exit Function
    End If

    Dim a As Long
    a = seed
    Dim output As String
    Dim x As Long
    Dim y As Long
    Dim value As String * 1
    Dim myMod As Integer
    Dim symbols As String * 5
    Dim symbolScheme As Integer
    output = prefix
    '//Data URI Prefix. URI前缀"data:text/plain;charset=utf-8,"，说明是纯文本形式。参考：
    '//https://blog.csdn.net/WuLex/article/details/109226587

    myMod = (a Mod 11) + 5
    '//Sparsity. 稀疏性，让mod等于a除以11的余数加5，mod值范围是[5,15]，值越大图中留白越多。

    symbolScheme = getScheme(a)
    If symbolScheme = 0 Then
        Exit Function
    ElseIf symbolScheme = 1 Then
        symbols = ".X/\."
    ElseIf symbolScheme = 2 Then
        symbols = ".+-|."
    ElseIf symbolScheme = 3 Then
        symbols = "./\.."
    ElseIf symbolScheme = 4 Then
        symbols = ".\|-/"
    ElseIf symbolScheme = 5 Then
        symbols = ".O|-."
    ElseIf symbolScheme = 6 Then
        symbols = ".\\.."
    ElseIf symbolScheme = 7 Then
        symbols = ".#|-+"
    ElseIf symbolScheme = 8 Then
        symbols = ".OO.."
    ElseIf symbolScheme = 9 Then
        symbols = ".#..."
    Else
        symbols = ".#O.."
    End If
    'Symbol Set. 选择符号，如果是0则标记错误并恢复当前的调用。10种组合符号出现的概率分别是20，15，
    '13，11，9，5，4，3，2，1。 每组5字节，头尾为空白2E，中间为值，不足的在后面补空白2E

    Dim i As Integer, j As Integer, v As Integer
    For i = 0 To SIZE - 1
        y = 2 * (i - HALF_SIZE) + 1
        If (a Mod 3) = 1 Then
            y = -y
        ElseIf (a Mod 3) = 2 Then
            y = Abs(y)
        End If
        'Vertical Symmetry. 垂直对称情况，a能整除3时对称，不能整除时不对称。
        '大循环64次，每循环完成一行

        y = y * a
        For j = 0 To SIZE - 1
            x = 2 * (j - HALF_SIZE) + 1
            If (a Mod 2) = 1 Then
                x = Abs(x)
            End If
            'Horizontal Symmetry. 水平对称情况，a是奇数时部分不对称，a是偶数时对称。
            '小循环64次，在某一行每次填一个符号，直至完成这一行。

            x = x * a
            v = Int(Abs(x / ONE * y)) Mod myMod
            If v < 5 Then
                value = Mid(symbols, v + 1, 1)
            Else
                value = "."
            End If
            'Sysbol Assignment. 符号分配，v小于5取symbol数组里的一个值，大于等于5取空白符号。根据
            'mod值（5-15）不同，v的范围从0-4到0-14。

            output = output + value
            '将符号加到output数组。
        Next j
        output = output + endmark
        'Row Termination. 每行末尾写入 %0A 三个字符作为分行符。
    Next i
    draw = output
End Function

Public Sub batchCreateGlyph(ByVal TotalSupply As Long, Optional ByVal cleanFolder As Boolean = False)
    Dim uri As String, fn As Integer
    Dim id As Long, a As Long, addNumID As Long
    Dim idToSeed As New Collection
    
    Dim buildPath As String, glyphPath As String, stringPath As String
    If Right(App.Path, 1) = "\" Then buildPath = App.Path & "build" Else buildPath = App.Path & "\build"
    If Dir(buildPath, vbDirectory) = "" Then MkDir buildPath
    glyphPath = buildPath & "\glyph"
    If Dir(glyphPath, vbDirectory) = "" Then MkDir glyphPath
    stringPath = buildPath & "\string"
    If Dir(stringPath, vbDirectory) = "" Then MkDir stringPath
    If cleanFolder = True Then
        If Dir(glyphPath & "\*.*") <> "" Then Kill glyphPath & "\*.*"
        If Dir(stringPath & "\*.*") <> "" Then Kill stringPath & "\*.*"
    End If
    
    allSeed = shuffle(2, 76933)
    addNumID = 0
    For id = 1 To TotalSupply
        a = allSeed(id + addNumID)
        Do While True
            If a = 32817 Or a = 76573 Then
                addNumID = addNumID + 1
                a = allSeed(id + addNumID)
            Else
                Exit Do
            End If
        Loop

        If id + addNumID > UBound(allSeed) Then
            MsgBox "Exceeds the maximum value, the total output is " & id & " tokens", vbInformation, ""
            Exit Sub
        End If
        uri = draw(a)
        SaveSvg uri, id
        idToSeed.Add a, Str(id)
        DoEvents
        Form1.LabelInfo.Caption = "Creating.. " & id & "/" & TotalSupply
        Form1.Preview uri
    Next id
    
    fn = FreeFile
    Open buildPath & "\" & "idToSeed.txt" For Output As #fn
    Dim i As Long
    For i = 1 To idToSeed.Count
        Print #fn, i & " -> " & idToSeed(i)
    Next i
    Close #fn
    Shell "explorer " & buildPath, 1
End Sub

Public Sub SaveSvg(ByVal uri As String, ByVal FileName As String)
    Dim buildPath As String, glyphPath As String, stringPath As String, fn As Integer
    Dim glyphFileName As String, stringFileName As String
    Dim i As Long, x As Long, y As Long
    Dim tempS As String
       
    Const x0 As Long = 120
    Const y0 As Long = 120
    Const cellWidth As Long = 10
    
    On Error Resume Next
    If Right(App.Path, 1) = "\" Then buildPath = App.Path & "build" Else buildPath = App.Path & "\build"
    If Dir(buildPath, vbDirectory) = "" Then MkDir buildPath
    glyphPath = buildPath & "\glyph"
    If Dir(glyphPath, vbDirectory) = "" Then MkDir glyphPath
    stringPath = buildPath & "\string"
    If Dir(stringPath, vbDirectory) = "" Then MkDir stringPath
    
    fn = FreeFile
    stringFileName = stringPath & "\" & FileName & ".txt"
    Open stringFileName For Output As #fn
    Print #fn, uri
    Close

    fn = FreeFile
    glyphFileName = glyphPath & "\" & FileName & ".svg"
    Open glyphFileName For Output As #fn
    Print #fn, LoadResString(101) 'svg文件头部相同内容（包括首行、符号定义、绘制背景画布）在资源文件101，调出来写入svg文件。
        
    x = x0
    y = y0
    For i = 31 To Len(uri)
        tempS = Mid(uri, i, 1)
        Select Case tempS
        Case "%"  '%0A
            If Mid(uri, i, 3) = "%0A" Then
                y = y + cellWidth
                x = x0
            End If
        Case "."
            x = x + cellWidth
        Case "O"
            Print #fn, "  <use xlink:href=""#O"" x=""" & x & """ y=""" & y & """/>" '调用头文件<defs>区域中ID为0的g内容，写入svg文件。
            x = x + cellWidth
        Case "+"
            Print #fn, "  <use xlink:href=""#+"" x=""" & x & """ y=""" & y & """/>";
            x = x + cellWidth
        Case "X"
            Print #fn, "  <use xlink:href=""#X"" x=""" & x & """ y=""" & y & """/>"
            x = x + cellWidth
        Case "|"
            Print #fn, "  <use xlink:href=""#|"" x=""" & x & """ y=""" & y & """/>"
            x = x + cellWidth
        Case "-"
            Print #fn, "  <use xlink:href=""#-"" x=""" & x & """ y=""" & y & """/>"
            x = x + cellWidth
        Case "\"
            Print #fn, "  <use xlink:href=""#\"" x=""" & x & """ y=""" & y & """/>"
            x = x + cellWidth
        Case "/"
            Print #fn, "  <use xlink:href=""#/"" x=""" & x & """ y=""" & y & """/>"
            x = x + cellWidth
        Case "#"
            Print #fn, "  <use xlink:href=""##"" x=""" & x & """ y=""" & y & """/>"
            x = x + cellWidth
'       Case Else
        End Select
    Next i
    Print #fn, "</svg>"
    Close
End Sub

'/**
' *给数组赋Min-Max之间的随机整数
'*/
Private Function shuffle(ByVal Min As Long, ByVal Max As Long) As Long()
    Dim i As Long, j As Long, tmp As Long
    Dim x() As Long
    ReDim x(Max - Min)

    For i = 0 To Max - Min
        x(i) = Min + i
    Next
    Randomize
    For i = Max To Min Step -1
        j = Int(Rnd * (i - Min)) + Min
        tmp = x(j - Min)
        x(j - Min) = x(i - Min)
        x(i - Min) = tmp
    Next
    shuffle = x
End Function

