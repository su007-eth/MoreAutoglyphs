Attribute VB_Name = "Module1"
Option Explicit

Const ONE As Long = 10939    'seed���76933��(��������32767)
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
'//getScheme���������趨�ĸ��ʷ���1-10��index��Χ��0-82������ֵ��1-10��Ƶ������
'//��20��15��13��11��9��5��4��3��2��1��
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
        '����=82ʱ������10
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
    '//Data URI Prefix. URIǰ׺"data:text/plain;charset=utf-8,"��˵���Ǵ��ı���ʽ���ο���
    '//https://blog.csdn.net/WuLex/article/details/109226587

    myMod = (a Mod 11) + 5
    '//Sparsity. ϡ���ԣ���mod����a����11��������5��modֵ��Χ��[5,15]��ֵԽ��ͼ������Խ�ࡣ

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
    'Symbol Set. ѡ����ţ������0���Ǵ��󲢻ָ���ǰ�ĵ��á�10����Ϸ��ų��ֵĸ��ʷֱ���20��15��
    '13��11��9��5��4��3��2��1�� ÿ��5�ֽڣ�ͷβΪ�հ�2E���м�Ϊֵ��������ں��油�հ�2E

    Dim i As Integer, j As Integer, v As Integer
    For i = 0 To SIZE - 1
        y = 2 * (i - HALF_SIZE) + 1
        If (a Mod 3) = 1 Then
            y = -y
        ElseIf (a Mod 3) = 2 Then
            y = Abs(y)
        End If
        'Vertical Symmetry. ��ֱ�Գ������a������3ʱ�Գƣ���������ʱ���Գơ�
        '��ѭ��64�Σ�ÿѭ�����һ��

        y = y * a
        For j = 0 To SIZE - 1
            x = 2 * (j - HALF_SIZE) + 1
            If (a Mod 2) = 1 Then
                x = Abs(x)
            End If
            'Horizontal Symmetry. ˮƽ�Գ������a������ʱ���ֲ��Գƣ�a��ż��ʱ�Գơ�
            'Сѭ��64�Σ���ĳһ��ÿ����һ�����ţ�ֱ�������һ�С�

            x = x * a
            v = Int(Abs(x / ONE * y)) Mod myMod
            If v < 5 Then
                value = Mid(symbols, v + 1, 1)
            Else
                value = "."
            End If
            'Sysbol Assignment. ���ŷ��䣬vС��5ȡsymbol�������һ��ֵ�����ڵ���5ȡ�հ׷��š�����
            'modֵ��5-15����ͬ��v�ķ�Χ��0-4��0-14��

            output = output + value
            '�����żӵ�output���顣
        Next j
        output = output + endmark
        'Row Termination. ÿ��ĩβд�� %0A �����ַ���Ϊ���з���
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
    Print #fn, LoadResString(101) 'svg�ļ�ͷ����ͬ���ݣ��������С����Ŷ��塢���Ʊ�������������Դ�ļ�101��������д��svg�ļ���
        
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
            Print #fn, "  <use xlink:href=""#O"" x=""" & x & """ y=""" & y & """/>" '����ͷ�ļ�<defs>������IDΪ0��g���ݣ�д��svg�ļ���
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
' *�����鸳Min-Max֮����������
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

