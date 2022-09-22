VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "More Autoglyphs"
   ClientHeight    =   10545
   ClientLeft      =   7710
   ClientTop       =   4470
   ClientWidth     =   13425
   Icon            =   "MoreAutoglyphs.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "MoreAutoglyphs.frx":7E6A
   ScaleHeight     =   10545
   ScaleWidth      =   13425
   Begin VB.VScrollBar VScroll2 
      Height          =   495
      Left            =   12240
      TabIndex        =   16
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10560
      Left            =   0
      ScaleHeight     =   10560
      ScaleWidth      =   10560
      TabIndex        =   15
      Top             =   0
      Width           =   10560
   End
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   10800
      Top             =   9840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Open uri txt file"
      Height          =   1335
      Left            =   10800
      TabIndex        =   11
      Top             =   7080
      Width           =   2295
      Begin VB.CommandButton CommandOpenFile 
         Caption         =   "Open"
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Batch Creat"
      Height          =   2535
      Left            =   10800
      TabIndex        =   6
      Top             =   4320
      Width           =   2295
      Begin VB.CheckBox CheckClean 
         Caption         =   "Clean Folder"
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton CommandBatch 
         Caption         =   "Creat"
         Height          =   495
         Left            =   480
         TabIndex        =   8
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox TextTotalSupply 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Text            =   "512"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "TotalSupply:"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   3855
      Left            =   10800
      TabIndex        =   5
      Top             =   240
      Width           =   2295
      Begin VB.CommandButton CommandShuffle 
         Caption         =   "Shuffle"
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000011&
         Height          =   1335
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "MoreAutoglyphs.frx":82F4
         Top             =   2280
         Width           =   1815
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   495
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton CommandSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox TextID 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Text            =   "10"
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   11880
      TabIndex        =   4
      Top             =   9840
      Width           =   1215
   End
   Begin VB.Label LabelInfo 
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   10800
      TabIndex        =   13
      Top             =   8520
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim preSeed() As Long

Private Sub Form_Load()
    Picture1.AutoRedraw = True
    Picture1.Scale (0, 0)-(880, 880)
    TextID_Change
    ReDim preSeed(0)
End Sub

Private Sub TextID_Change()
    Preview draw(Val(TextID))
    If Val(TextID) >= VScroll1.Min And Val(TextID) <= VScroll1.Max Then VScroll1.value = Val(TextID)
End Sub

Public Sub Preview(ByVal uri As String)
    Dim i As Long, x As Long, y As Long
    Dim tempS As String * 1
    Const x0 As Long = 120
    Const y0 As Long = 120
    Picture1.Cls
    Picture1.ForeColor = vbBlack
    Picture1.DrawWidth = 1
    On Error Resume Next
    x = x0
    y = y0
    For i = 31 To Len(uri)
        tempS = Mid(uri, i, 1)
        Select Case tempS
        Case "%"  '%0A
            If Mid(uri, i, 3) = "%0A" Then
                y = y + 10
                x = x0
            End If
        Case "."
            x = x + 10
        Case "O"
             Picture1.Circle (x + 5, y + 5), 5, vbBlack
            x = x + 10
        Case "+"
             Picture1.Line (x, y + 5)-(x + 10, y + 5), vbBlack
             Picture1.Line (x + 5, y)-(x + 5, y + 10), vbBlack
            x = x + 10
        Case "X"
             Picture1.Line (x, y)-(x + 10, y + 10), vbBlack
             Picture1.Line (x, y + 10)-(x + 10, y), vbBlack
            x = x + 10
        Case "|"
             Picture1.Line (x + 5, y)-(x + 5, y + 10), vbBlack
            x = x + 10
        Case "-"
             Picture1.Line (x, y + 5)-(x + 10, y + 5), vbBlack
            x = x + 10
        Case "\"
             Picture1.Line (x, y)-(x + 10, y + 10), vbBlack
            x = x + 10
        Case "/"
             Picture1.Line (x, y + 10)-(x + 10, y), vbBlack
            x = x + 10
        Case "#"
             Picture1.Line (x, y)-Step(10, 10), , BF
            x = x + 10
       Case Else
        End Select
    Next i
'        Picture1.ForeColor = RGB(91, 194, 231)
'        Picture1.DrawWidth = 5
'    Picture1.Line (440 - 100, 440 - 50)-Step(200, 100), , BF
End Sub

Private Sub VScroll1_Change()
    TextID = VScroll1.value
End Sub

Private Sub CommandShuffle_Click()
    Randomize
    TextID = Int(Rnd() * 76933) + 1
    ReDim Preserve preSeed(UBound(preSeed) + 1)
    preSeed(UBound(preSeed)) = TextID
    VScroll2.Max = UBound(preSeed)
    VScroll2.value = UBound(preSeed)
    If Val(TextID) >= VScroll1.Min And Val(TextID) <= VScroll1.Max Then VScroll1.value = Val(TextID)
End Sub

Private Sub CommandSave_Click()
    Dim uri As String
    LabelInfo.Caption = ""
    uri = draw(Val(TextID))
    SaveSvg uri, TextID
    LabelInfo.Caption = "Done. build\glyph\" & TextID & ".svg"
End Sub

Private Sub CommandBatch_Click()
    If Val(TextTotalSupply) > 76000 Or Val(TextTotalSupply) < 1 Then
        MsgBox "Out of range, 1-76000.", vbCritical, ""
        Exit Sub
    End If
    Dim cleanFolder As Boolean
    LabelInfo.Caption = "Creating........"
    DoEvents
    If CheckClean.value = 1 Then cleanFolder = True Else cleanFolder = False
    batchCreateGlyph Val(TextTotalSupply.Text), cleanFolder
    LabelInfo.Caption = "Done. build\glyph\"
End Sub

Private Sub CommandOpenFile_Click()
    Dim uri As String, uriPreview As String, fn As Integer
    Dim shortName As String
    Dim tempS As String
    LabelInfo.Caption = ""
    With cmnDlg
        .DialogTitle = "Select uri txt File"
        .Flags = cdlOFNExplorer Or cdlOFNFileMustExist
        .Filter = "txt(.txt)|*.txt|" & "All Files|*.*"
    End With
    cmnDlg.ShowOpen
    If cmnDlg.FileTitle = "" Or LCase(Right(cmnDlg.FileTitle, 3)) <> "txt" Then
        Exit Sub
    End If
    
    uri = ""
    shortName = Left(cmnDlg.FileTitle, Len(cmnDlg.FileTitle) - 4)
    fn = FreeFile
    Open cmnDlg.FileName For Input As #fn
    Do
        Input #fn, tempS
        uri = uri & tempS & vbCrLf
    Loop Until EOF(fn)
    Close #fn
    Preview uri
    SaveSvg uri, shortName
    LabelInfo.Caption = "Done. \build\glyph\" & shortName & ".svg"
End Sub

Private Sub CommandExit_Click()
    End
End Sub

Private Sub TextID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
    TextID = TextID - 1
    End If
    If KeyCode = vbKeyDown Then
        TextID = TextID + 1
    End If
End Sub

Private Sub CommandShuffle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
    TextID = TextID - 1
    End If
    If KeyCode = vbKeyDown Then
        TextID = TextID + 1
    End If
End Sub

Private Sub VScroll2_Change()
    On Error Resume Next
    VScroll2.Min = 1
    If UBound(preSeed) > 0 Then VScroll2.Max = UBound(preSeed) Else VScroll2.Max = 1
    TextID = preSeed(VScroll2.value)
    TextID_Change
End Sub

