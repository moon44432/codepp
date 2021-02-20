VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Untitled - Code++"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   11760
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   11760
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6870
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Bevel           =   0
            Object.Width           =   847
            MinWidth        =   847
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   847
            MinWidth        =   847
            TextSave        =   "INS"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11160
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   9360
      TabIndex        =   3
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   9360
      TabIndex        =   2
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   9360
      TabIndex        =   1
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "새 파일(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnubar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "열기(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "저장(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveOtherName 
         Caption         =   "다른 이름으로 저장(&A)..."
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "종료(&Q)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "편집(&E)"
      Begin VB.Menu mnuCut 
         Caption         =   "잘라내기(&X)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "복사(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "붙여넣기(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "모두 선택(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnubar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "찾기(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "바꾸기(&R)"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "보기(&V)"
      Begin VB.Menu mnuSetFont 
         Caption         =   "글꼴 설정"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "실행(&R)"
      Begin VB.Menu mnuRunApp 
         Caption         =   "실행(&R)..."
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "컴파일(&C)..."
         Enabled         =   0   'False
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnubar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "컴파일 및 실행 설정"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "도움말(&H)"
      Begin VB.Menu mnuHelp 
         Caption         =   "Code++ 도움말"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnubar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "Code++ 정보"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type INITCOMMONCONTROLSEX_TYPE
dwSize As Long
dwICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As INITCOMMONCONTROLSEX_TYPE) As Long
Private Const ICC_INTERNET_CLASSES = &H800

Dim OpenedFile As String
Dim ifFileChanged As Boolean
Dim indent As Integer
Dim ifFileOpened As Boolean


Function SaveFile()
On Error GoTo Err
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "모든 파일|*.*"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.DialogTitle = "파일 저장"
    CommonDialog1.InitDir = "C:\"
    CommonDialog1.FileName = "Untitled.txt"
    CommonDialog1.ShowSave
    
    OpenedFile = CommonDialog1.FileName
    
    Open CommonDialog1.FileName For Output As #1
    Print #1, Text1.Text
    Close #1
    ifFileChanged = False
    frmMain.Caption = OpenedFile & " - Code++"
    ifFileOpened = True
Err:
End Function

Function JustSave()
    Open OpenedFile For Output As #1
    Print #1, Text1.Text
    Close #1
    frmMain.Caption = OpenedFile & " - Code++"
    ifFileChanged = False
End Function

Function LoadFile()
On Error GoTo Err
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "모든 파일|*.*"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.DialogTitle = "파일 열기"
    CommonDialog1.InitDir = "C:\"
    CommonDialog1.ShowOpen
    
    OpenedFile = CommonDialog1.FileName
    
    Text1.Text = ""
    
    Dim str As String
    Open OpenedFile For Input As #1
        Do Until EOF(1)
        Line Input #1, str
        frmMain.Text1 = frmMain.Text1 + str + vbCrLf
        Loop
    Close #1
    
    frmMain.Caption = OpenedFile & " - Code++"
    ifFileChanged = False
    ifFileOpened = True
Err:
End Function


Private Sub Form_Load()
    Dim comctls As INITCOMMONCONTROLSEX_TYPE
    Dim retval As Long
    With comctls
    .dwSize = Len(comctls)
    .dwICC = ICC_INTERNET_CLASSES
    End With
    
    retval = InitCommonControlsEx(comctls)
    
    Dim str As String
    
    Open App.Path & "\config.ini" For Input As #1
    Line Input #1, str
    Label1.Caption = str
    Line Input #1, str
    Label2.Caption = str
    Line Input #1, str
    Label3.Caption = str
    Close #1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Response As Integer
    If ifFileChanged = True Then
        Response = MsgBox("파일을 저장하시겠습니까?" & vbCrLf & "취소하면 저장하지 않은 내용은 없어집니다.", vbYesNo + vbQuestion, "종료")
        If Response = vbYes Then
            SaveFile
        End If
    End If
End Sub

Private Sub Form_Resize()
On Error GoTo Err
    If frmMain.Height > 900 Then
        Text1.Height = Me.Height - 900 - 255
    End If
    If frmMain.Width > 250 Then
        Text1.Width = Me.Width - 250
    End If
Err:
End Sub

Private Sub mnuAbout_Click()
    frmInfo.Show vbModal
End Sub

Private Sub mnuCompile_Click()
    Shell """" & Label1.Caption & """ " & Label2.Caption & " """ & OpenedFile & """", vbNormalFocus
    MsgBox """" & Label1.Caption & """ " & Label2.Caption & " """ & OpenedFile & """", , "컴파일 중..."
End Sub

Private Sub mnuCopy_Click()
   Clipboard.Clear
   Clipboard.SetText Text1.SelText
End Sub

Private Sub mnuCut_Click()
   Clipboard.Clear
   Clipboard.SetText Text1.SelText
   Text1.SelText = ""
End Sub

Private Sub mnuFind_Click()
    frmFindReplace.Show
End Sub

Private Sub mnuHelp_Click()
    MsgBox "아직 준비 중인 기능입니다.", vbInformation, "Code++"
End Sub

Private Sub mnuOption_Click()
    frmConfig.Show
End Sub

Private Sub mnuPaste_Click()
   Text1.SelText = Clipboard.GetText()
End Sub

Private Sub mnuNew_Click()
    SaveFile
    Text1.Text = ""
    frmMain.Caption = "Untitled - Code++"
    mnuRunApp.Enabled = False
    mnuCompile.Enabled = False
    ifFileOpened = False
    ifFileChanged = False
End Sub

Private Sub mnuOpen_Click()
    LoadFile
    mnuRunApp.Enabled = True
    mnuCompile.Enabled = True
End Sub

Private Sub mnuQuit_Click()
    End
End Sub


Private Sub mnuReplace_Click()
    frmFindReplace.Show
End Sub

Private Sub mnuRunApp_Click()
    Shell """" & Label3.Caption & """", vbNormalFocus
End Sub

Private Sub mnuSave_Click()
If ifFileOpened = True Then
    JustSave
Else
    SaveFile
End If
mnuRunApp.Enabled = True
mnuCompile.Enabled = True
End Sub

Private Sub mnuSaveOtherName_Click()
    SaveFile
End Sub

Private Sub mnuUndo_Click()
    Text1.SetFocus
End Sub

Private Sub mnuSelectAll_Click()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_Change()
    ifFileChanged = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then
        KeyAscii = 32
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "   "
    End If
    
    'If KeyAscii = 123 Then
    '    indent = indent + 1
    'End If
    
    'If KeyAscii = 125 Then
    '    If indent > 0 Then
    '        indent = indent - 1
    '    End If
    'End If
    
    'If KeyAscii = 13 Then
    'Dim i As Integer
    'For i = 1 To indent
    '    Set WshShell = CreateObject("WScript.Shell")
    '    WshShell.SendKeys "    "
    'Next
    'End If
    
    If 0 Then
        '^(\s+)  //search for
        '$1$1$1$1
    End If
End Sub


