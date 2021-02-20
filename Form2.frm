VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Code++ 정보"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7665
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "Naissoft Code++는 오픈 소스 소프트웨어입니다. 상업 및 재배포 목적을 제외한 어떤 목적이라도 코드를 활용할 수 있습니다."
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "ⓒ 2016~2021 Naissoft. All rights reserved."
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   3615
   End
   Begin VB.Label lblVer 
      BackStyle       =   0  '투명
      Caption         =   "lblVer"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "Naissoft Code++"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   0
      Picture         =   "Form2.frx":1542
      Top             =   0
      Width           =   7680
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload frmInfo
End Sub

Private Sub Form_Load()
    lblVer.Caption = "Version " & App.Major & "." & App.Minor & " (Build " & App.Revision & ")"
End Sub

