VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "도움말"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7110
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdOK 
      Caption         =   "나가기(&Q)"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      Height          =   5175
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmHelp.frx":1542
      Top             =   120
      Width           =   4575
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   9128
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
