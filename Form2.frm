VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '¥‹¿œ ∞Ì¡§
   Caption         =   "Code++ ¡§∫∏"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7665
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows ±‚∫ª∞™
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'ºˆ¡˜
      TabIndex        =   4
      Text            =   "Form2.frx":1542
      Top             =   4080
      Width           =   7215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "»Æ¿Œ"
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
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
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '≈ı∏Ì
      Caption         =   "®œ 2016~2021 Naissoft. All rights reserved."
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
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
      Top             =   5640
      Width           =   3615
   End
   Begin VB.Label lblVer 
      BackStyle       =   0  '≈ı∏Ì
      Caption         =   "lblVer"
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
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
      BackStyle       =   0  '≈ı∏Ì
      Caption         =   "Naissoft Code++"
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
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
      Picture         =   "Form2.frx":15B5
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

