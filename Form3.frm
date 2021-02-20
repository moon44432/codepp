VERSION 5.00
Begin VB.Form frmFindReplace 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "ã�� / �ٲٱ�"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5760
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdReplaceAll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��� �ٲٱ�"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ã��(&F)"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ã�� / �ٲٱ�"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton cmdReplace 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ٲٱ�(&R)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdFindNext 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���� ã��"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Text            =   "�ٲ� �ؽ�Ʈ"
         Top             =   1440
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "ã�� �ؽ�Ʈ"
         Top             =   360
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PostFind As Long
Private TextFind As String
Private FindText As Long

Private Sub cmdFind_Click()
    cmdFindNext.Enabled = False
    cmdReplace.Enabled = False
    TextFind = Text1.Text
    If TextFind & "" <> "" Then
        PostFind = 1
        cmdFindNext.Enabled = True
        cmdReplace.Enabled = True
        cmdReplaceAll.Enabled = True
        cmdFindNext_Click
    End If
End Sub

Private Sub cmdFindNext_Click()
frmMain.Text1.SetFocus
FindText = InStr(PostFind, frmMain.Text1, TextFind)
If FindText > 0 Then
    frmMain.Text1.SelStart = FindText - 1
    frmMain.Text1.SelLength = Len(TextFind)
    cmdReplace.Enabled = True
    cmdReplaceAll.Enabled = True
    PostFind = FindText + 1
    cmdReplace.Enabled = True
    cmdReplaceAll.Enabled = True
Else
    MsgBox "�˻��� �������ϴ�.", vbExclamation, "ã��"
End If
End Sub

Private Sub cmdReplace_Click()
   Dim TextReplace As String
    TextReplace = Text2.Text
    If TextReplace & "" <> "" Then
        'Replace only one which is selected
        If (frmMain.Text1.SelText = "") = False Then
            frmMain.Text1.SelText = TextReplace
        End If
    End If
End Sub


Private Sub cmdReplaceAll_Click()
    Dim TextReplace As String
    TextReplace = Text2.Text
    If TextReplace & "" <> "" Then
        Dim QueryReplace As Integer
        QueryReplace = MsgBox("��� �ٲٽðڽ��ϱ�?", vbYesNo + vbExclamation, "�ٲٱ�")
        frmMain.Text1.SetFocus
        If QueryReplace = vbYes Then
           'Replace all in the text
           FindText = 1
           Do Until FindText = 0
                FindText = InStr(FindText, frmMain.Text1, TextFind)
                If FindText > 0 Then
                    frmMain.Text1.SelStart = FindText - 1
                    frmMain.Text1.SelLength = Len(TextFind)
                    frmMain.Text1.SelText = TextReplace
                End If
           Loop
        End If
    End If
End Sub
