VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "���@�\���[�J�[ (�d��)"
   ClientHeight    =   4755
   ClientLeft      =   6675
   ClientTop       =   4080
   ClientWidth     =   8370
   Icon            =   "�d��.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4755
   ScaleWidth      =   8370
   Begin VB.CommandButton Command6 
      Caption         =   "�z�[���֖߂�"
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   2520
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "�v�Z����"
      Height          =   1215
      Left            =   4920
      TabIndex        =   6
      Top             =   1200
      Width           =   2535
      Begin VB.Label Kotae 
         Alignment       =   2  '��������
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�~"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�|"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�{"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Text            =   "�����ɔ��p�Ő��l����͂��Ă��������B"
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "�����ɔ��p�Ő��l����͂��Ă��������B"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "���l����͂��Ă���A�u�{�v�u�|�v�@�u�~�v�u���v�̂ǂꂩ�̃{�^�����N���b�N���Ă��������B"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   4080
      Width           =   6255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    a = Val(Text1.Text)
    b = Val(Text2.Text)
    Kotae.Caption = a + b
End Sub

Private Sub Command2_Click()
    a = Val(Text1.Text)
    b = Val(Text2.Text)
    Kotae.Caption = a - b
End Sub

Private Sub Command3_Click()
    a = Val(Text1.Text)
    b = Val(Text2.Text)
    Kotae.Caption = a * b
End Sub

Private Sub Command4_Click()
    a = Val(Text1.Text)
    b = Val(Text2.Text)
    Kotae.Caption = a / b
End Sub

Private Sub Command5_Click()
    Text1.Text = ""
    Text2.Text = ""
    Kotae.Caption = ""
End Sub

Private Sub Command6_Click()
        Dim Answer As Long
Answer = MsgBox("���@�\���[�J�[ (�d��)���I�����܂����H", vbOKCancel Or vbQuestion, "���@�\���[�J�[ (�d��)")

Select Case Answer

    Case vbOK

        MsgBox "���@�\���[�J�[ (�d��)���I�����܂�", vbInformation, "���@�\���[�J�[ (�d��)"
        Call Unload(Me)
End Select
End Sub

Private Sub Text1_Click()
    Text1.Text = ""
End Sub

Private Sub Text2_Click()
    Text2.Text = ""
End Sub
