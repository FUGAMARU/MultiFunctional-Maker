VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "���@�\���[�J�[ (���v)"
   ClientHeight    =   3990
   ClientLeft      =   8355
   ClientTop       =   4500
   ClientWidth     =   9045
   Icon            =   "���v.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3990
   ScaleWidth      =   9045
   Begin VB.CommandButton Command5 
      Caption         =   "�z�[���֖߂�"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���Ԃ�\������"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   1920
      Width           =   1755
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���t��\������"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   1080
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���t��\�����Ȃ�"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox hiduke 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���Ԃ�\�����Ȃ�"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4320
      Top             =   3000
   End
   Begin VB.TextBox jikan 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4320
      Top             =   3480
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    jikan.Text = ""
    Timer1.Enabled = False
End Sub

Private Sub Command2_Click()
    hiduke.Text = ""
    Timer2.Enabled = False
End Sub

Private Sub Command3_Click()
    Timer2.Enabled = True
End Sub

Private Sub Command4_Click()
    Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
    Dim Answer As Long
Answer = MsgBox("���@�\���[�J�[ (���v)���I�����܂����H", vbOKCancel Or vbQuestion, "���@�\���[�J�[ (���v)")

Select Case Answer

    Case vbOK

        MsgBox "���@�\���[�J�[ (���v)���I�����܂�", vbInformation, "���@�\���[�J�[ (���v)"
        Call Unload(Me)
End Select
End Sub

Private Sub Timer1_Timer()
    jikan.Text = "        ���ԁF    " & Time
End Sub

Private Sub Timer2_Timer()
    hiduke.Text = "���t:    " & Date
End Sub
