VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "���@�\���[�J�[ (�J�E���^�[)"
   ClientHeight    =   3810
   ClientLeft      =   6675
   ClientTop       =   4290
   ClientWidth     =   7410
   Icon            =   "�J�E���^�[.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   3810
   ScaleWidth      =   7410
   Begin VB.CommandButton Command3 
      Caption         =   "�z�[���֖߂�"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���Z�b�g"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�J�E���g"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer

Private Sub Command1_Click()
    n = n + 1
Text1.Text = n
End Sub

Private Sub Command2_Click()
    Set Form5 = Nothing
    Call Unload(Me)
    Form5.Show
End Sub

Private Sub Command3_Click()
                Dim Answer As Long
Answer = MsgBox("���@�\���[�J�[ (�J�E���^�[)���I�����܂����H", vbOKCancel Or vbQuestion, "���@�\���[�J�[ (�J�E���^�[)")

Select Case Answer

    Case vbOK

        MsgBox "���@�\���[�J�[ (�J�E���^�[)���I�����܂�", vbInformation, "���@�\���[�J�[ (�J�E���^�[)"
        Call Unload(Me)
End Select
End Sub
