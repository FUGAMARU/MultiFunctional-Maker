VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "���@�\���[�J�[ (�A���[��)"
   ClientHeight    =   4845
   ClientLeft      =   1815
   ClientTop       =   3435
   ClientWidth     =   16755
   Icon            =   "�A���[��.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   4845
   ScaleWidth      =   16755
   Begin VB.CommandButton Command4 
      Caption         =   "�z�[���֖߂�"
      Height          =   495
      Left            =   10680
      TabIndex        =   17
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���Z�b�g"
      Height          =   615
      Left            =   6240
      TabIndex        =   15
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text4 
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
      Height          =   495
      Left            =   5640
      TabIndex        =   13
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�\�������"
      Height          =   615
      Left            =   4680
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "�\��ς�"
      Height          =   1575
      Left            =   9720
      TabIndex        =   10
      Top             =   960
      Width           =   3615
      Begin VB.Label Label7 
         Alignment       =   2  '��������
         Height          =   1215
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8640
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�\��"
      Height          =   615
      Left            =   3120
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text3 
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
      Height          =   495
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text2 
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
      Height          =   495
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   3000
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label9 
      Caption         =   "��F�ߌ�6��4��5�b�Ɂu����΂�́v�ƕ\������ꍇ��18��04��05�b�ɂ���ɂ��͂ƒʒm����B�u�\��v���N���b�N�B"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   4080
      Width           =   8415
   End
   Begin VB.Label Label8 
      Alignment       =   2  '��������
      Caption         =   "�b"
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
      Left            =   6240
      TabIndex        =   14
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  '��������
      Caption         =   "24H"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFFFF&
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
      Left            =   6960
      TabIndex        =   8
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��������
      Caption         =   "�Ɓ@�ʒm����"
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
      Left            =   6840
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      Caption         =   "��"
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
      Left            =   6720
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "��"
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
      Left            =   5040
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "��"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Timer1.Enabled = True
    MsgBox Text2.Text & "��" & Text3.Text & "����" & Text1.Text & "�ƒʒm���܂��B", vbInformation, "���@�\���[�J�[ (�A���[��)"
    
Dim Answer As Long
Answer = MsgBox("��낵���ł����H", vbQuestion Or vbOKCancel, "���@�\���[�J�[ (�A���[��)")

Select Case Answer

Case vbOK

MsgBox "�\�񂵂܂����B", vbInformation, "���@�\���[�J�[ (�A���[��)"
Label7.Caption = Text2.Text & "��" & Text3.Text & "��" & Text4.Text & "�b��" & Text1.Text & "�ƒʒm���܂��B"
End Select

MsgBox "�ʒm�����܂ł� ���@�\���[�J�[���I�����Ȃ��ł��������B", vbInformation, "���@�\���[�J�[ (�A���[��)"
End Sub

Private Sub Command2_Click()
    Set Form6 = Nothing
    Call Unload(Me)
    Form6.Show
End Sub

Private Sub Command3_Click()
    Set Form6 = Nothing
    Call Unload(Me)
    Form6.Show
End Sub

Private Sub Command4_Click()
                    Dim Answer As Long
Answer = MsgBox("���@�\���[�J�[ (�A���[��)���I�����܂����H", vbOKCancel Or vbQuestion, "���@�\���[�J�[ (�A���[��)")

Select Case Answer

    Case vbOK

        MsgBox "���@�\���[�J�[ (�A���[��)���I�����܂�", vbInformation, "���@�\���[�J�[ (�A���[��)"
        Call Unload(Me)
End Select
End Sub

Private Sub Timer1_Timer()
    Label5.Caption = "���ԁF" & Time
    If Label5.Caption = "���ԁF" & Text2.Text & ":" & Text3.Text & ":" & Text4.Text Then
    
Dim Answer As Long
Answer = MsgBox(Text1.Text, vbInformation Or vbOKOnly, "���@�\���[�J�[ (�A���[��)")

Select Case Answer

    Case vbOK
        
   Set Form6 = Nothing
    Call Unload(Me)
    Form6.Show
End Select
End If
End Sub
