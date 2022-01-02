VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "多機能メーカー (電卓)"
   ClientHeight    =   4755
   ClientLeft      =   6675
   ClientTop       =   4080
   ClientWidth     =   8370
   Icon            =   "電卓.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4755
   ScaleWidth      =   8370
   Begin VB.CommandButton Command6 
      Caption         =   "ホームへ戻る"
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "計算結果"
      Height          =   1215
      Left            =   4920
      TabIndex        =   6
      Top             =   1200
      Width           =   2535
      Begin VB.Label Kotae 
         Alignment       =   2  '中央揃え
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "－"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "＋"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Text            =   "ここに半角で数値を入力してください。"
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "ここに半角で数値を入力してください。"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "数値を入力してから、「＋」「－」　「×」「÷」のどれかのボタンをクリックしてください。"
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
Answer = MsgBox("多機能メーカー (電卓)を終了しますか？", vbOKCancel Or vbQuestion, "多機能メーカー (電卓)")

Select Case Answer

    Case vbOK

        MsgBox "多機能メーカー (電卓)を終了します", vbInformation, "多機能メーカー (電卓)"
        Call Unload(Me)
End Select
End Sub

Private Sub Text1_Click()
    Text1.Text = ""
End Sub

Private Sub Text2_Click()
    Text2.Text = ""
End Sub
