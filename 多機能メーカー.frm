VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "多機能メーカー (ホーム)"
   ClientHeight    =   5070
   ClientLeft      =   2865
   ClientTop       =   3645
   ClientWidth     =   14805
   Icon            =   "多機能メーカー.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   14805
   Begin VB.CommandButton jyouho 
      Caption         =   "情報"
      Height          =   495
      Left            =   12120
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton ararm 
      Caption         =   "アラーム"
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton kaunter 
      Caption         =   "カウンター"
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton StopWatch 
      Caption         =   "ストップウォッチ"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Dentaku 
      Caption         =   "電卓"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton shuryo 
      Caption         =   "終了"
      Height          =   495
      Left            =   12120
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Tokei 
      Caption         =   "時計"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      Caption         =   "バージョン　1.O"
      Height          =   255
      Left            =   12000
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "多機能メーカー "
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Top             =   4080
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ararm_Click()
    Form6.Show
End Sub

Private Sub Dentaku_Click()
Form3.Show
End Sub

Private Sub jyouho_Click()
    Form7.Show
End Sub

Private Sub kaunter_Click()
    Form5.Show
End Sub

Private Sub shuryo_Click()
                        Dim Answer As Long
Answer = MsgBox("多機能メーカーを終了しますか？", vbOKCancel Or vbQuestion, "多機能メーカー")

Select Case Answer

    Case vbOK

        MsgBox "多機能メーカーを終了します", vbInformation, "多機能メーカー"
        Call Unload(Me)
End Select
End Sub

Private Sub StopWatch_Click()
    Form4.Show
End Sub

Private Sub taimer_Click()
    Form6.Show
End Sub

Private Sub Tokei_Click()
    Form2.Show
End Sub
