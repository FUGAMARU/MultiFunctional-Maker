VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "多機能メーカー (情報)"
   ClientHeight    =   5355
   ClientLeft      =   6240
   ClientTop       =   3645
   ClientWidth     =   10110
   Icon            =   "情報.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   5355
   ScaleWidth      =   10110
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   4440
      Picture         =   "情報.frx":0ECA
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ホームへ戻る"
      Height          =   495
      Left            =   7920
      TabIndex        =   6
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  '中央揃え
      Caption         =   "開発・著作権 ："
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "FUGA SHIMIZU"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   2  '中央揃え
      Caption         =   "ソフトウェア バージョン ："
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      Caption         =   "バージョン 1.0"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      Caption         =   "ソフトウェア タイトル ："
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "多機能メーカー"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   2280
      Width           =   3375
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox "ホームへ戻ります", vbInformation, "多機能メーカー (情報)"
    Call Unload(Me)
End Sub
