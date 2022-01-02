VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "多機能メーカー (ストップウォッチ)"
   ClientHeight    =   4605
   ClientLeft      =   6450
   ClientTop       =   3645
   ClientWidth     =   7815
   Icon            =   "ストップウォッチ.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4605
   ScaleWidth      =   7815
   Begin VB.CommandButton Command4 
      Caption         =   "ホームへ戻る"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "リセット"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   3360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開始"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
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
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   3975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  
  Timer1.Enabled = True  'タイマーを有効にする
 
End Sub

Private Sub Command2_Click()
  
  Timer1.Enabled = False  'タイマーを無効にする
  
End Sub

Private Sub Command3_Click()
    Set Form4 = Nothing
    Call Unload(Me)
    Form4.Show
End Sub

Private Sub Command4_Click()
            Dim Answer As Long
Answer = MsgBox("多機能メーカー (ストップウォッチ)を終了しますか？", vbOKCancel Or vbQuestion, "多機能メーカー (ストップウォッチ)")

Select Case Answer

    Case vbOK

        MsgBox "多機能メーカー (ストップウォッチ)を終了します", vbInformation, "多機能メーカー (ストップウォッチ)"
        Call Unload(Me)
End Select
End Sub

Private Sub Timer1_Timer()
  
  Static iSec As Integer  '秒
  Static iMin As Integer  '分
  Static iHour As Integer  '時
    
  iSec = iSec + 1  '１秒進める
    
  If iSec >= 60 Then  '６０秒＝１分
    iMin = iMin + 1
    iSec = 0
   
    If iMin >= 60 Then  '６０分＝１時間
     iHour = iHour + 1
     iMin = 0
    End If
   
  End If
  
  '[Label1]に書式を指定して表示
  Text1.Text = Format(iHour, "00") & "：" & _
                   Format(iMin, "00") & "：" & _
                   Format(iSec, "00")
  
End Sub



