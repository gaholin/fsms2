VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   1  '單線固定
   Caption         =   "程式說明"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10950
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   1680
      Picture         =   "FrmHelp.frx":0E42
      ScaleHeight     =   3195
      ScaleWidth      =   6915
      TabIndex        =   2
      Top             =   120
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   0
      Top             =   3600
      Width           =   9135
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  txtHelp = "資料庫處理方式" & vbCrLf
  txtHelp = txtHelp & "  1. 資料庫以Access方式處理, 也可以利用Access方式處理" & vbCrLf
  txtHelp = txtHelp & "  2. 以月為單位做計算" & vbCrLf
  txtHelp = txtHelp & "  3. 公式: {前月結款} + {本月交易} - {本月收款} = {本月結餘}" & vbCrLf
  txtHelp = txtHelp & "  4. 故每當換到新的月份時, 必須要結款上個月餘款" & vbCrLf
  txtHelp = txtHelp & "  5. 每日會有交易及收款, 故每日必須輸入交易資料與收款資料" & vbCrLf
  txtHelp = txtHelp & "  6. 過完帳後, 即可查詢結餘, 及列印本日報表" & vbCrLf
  txtHelp = txtHelp & "" & vbCrLf
  txtHelp = txtHelp & "每日工作" & vbCrLf
  txtHelp = txtHelp & "1. 輸入每日交易及收款款項" & vbCrLf
  txtHelp = txtHelp & "  1) 執行[交易資料新增與修改] -> 輸入交易資料" & vbCrLf
  txtHelp = txtHelp & "  2) 執行[收款資料新增與修改] -> 輸入收款資料" & vbCrLf
  txtHelp = txtHelp & "2. 確定本日資料正確無誤, 將交易及收款過帳" & vbCrLf
  txtHelp = txtHelp & "  1) 執行[交易資料過帳] -> [過帳]" & vbCrLf
  txtHelp = txtHelp & "  2) 執行[收款資料過帳] -> [過帳]" & vbCrLf
  txtHelp = txtHelp & "3. 列印營業彙總表" & vbCrLf
  txtHelp = txtHelp & "  1) [營業彙總表] -> [列印]" & vbCrLf
  txtHelp = txtHelp & "4. 列印對帳單" & vbCrLf
  txtHelp = txtHelp & "  1) 選擇所要列印對帳單的人員, 可利用一一新增, 也可以自動新增" & vbCrLf
  txtHelp = txtHelp & "  2) 選取或取消所要列印之人員" & vbCrLf
  txtHelp = txtHelp & "  3) [對帳單列印] -> [列印]" & vbCrLf
  txtHelp = txtHelp & "5. 備份資料" & vbCrLf
  txtHelp = txtHelp & "  1) 執行[資料備份]" & vbCrLf
  txtHelp = txtHelp & "  2) 開啟存放路徑" & vbCrLf
  txtHelp = txtHelp & "  3) 輸入儲存名稱" & vbCrLf
  txtHelp = txtHelp & "  4) [開啟]" & vbCrLf
  txtHelp = txtHelp & "" & vbCrLf
  txtHelp = txtHelp & "每月工作" & vbCrLf
  txtHelp = txtHelp & "1. 結算前月餘款 (結算上個月餘款)" & vbCrLf
  txtHelp = txtHelp & "  1) 若要結算 99年11月份, 必須將處理日期選擇99年12月份的任何一天" & vbCrLf
  txtHelp = txtHelp & "  2) 執行[前月結算] -> [結算]" & vbCrLf
  txtHelp = txtHelp & "  " & vbCrLf
  txtHelp = txtHelp & "其他" & vbCrLf
  txtHelp = txtHelp & "1. 初始資料" & vbCrLf
  txtHelp = txtHelp & "  1) 建立基本資料 (客戶資料及魚貨資料)" & vbCrLf
  txtHelp = txtHelp & "  2) 建立初始上月結餘/初始過帳交易資料/初始過帳收款資料" & vbCrLf
  txtHelp = txtHelp & "     (必須使用Access匯入或輸入資料)" & vbCrLf
  txtHelp = txtHelp & "2. 當資料錯誤時,可將先前資料給回存" & vbCrLf
  txtHelp = txtHelp & "  1) 執行[資料回存]" & vbCrLf
  txtHelp = txtHelp & "  2) 選擇開啟路徑" & vbCrLf
  txtHelp = txtHelp & "  3) 輸入開啟名稱" & vbCrLf
  txtHelp = txtHelp & "  4) [開啟]" & vbCrLf
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub
