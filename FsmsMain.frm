VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FsmsMain 
   BorderStyle     =   1  '單線固定
   Caption         =   "魚貨業買賣管理系統"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   Icon            =   "FsmsMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11745
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame5 
      Caption         =   "進階功能"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   8280
      TabIndex        =   24
      Top             =   3600
      Width           =   3255
      Begin VB.CommandButton Command36 
         Caption         =   "客戶漁獲條件搜尋"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   28
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton Command34 
         Caption         =   "資料回存"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   27
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton Command35 
         Caption         =   "程式說明"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   26
         Top             =   3000
         Width           =   2535
      End
      Begin VB.CommandButton Command32 
         Caption         =   "刪除歷史資料"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   25
         Top             =   480
         Width           =   2535
      End
   End
   Begin MSComDlg.CommonDialog fileDialog 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame0 
      Caption         =   "基本資料維護作業"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   600
      TabIndex        =   20
      Top             =   1080
      Width           =   3255
      Begin VB.CommandButton Command03 
         Caption         =   "稅別資料修改"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton Command01 
         Caption         =   "客戶資料新增與修改"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   0
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton Command02 
         Caption         =   "貨品資料新增與修改"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   1320
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "資料過帳作業"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   600
      TabIndex        =   19
      Top             =   4440
      Width           =   3255
      Begin VB.CommandButton Command43 
         Caption         =   "前月結算"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   22
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton Command42 
         Caption         =   "收款資料過帳"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton Command41 
         Caption         =   "交易資料過帳"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "系統作業"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8280
      TabIndex        =   18
      Top             =   1080
      Width           =   3255
      Begin VB.CommandButton Command31 
         Caption         =   "設定處理日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton Command33 
         Caption         =   "資料備份"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "查詢 / 列印作業"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   4440
      TabIndex        =   17
      Top             =   3600
      Width           =   3255
      Begin VB.CommandButton Command25 
         Caption         =   "營業彙總表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton Command26 
         Caption         =   "對帳單"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   3000
         Width           =   2535
      End
      Begin VB.CommandButton Command23 
         Caption         =   "已過帳－交易明細表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton Command24 
         Caption         =   "已過帳－收款明細表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "交易資料維護作業"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4440
      TabIndex        =   13
      Top             =   1080
      Width           =   3255
      Begin VB.CommandButton Command12 
         Caption         =   "收款資料新增與修改"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton Command11 
         Caption         =   "交易資料新增與修改"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      Caption         =   "視窗請工作於: 800 x 600"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Main_Date 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   21
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Main_DateLabel2 
      Caption         =   "]"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   16
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Main_DateLabel 
      Caption         =   "處理日期：["
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   15
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Main_Title 
      Caption         =   "魚貨業買賣管理系統-FSMS- V1.08"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "FsmsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command01_Click()
  FrmCustom.Show
  Me.Visible = False
End Sub

Private Sub Command02_Click()
  FrmFish.Show
  Me.Visible = False
End Sub

Private Sub Command03_Click()
  FrmTax.Show
  Me.Visible = False
End Sub

Private Sub Command11_Click()
  FrmDeal.Show
  Me.Visible = False
End Sub

Private Sub Command12_Click()
  FrmReceive.Show
  Me.Visible = False
End Sub

Private Sub Command23_Click()
  CheckMode = 2
  FrmCheckIn.Show
  Me.Visible = False
End Sub
Private Sub Command24_Click()
  CheckMode = 3
  FrmCheckIn.Show
  Me.Visible = False
End Sub

Private Sub Command25_Click()
  FrmMonthSummary.Show
  Me.Visible = False
End Sub

Private Sub Command26_Click()
  FrmReceipt.Show
  Me.Visible = False
End Sub

Private Sub Command31_Click()
  FrmCalendar.Show
  Me.Visible = False
End Sub

Private Sub Command32_Click()
  FrmClearDateBase.Show
  Me.Visible = False

End Sub

Private Sub Command33_Click()
  Dim filename As String
  On Error GoTo ErrHandler1
  'Dim filename As String
  'Dim index As Integer
  Dim save_flag As Integer
  Dim cmd_flag As Integer
  Dim fso As FileSystemObject
  Set fso = CreateObject("Scripting.FileSystemObject")

  'fileDialog.FileTitle = "資料庫備份"
  fileDialog.Filter = "FSMS Database (*.mdb)|*.mdb"
  fileDialog.FilterIndex = 0
  fileDialog.filename = fsmsfile
  fileDialog.CancelError = True
  fileDialog.Action = 2
  'filename = fsmsfile
  'fileDialog.ShowOpen
  
  'index = InStrRev(fileDialog.filename, "/")
  'filename = Mid(fileDialog.filename, index + 1, Len(fileDialog.filename) - index - 1)
  
  save_flag = 0
  If fileDialog.filename <> FSMS_file Then
    If fso.FileExists(fileDialog.filename) Then
      cmd_flag = MsgBox("是否覆蓋資料?", vbYesNo, "資料已存在")
      If cmd_flag = 6 Then
        save_flag = 1
      End If
    Else
      save_flag = 1
    End If
  Else
    MsgBox "備份失敗", vbOKOnly + vbExclamation, "錯誤"
  End If
  If save_flag Then
    FileCopy FSMS_file, fileDialog.filename
  End If
ErrHandler1:
  Exit Sub
End Sub

Private Sub Command34_Click()
  On Error GoTo ErrHandler2
  Dim load_flag As Integer
  Dim cmd_flag As Integer
  Dim fso As FileSystemObject
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  'fileDialog.FileTitle = "資料庫回存"
  fileDialog.Filter = "FSMS Database (*.mdb)|*.mdb"
  fileDialog.filename = ""
  fileDialog.FilterIndex = 0
  fileDialog.CancelError = True
  fileDialog.Action = 1
  'fileDialog.ShowOpen
  load_flag = 0
  If fso.FileExists(fileDialog.filename) Then
    load_flag = 1
  End If
  If load_flag Then
    FileCopy fileDialog.filename, FSMS_file
  End If
ErrHandler2:
  Exit Sub

End Sub

Private Sub Command35_Click()
  FrmHelp.Show
  Me.Visible = False
End Sub

Private Sub Command36_Click()
  FrmCondition.Show
  Me.Visible = False
End Sub

Private Sub Command41_Click()
  CheckMode = 0
  FrmCheckIn.Show
  Me.Visible = False
End Sub
Private Sub Command42_Click()
  CheckMode = 1
  FrmCheckIn.Show
  Me.Visible = False
End Sub

Private Sub Command43_Click()
  FrmAccount.Show
  Me.Visible = False
End Sub

Private Sub Form_Activate()
  Main_Date.Caption = update_date
End Sub

Private Sub Form_Load()
  Dim TmpDate As Date
  Dim tmp As String
  Dim conn As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  Dim cmd_flag As Integer
  Call RCYear
  fsms_name = "fsms.mdb"
  FSMS_file = App.Path & "\" & fsms_name
  tmp = App.Path & "tmp.ddd"
  On Error Resume Next
  Dim strFile As String
  Dim oAccess As Object
  Set oAccess = CreateObject("Access.Application")
  oAccess.CompactRepair FSMS_file, tmp, True
  oAccess.Quit
  Set oAccess = Nothing
  If Dir(tmp) <> "" Then
    Kill FSMS_file
    Name tmp As FSMS_file
  End If

  database_string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FSMS_file & ";Persist Security Info=False"
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  vDD = Day(Now)
  vMM = Month(Now)
  vYY = Year(Now)
  CalendarValue = DateValue(Now)
  If vYY > 1900 Then
    vYY = vYY - 1911
  End If
  TmpDate = CDate(vYY & "/" & vMM & "/" & vDD)
  
  conn.Open database_string
  rs.Open "系統資料表", conn
  AccountDate = ""
  If dbDataBase1.Recordset.EOF <> True Then
   AccountDate = rs.Fields("前次結算日期")
  End If
  rs.Close
  conn.Close
  Set conn = Nothing
  If vMM <> Month(AccountDate) Then
    Beep
    cmd_flag = MsgBox("是否結算上月餘額?", vbYesNo, "是否結算")
    If cmd_flag = 6 Then
      FrmAccount.Show
      Me.Visible = False
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call ADYear
End Sub

