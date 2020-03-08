VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmReceipt 
   BorderStyle     =   1  '單線固定
   Caption         =   "對帳單"
   ClientHeight    =   8205
   ClientLeft      =   285
   ClientTop       =   915
   ClientWidth     =   13830
   Icon            =   "FrmReceipt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13830
   Begin VB.Frame Frame4 
      Caption         =   "自動新增"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7680
      TabIndex        =   23
      Top             =   840
      Width           =   2055
      Begin VB.CommandButton cmdAutoAdd 
         Caption         =   "自動新增"
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
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "所有人員"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   24
         ToolTipText     =   "未勾擇表示, 今日交易人員"
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Height          =   735
      Left            =   720
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton cmdPage 
         Caption         =   "第一頁"
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPage 
         Caption         =   "上一頁"
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   3000
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPage 
         Caption         =   "下一頁"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   4320
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPage 
         Caption         =   "最後一頁"
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   5640
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "列印"
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblPage 
         Alignment       =   2  '置中對齊
         Caption         =   "0 / 0"
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
         Left            =   7080
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   2520
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Label Label4 
         Caption         =   "資料處理中...請稍後"
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
         Left            =   1680
         TabIndex        =   15
         Top             =   1560
         Width           =   2295
      End
   End
   Begin RichTextLib.RichTextBox Preview 
      Height          =   6015
      Left            =   840
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10610
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"FrmReceipt.frx":0E42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox Preview1 
      Height          =   855
      Left            =   720
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmReceipt.frx":0EE7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "細明體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   720
   End
   Begin MSAdodcLib.Adodc dbDataBase3 
      Height          =   330
      Left            =   7200
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dbDataBase3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc dbDataBase2 
      Height          =   330
      Left            =   5280
      Top             =   120
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dbDataBase2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc dbDataBase1 
      Height          =   330
      Left            =   3360
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dbDataBase1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "新增範圍"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   6735
      Begin VB.TextBox cid 
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
         Index           =   1
         Left            =   3600
         MaxLength       =   3
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox cid 
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
         Index           =   0
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdAddCustom 
         Caption         =   "新增"
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
         Left            =   5040
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker checkDate 
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364737
         CurrentDate     =   40345
      End
      Begin MSComCtl2.DTPicker checkDate 
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364737
         CurrentDate     =   40345
      End
      Begin VB.Label Label3 
         Caption         =   "日期區間"
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
         Left            =   600
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "客戶區間                         至"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   1200
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView listReport 
      Height          =   4695
      Left            =   720
      TabIndex        =   1
      Top             =   2880
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8281
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "客戶編號"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "客戶姓名"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "起始日期"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "結束日期"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "前期結額"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "合計結欠"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7590
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   13388
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "對帳單設定"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "對帳單列印"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog PrintDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox PreviewTmp 
      Height          =   2055
      Left            =   11520
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3625
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmReceipt.frx":0F8C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TotalList As Integer
Private StartId As String
Private StopId As String
Private StartDate As String
Private StopDate As String
Private Const REPORT_COLUMN_SPAN = "          "
Private Const REPORT_ROW_SPAN = "---------------------------------------------"
Private Const REPORT_FINAL_SPAN = "                    (聯)                     "
Private Const REPORT_NULL_SPAN = "                                             "
Private UpdateFlag As Integer
Private TimerMode As Integer
Private PageBuf(100) As String
Private NumList As Byte
Private ListBufPage(120) As String
Private NumListBuf14 As Byte
Private ListBuf14pt(12) As String
Private ListBuf14ptLine(12) As Byte
Private ListBuf14ptType(12) As Byte

Private AutoCid(1024) As String
Private AutoName(1024) As String
Private AutoMoney(1024) As Long
Private TotalPage As Integer
Private CurrentPage As Integer
Private Const NUM_LIST_PRE_PAGE = 14 '17

Private Const PRINTER_X_OFFSET = 100
Private Const PRINTER_Y_OFFSET = 0
Private Const PRINTER_X2_OFFSET = 6223
Private Const PRINTER_MONEY_X1_OFFSET = 1200
Private Const PRINTER_MONEY_X2_OFFSET = 3280

Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, _
ByVal wp As Long, lp As Any) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" _
  (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const WM_USER = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57

Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Type Rect
   Left           As Long
   Top            As Long
   Right          As Long
   Bottom         As Long
End Type

Private Type CharRange
   cpMin          As Long
   cpMax          As Long
End Type

Private Type FormatRange
   hdc            As Long
   hdcTarget      As Long
   rc             As Rect
   rcPage         As Rect
   chrg           As CharRange
End Type

Private Function getSummary(customId As String, procDate As Date) As Long
  Dim cmd As String
  Dim StrDate As String
  Dim DateYear As Integer
  Dim DateMonth As Integer
  
  DateYear = Year(procDate)
  DateMonth = Month(procDate)
  StrDate = DateYear & "/" & DateMonth & "/1"
  
  'cmd = "SELECT 營業彙總表.客戶編號, Sum(營業彙總表.結算金額) AS 期初餘額, Sum(營業彙總表.交易金額) "
  'cmd = cmd & "AS 本期銷售, Sum(營業彙總表.收款金額) AS 本期收款, [期初餘額]+[本期銷售]-[本期收款] AS 應收餘額 "
  'cmd = cmd & "FROM 客戶資料表 INNER JOIN 營業彙總表 ON 客戶資料表.客戶編號 = 營業彙總表.客戶編號 "
  'cmd = cmd & "WHERE (((營業彙總表.結算日期) Between #" & StrDate & "# And #" & CStr(procDate) & "#)) "
  'cmd = cmd & "GROUP BY 營業彙總表.客戶編號 "
  'cmd = cmd & "HAVING ((營業彙總表.客戶編號) = '" & customId & "')"
  'dbDataBase2.CommandType = adCmdText
  'dbDataBase2.RecordSource = cmd
  'dbDataBase2.Refresh
  'If dbDataBase2.Recordset.EOF = True Then
  '  getSummary = 0
  'Else
  '  getSummary = dbDataBase2.Recordset.Fields("應收餘額")
  'End If
  cmd = "SELECT 結算資料表.結算金額 "
  cmd = cmd & "FROM 結算資料表 "
  cmd = cmd & "WHERE (((結算資料表.客戶編號)='" & customId & "') AND ((結算資料表.結算日期)=#" & StrDate & "#));"
  dbDataBase2.CommandType = adCmdText
  dbDataBase2.RecordSource = cmd
  dbDataBase2.Refresh
  If dbDataBase2.Recordset.EOF = True Then
    getSummary = 0
  Else
    getSummary = dbDataBase2.Recordset.Fields("結算金額")
  End If
  'cmd = "SELECT Sum(過帳交易資料表.合計) AS 合計之總計, 過帳交易資料表.客戶編號 "
  'cmd = cmd & "FROM 過帳交易資料表 "
  'cmd = cmd & "WHERE (((過帳交易資料表.客戶編號)='" & customId & "') AND ((過帳交易資料表.交易日期) Between #" & StrDate & "# And #" & procDate & "#));"
  cmd = "SELECT Sum(過帳交易資料表.合計) AS 合計之總計, 過帳交易資料表.客戶編號 "
  cmd = cmd & "FROM 過帳交易資料表 "
  cmd = cmd & "WHERE (((過帳交易資料表.交易日期) Between #" & StrDate & "# And #" & procDate & "#)) "
  cmd = cmd & "GROUP BY 過帳交易資料表.客戶編號 "
  cmd = cmd & "HAVING (((過帳交易資料表.客戶編號)='" & customId & "'));"

  dbDataBase2.CommandType = adCmdText
  dbDataBase2.RecordSource = cmd
  dbDataBase2.Refresh
  If dbDataBase2.Recordset.EOF <> True Then
    getSummary = getSummary + dbDataBase2.Recordset.Fields("合計之總計")
  End If
  'cmd = "SELECT Sum(過帳收款資料表.收款金額) AS 收款金額之總計 "
  'cmd = cmd & "FROM 過帳收款資料表 "
  'cmd = cmd & "WHERE (((過帳收款資料表.客戶編號)='" & customId & "') AND ((過帳收款資料表.收款日期) Between #" & StrDate & "# And #" & procDate & "#));"
  cmd = "SELECT Sum(過帳收款資料表.收款金額) AS 收款金額之總計, 過帳收款資料表.客戶編號 "
  cmd = cmd & "FROM 過帳收款資料表 "
  cmd = cmd & "WHERE (((過帳收款資料表.收款日期) Between #" & StrDate & "# And #" & procDate & "#)) "
  cmd = cmd & "GROUP BY 過帳收款資料表.客戶編號 "
  cmd = cmd & "HAVING (((過帳收款資料表.客戶編號)='" & customId & "'));"
  dbDataBase2.CommandType = adCmdText
  dbDataBase2.RecordSource = cmd
  dbDataBase2.Refresh
  If dbDataBase2.Recordset.EOF <> True Then
    getSummary = getSummary - dbDataBase2.Recordset.Fields("收款金額之總計")
  End If
End Function

Private Sub checkDate_Change(Index As Integer)
  If Index = 0 Then
    StartDate = checkDate(0)
  ElseIf Index = 1 Then
    StopDate = checkDate(1)
  End If
End Sub

Private Sub cid_Change(Index As Integer)
  'If cid(Index) <> "" Then
  '  If IsNumeric(cid(Index)) Then
  '    If Index = 0 Then
  '      StartId = String(3 - Len(cid(Index)), "0") & cid(Index)
  '    Else
  '      StopId = String(3 - Len(cid(Index)), "0") & cid(Index)
  '    End If
  '  End If
  'End If
  
  
  
  If cid(0) = "" And cid(1) = "" Then
    StartId = "000"
    StopId = "999"
  Else
    If cid(0) = "" Then
    StartId = "000"
    Else
      StartId = String(3 - Len(cid(0)), "0") & cid(0)
    End If
    If cid(1) = "" Then
      StopId = StartId
    Else
      StopId = String(3 - Len(cid(1)), "0") & cid(1)
    End If
  End If
  If StartId > StopId Then
    Beep
    cid(1).SetFocus
  End If
End Sub

Private Sub cId_KeyPress(Index As Integer, KeyAscii As Integer)
  If ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 8)) = False Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub cid_LostFocus(Index As Integer)
  'If cid(Index) <> "" Then
  '  If IsNumeric(cid(Index)) Then
  '    cid(Index) = String(3 - Len(cid(Index)), "0") & cid(Index)
  '  End If
  'End If
  If Index = 0 Then
    If cid(0) <> "" Then
      cid(0) = String(3 - Len(cid(0)), "0") & cid(0)
    End If
  ElseIf cid(1) <> "" Then
    cid(1) = String(3 - Len(cid(1)), "0") & cid(1)
  End If
End Sub
Private Sub AutoFun()
  Dim AutoCnt As Integer
  Dim cmd As String
  Dim i, j As Integer
  Dim TargetMoney As Long
  Dim SumMoney As Long
  Dim CurrentMoney As Long
  Dim LastDate As String
  Dim firstDate As String
  Dim item As String
  Dim list_item As ListItem
  Dim TmpDate As Date
  Dim Index As Integer
  Dim Date1 As String
  Dim Date2 As String
  Dim Date3 As String
  
  AutoCnt = 0
  UpdateFlag = 1

  If vYY < 1900 Then
    Date1 = CStr(vYY + 1911) & "/" & vMM & "/1"
    Date2 = CStr(vYY + 1911) & "/" & String(2 - Len(CStr(vMM)), "0") & vMM & "/" & String(2 - Len(CStr(vDD)), "0") & vDD
  Else
    Date1 = vYY & "/" & vMM & "/1"
    Date2 = vYY & "/" & String(2 - Len(CStr(vMM)), "0") & vMM & "/" & String(2 - Len(CStr(vDD)), "0") & vDD
  End If
  cmd = "SELECT 營業彙總表.客戶編號, 客戶資料表.客戶姓名, Sum(營業彙總表.結算金額) "
  cmd = cmd & "AS 期初餘額, Sum(營業彙總表.交易金額) AS 本期銷售, Sum(營業彙總表.收款金額) "
  cmd = cmd & "AS 本期收款, [期初餘額]+[本期銷售]-[本期收款] AS 應收餘額 "
  cmd = cmd & "FROM 客戶資料表 INNER JOIN 營業彙總表 ON 客戶資料表.客戶編號=營業彙總表.客戶編號 "
  cmd = cmd & "WHERE (((營業彙總表.結算日期) Between #" & Date1 & "# And #" & Date2 & "#))"
  cmd = cmd & "GROUP BY 營業彙總表.客戶編號, 客戶資料表.客戶姓名;"
  dbDataBase2.CommandType = adCmdText
  dbDataBase2.RecordSource = cmd
  dbDataBase2.Refresh
  Do Until dbDataBase2.Recordset.EOF
    item = dbDataBase2.Recordset.Fields("客戶編號")
    AutoCid(AutoCnt) = item
    item = dbDataBase2.Recordset.Fields("客戶姓名")
    AutoName(AutoCnt) = item
    item = dbDataBase2.Recordset.Fields("應收餘額")
    AutoMoney(AutoCnt) = CLng(item)
    dbDataBase2.Recordset.MoveNext
    AutoCnt = AutoCnt + 1
  Loop
  
  For i = 1 To AutoCnt
    Index = 0
    For j = 1 To TotalList
      If AutoCid(i - 1) = listReport.ListItems(j).Text Then
        Index = j
        Exit For
      End If
    Next
    'cmd = "SELECT [40_交易收款日期查詢].客戶編號, 客戶資料表.客戶姓名, Min([40_交易收款日期查詢].交易日期) "
    'cmd = cmd & "AS 起始日期, Max([40_交易收款日期查詢].交易日期) AS 結束日期 "
    'cmd = cmd & "FROM 客戶資料表 INNER JOIN 40_交易收款日期查詢 ON 客戶資料表.客戶編號 = [40_交易收款日期查詢].客戶編號 "
    'cmd = cmd & "WHERE (([40_交易收款日期查詢].交易日期) Between #" & Date1 & "# And #" & Date2 & "#) "
    'cmd = cmd & "GROUP BY [40_交易收款日期查詢].客戶編號, 客戶資料表.客戶姓名 "
    'cmd = cmd & "HAVING (([40_交易收款日期查詢].客戶編號) Between '" & AutoCid(i - 1) & "' And '" & AutoCid(i - 1) & "');"
    'dbDataBase2.CommandType = adCmdText
    'dbDataBase2.RecordSource = cmd
    'dbDataBase2.Refresh
    'If dbDataBase2.Recordset.EOF = False And dbDataBase2.Recordset.BOF = False Then
    'End If
    cmd = "SELECT 過帳收款資料表.客戶編號, Max(過帳收款資料表.收款日期) AS 收款日期之最大值 "
    cmd = cmd & "From 過帳收款資料表 "
    cmd = cmd & "GROUP BY 過帳收款資料表.客戶編號 "
    cmd = cmd & "HAVING (((過帳收款資料表.客戶編號)='" & AutoCid(i - 1) & "'));"
    dbDataBase2.CommandType = adCmdText
    dbDataBase2.RecordSource = cmd
    dbDataBase2.Refresh
    LastDate = ""
    If dbDataBase2.Recordset.BOF = False Then
      Date3 = dbDataBase2.Recordset.Fields("收款日期之最大值")
      If Date3 <= Date2 Then
        LastDate = dbDataBase2.Recordset.Fields("收款日期之最大值")
      End If
    End If
    

    'cmd = "SELECT 過帳交易資料表.客戶編號, 客戶資料表.客戶姓名, 過帳交易資料表.交易日期, "
    'cmd = cmd & "Sum(過帳交易資料表.合計) AS 合計之總計 "
    'cmd = cmd & "FROM 過帳交易資料表 INNER JOIN 客戶資料表 ON 過帳交易資料表.客戶編號 = 客戶資料表.客戶編號 "
    'cmd = cmd & "GROUP BY 過帳交易資料表.客戶編號, 客戶資料表.客戶姓名, 過帳交易資料表.交易日期 "
    'cmd = cmd & "HAVING (((過帳交易資料表.客戶編號)='" & AutoCid(i - 1) & "')) "
    'cmd = cmd & "ORDER BY 過帳交易資料表.交易日期 DESC;"
    cmd = "SELECT 過帳交易資料表.客戶編號, 過帳交易資料表.交易日期, "
    cmd = cmd & "Sum(過帳交易資料表.合計) AS 合計之總計 "
    cmd = cmd & "FROM 過帳交易資料表 "
    cmd = cmd & "GROUP BY 過帳交易資料表.客戶編號, 過帳交易資料表.交易日期 "
    cmd = cmd & "HAVING (((過帳交易資料表.客戶編號)='" & AutoCid(i - 1) & "')) "
    cmd = cmd & "ORDER BY 過帳交易資料表.交易日期 DESC;"
    dbDataBase2.CommandType = adCmdText
    dbDataBase2.RecordSource = cmd
    dbDataBase2.Refresh
    SumMoney = 0
    firstDate = ""
    TargetMoney = AutoMoney(i - 1)
    If dbDataBase2.Recordset.BOF = False Then
      If LastDate = "" Then
        LastDate = dbDataBase2.Recordset.Fields("交易日期")
      ElseIf LastDate < dbDataBase2.Recordset.Fields("交易日期") Then
        LastDate = dbDataBase2.Recordset.Fields("交易日期")
      End If
    End If
    
    If chkAll.value = 1 Or LastDate = Date2 Then
      While dbDataBase2.Recordset.EOF = False And SumMoney < TargetMoney
        CurrentMoney = dbDataBase2.Recordset.Fields("合計之總計")
        If SumMoney < TargetMoney Then
          firstDate = dbDataBase2.Recordset.Fields("交易日期")
          SumMoney = SumMoney + CurrentMoney
        End If
        dbDataBase2.Recordset.MoveNext
      Wend
      If LastDate <> "" And firstDate <> "" Then
        If Index = 0 Then
          Set list_item = listReport.ListItems.Add(, , AutoCid(i - 1))
          TotalList = TotalList + 1
        Else
          Set list_item = listReport.ListItems(Index)
        End If
        list_item.SubItems(1) = AutoName(i - 1)
        list_item.SubItems(2) = DC2PC(CDate(firstDate))
        list_item.SubItems(3) = DC2PC(CDate(LastDate))
        list_item.SubItems(5) = TargetMoney
        TmpDate = CDate(firstDate)
        TmpDate = DateAdd("d", -1, TmpDate)
        TargetMoney = getSummary(AutoCid(i - 1), TmpDate)
        list_item.SubItems(4) = TargetMoney
        list_item.Checked = True
      End If
    End If
  Next
End Sub
  
Private Sub AddFun()
  Dim cmd As String
  Dim customId As String
  Dim cName As String
  Dim dateStart As String
  Dim dateStop As String
  Dim idStart As String
  Dim idStop As String
  Dim Index As Integer
  Dim list_item As ListItem
  Dim i As Integer
  Dim SumMoney As Long
  Dim TmpDate As Date
  Dim SkipMach As Boolean
  
  UpdateFlag = 1
  cmd = "SELECT [40_交易收款日期查詢].客戶編號, 客戶資料表.客戶姓名, Min([40_交易收款日期查詢].交易日期) "
  cmd = cmd & "AS 起始日期, Max([40_交易收款日期查詢].交易日期) AS 結束日期 "
  cmd = cmd & "FROM 客戶資料表 INNER JOIN 40_交易收款日期查詢 ON 客戶資料表.客戶編號 = [40_交易收款日期查詢].客戶編號 "
  cmd = cmd & "WHERE (([40_交易收款日期查詢].交易日期) Between #" & StartDate & "# And #" & StopDate & "#) "
  cmd = cmd & "GROUP BY [40_交易收款日期查詢].客戶編號, 客戶資料表.客戶姓名 "
  cmd = cmd & "HAVING (([40_交易收款日期查詢].客戶編號) Between '" & StartId & "' And '" & StopId & "');"
  dbDataBase3.CommandType = adCmdText
  dbDataBase3.RecordSource = cmd
  dbDataBase3.Refresh
  'If StartId = "000" And StopId = "999" Then
  '  SkipMach = True
  '  For i = 1 To TotalList
  '    list_item.ListItems.Remove i
  '  Next
  '  TotalList = 0
  'End If
  'If TotalList = 0 Then
  '  SkipMach = True
  'End If
  Index = 0
  Do Until dbDataBase3.Recordset.EOF
    customId = dbDataBase3.Recordset.Fields("客戶編號")
    cName = dbDataBase3.Recordset.Fields("客戶姓名")
    dateStart = dbDataBase3.Recordset.Fields("起始日期")
    dateStop = dbDataBase3.Recordset.Fields("結束日期")
    'If SkipMach = False Then
      Index = 0
      For i = 1 To TotalList
        If customId = listReport.ListItems(i).Text Then
          Index = i
          Exit For
        End If
      Next
    'End If
    If Index <> 0 Then
      Set list_item = listReport.ListItems(Index)
    Else
      Set list_item = listReport.ListItems.Add(, , customId)
      TotalList = TotalList + 1
    End If
    list_item.SubItems(1) = cName
    list_item.SubItems(2) = DC2PC(CDate(dateStart))
    list_item.SubItems(3) = DC2PC(CDate(dateStop))
    SumMoney = getSummary(customId, CDate(dateStop))
    list_item.SubItems(5) = SumMoney
    TmpDate = CDate(dateStart)
    TmpDate = DateAdd("d", -1, TmpDate)
    SumMoney = getSummary(customId, TmpDate)
    list_item.SubItems(4) = SumMoney
    list_item.Checked = True
    dbDataBase3.Recordset.MoveNext
  Loop
End Sub
Private Sub WaitProc()
  Frame1.Visible = True
  Frame2.Enabled = False
  Frame4.Enabled = False
  listReport.Enabled = False
  TabStrip1.Enabled = False
  Preview.Enabled = False
  cmdPrint.Enabled = False
End Sub
Private Sub ProcDone()
  Frame1.Visible = False
  Frame2.Enabled = True
  Frame4.Enabled = True
  listReport.Enabled = True
  If TotalPage <> 0 Then
    cmdPrint.Enabled = True
  End If
  TabStrip1.Enabled = True
  Preview.Enabled = True
End Sub
Private Sub cmdAutoAdd_Click()
  Call WaitProc
  TimerMode = 1
  Timer1.Enabled = True
End Sub
Private Sub cmdAddCustom_Click()
  Call WaitProc
  TimerMode = 2
  Timer1.Enabled = True
End Sub
Private Sub Print_Receipt()
  If UpdateFlag = 1 Then
    Call WaitProc
    TimerMode = 3
    Timer1.Enabled = True
  Else
    If TotalPage <> 0 Then
      cmdPrint.Enabled = True
      If TotalPage <> 1 Then
        cmdPage(2).Enabled = True
        cmdPage(3).Enabled = True
      End If
      Preview.Text = PageBuf(CurrentPage)
      lblPage = CurrentPage & " / " & TotalPage
    End If
  End If
End Sub

Private Sub ReceiptFun()
  Dim i, j, k As Integer
  Dim customId As String
  Dim customName As String
  Dim dateStart As String
  Dim dateStop As String
  Dim dateStartPrevsum As Date
  Dim dateStopPrevsum As Date
  Dim list_item As ListItem
  Dim cmd As String
  Dim prevSum As Long
  Dim finalSum As Long
  Dim dealSum As Long
  Dim receiveSum As Long
  Dim UintKg As Double
  Dim UintBag As Integer
  
  Dim fDate As String
  Dim fName As String
  Dim fWeight As String
  Dim fUnit As String
  Dim fMoney As String
  Dim dMoney As String
  
  Dim ListNum As Integer
  Dim ListBuf(110) As String
  Dim ListHeader(4) As String
  Dim SmallPageBuf(32) As String
  Dim toggleList As Integer
  Dim CurrentList As Integer
  Dim NumSmallPage As Integer
  Dim SmallPageCnt As Integer
  
  UpdateFlag = 0
  toggleList = 0
  NumSmallPage = 0
  ListNum = 0
  Preview.Text = ""
  cmdPrint.Enabled = False
  SmallPageCnt = 0
  TotalPage = 0
  For k = 1 To TotalList
  'k = 1
    Set list_item = listReport.ListItems(k)
    If list_item.Checked = True Then
      customId = list_item.Text
      customName = list_item.SubItems(1)
      dateStart = PC2DC(list_item.SubItems(2))
      dateStop = PC2DC(list_item.SubItems(3))
      prevSum = list_item.SubItems(4)
      finalSum = list_item.SubItems(5)
      cmd = "SELECT 過帳交易資料表.客戶編號, Sum(過帳交易資料表.重量) AS 重量之總計 "
      cmd = cmd & "FROM 過帳交易資料表 "
      cmd = cmd & "WHERE (((過帳交易資料表.交易日期) Between #" & dateStart & "# And #" & dateStop & "#))"
      cmd = cmd & "GROUP BY 過帳交易資料表.客戶編號 "
      cmd = cmd & "HAVING (((過帳交易資料表.客戶編號)='" & customId
      cmd = cmd & "'));"
      dbDataBase1.CommandType = adCmdText
      dbDataBase1.RecordSource = cmd
      dbDataBase1.Refresh
      If dbDataBase1.Recordset.EOF = False Then
        UintKg = dbDataBase1.Recordset.Fields("重量之總計")
      Else
        UintKg = 0
      End If
      'cmd = "SELECT 過帳交易資料表.客戶編號, 過帳交易資料表.單位, Sum(過帳交易資料表.重量) AS 重量之總計 "
      'cmd = cmd & "FROM 過帳交易資料表 "
      'cmd = cmd & "GROUP BY 過帳交易資料表.客戶編號, 過帳交易資料表.交易日期, 過帳交易資料表.單位 "
      'cmd = cmd & "HAVING (((過帳交易資料表.客戶編號) = '" & customId
      'cmd = cmd & "') AND ((過帳交易資料表.交易日期) Between #" & dateStart & "# And #" & dateStop
      'cmd = cmd & "#) AND ((過帳交易資料表.單位)='台斤'));"
      'dbDataBase1.RecordSource = cmd
      'dbDataBase1.Refresh
      'If dbDataBase1.Recordset.EOF = False Then
      '  UintKg = UintKg + dbDataBase1.Recordset.Fields("重量之總計") * 0.6
      'End If
      'cmd = "SELECT 過帳交易資料表.客戶編號, 過帳交易資料表.單位, Sum(過帳交易資料表.重量) AS 重量之總計 "
      'cmd = cmd & "FROM 過帳交易資料表 "
      'cmd = cmd & "GROUP BY 過帳交易資料表.客戶編號, 過帳交易資料表.交易日期, 過帳交易資料表.單位 "
      'cmd = cmd & "HAVING (((過帳交易資料表.客戶編號) = '" & customId
      'cmd = cmd & "') AND ((過帳交易資料表.交易日期) Between #" & dateStart & "# And #" & dateStop
      'cmd = cmd & "#) AND ((過帳交易資料表.單位)='包'));"
      'dbDataBase1.RecordSource = cmd
      'dbDataBase1.Refresh
      'If dbDataBase1.Recordset.EOF = False Then
      '  UintBag = dbDataBase1.Recordset.Fields("重量之總計")
      'Else
      '  UintBag = 0
      'End If
      cmd = "SELECT 過帳交易資料表.交易日期, 魚貨資料表.魚貨名稱, 過帳交易資料表.重量, "
      cmd = cmd & "過帳交易資料表.單位, 過帳交易資料表.單價, 過帳交易資料表.合計 "
      cmd = cmd & "FROM 過帳交易資料表 INNER JOIN 魚貨資料表 ON 過帳交易資料表.魚貨代號 = "
      cmd = cmd & "魚貨資料表.魚貨代號 "
      cmd = cmd & "GROUP BY 過帳交易資料表.交易日期, 魚貨資料表.魚貨名稱, 過帳交易資料表.重量, "
      cmd = cmd & "過帳交易資料表.單位, 過帳交易資料表.單價, 過帳交易資料表.合計, 過帳交易資料表.識別碼, 過帳交易資料表.客戶編號 "
      cmd = cmd & "HAVING (((過帳交易資料表.客戶編號)= '" & customId
      cmd = cmd & "') AND ((過帳交易資料表.交易日期) Between #"
      cmd = cmd & dateStart & "# And #" & dateStop & "#)) "
      cmd = cmd & "ORDER BY 過帳交易資料表.識別碼;"
      dbDataBase1.RecordSource = cmd
      dbDataBase1.Refresh
      
      ' -----------------------------------------------------------------
      'ListBuf(ListNum) = " 寶 號 : " & customId & customName
      ListHeader(0) = " 寶 號 : " & StrAppendSpace(customId & " " & customName, 16, StrAppendLeft)
      ListHeader(1) = REPORT_NULL_SPAN
      ListHeader(2) = "日 期 : " & StrAppendSpace(DC2PC(CDate(dateStart)), 9, StrAppendLeft) & " 至 " & StrAppendSpace(DC2PC(CDate(dateStop)), 9, StrAppendLeft) & String(15, " ")
      ListNum = 0
      If prevSum = 0 Then
        ListBuf(ListNum) = "前日結欠:" & StrAppendSpace(CStr(prevSum), 36, StrAppendRight)
      Else
        ListBuf(ListNum) = "前日結欠:" & StrAppendSpace(Format(CStr(prevSum), "###,###,###"), 36, StrAppendRight)
      End If
      ListNum = ListNum + 1
      dealSum = 0
      Do Until dbDataBase1.Recordset.EOF
        fDate = dbDataBase1.Recordset.Fields("交易日期")
        fName = dbDataBase1.Recordset.Fields("魚貨名稱")
        fWeight = dbDataBase1.Recordset.Fields("重量")
        fUnit = dbDataBase1.Recordset.Fields("單位")
        fMoney = dbDataBase1.Recordset.Fields("單價")
        dMoney = dbDataBase1.Recordset.Fields("合計")
        dbDataBase1.Recordset.MoveNext
        If ListNum < 100 Then
          ListBuf(ListNum) = StrAppendSpace(DC2PC(CDate(fDate)), 9, StrAppendRight) & " " & StrAppendSpace(fName, 8, StrAppendLeft) & StrAppendSpace(StrFraction(CStr(fWeight), 2), 6, StrAppendRight) & " " & StrAppendSpace(fUnit, 4, StrAppendLeft) & " " & StrAppendSpace(StrFraction(CStr(fMoney), 1), 7, StrAppendRight) & " " & StrAppendSpace(Format(CStr(dMoney), "###,###,###"), 7, StrAppendRight)
          ListNum = ListNum + 1
        End If
        dealSum = dealSum + dMoney
      Loop
      If dealSum <> 0 Then
        ListBuf(ListNum) = "< 銷售合計 >"
        If UintKg <> 0 Then
          ListBuf(ListNum) = ListBuf(ListNum) & StrAppendSpace(StrFraction(CStr(UintKg), 2) & " Kg", 15, StrAppendRight)
        Else
          ListBuf(ListNum) = ListBuf(ListNum) & String(15, " ")
        End If
        If UintBag <> 0 Then
          ListBuf(ListNum) = ListBuf(ListNum) & StrAppendSpace(" " & UintBag & " Bag", 7, StrAppendLeft)
        Else
          ListBuf(ListNum) = ListBuf(ListNum) & String(7, " ")
        End If
        ListBuf(ListNum) = ListBuf(ListNum) & "   " & StrAppendSpace(Format(CStr(dealSum), "###,###,###"), 8, StrAppendRight)
        ListNum = ListNum + 1
        ListBuf(ListNum) = REPORT_ROW_SPAN
        ListNum = ListNum + 1
      End If
      
      cmd = "SELECT 過帳收款資料表.收款日期, 過帳收款資料表.收款金額 "
      cmd = cmd & "FROM 過帳收款資料表 "
      cmd = cmd & "GROUP BY 過帳收款資料表.客戶編號, 過帳收款資料表.收款日期, 過帳收款資料表.收款金額 "
      cmd = cmd & "HAVING (((過帳收款資料表.客戶編號)= '" & customId
      cmd = cmd & "') AND ((過帳收款資料表.收款日期) Between #"
      cmd = cmd & dateStart & "# And #" & dateStop & "#));"
      dbDataBase1.RecordSource = cmd
      dbDataBase1.Refresh
      receiveSum = 0
      Do Until dbDataBase1.Recordset.EOF
        fDate = dbDataBase1.Recordset.Fields("收款日期")
        dMoney = dbDataBase1.Recordset.Fields("收款金額")
        dbDataBase1.Recordset.MoveNext
        If ListNum < 100 Then
          ListBuf(ListNum) = StrAppendSpace(DC2PC(CDate(fDate)), 9, StrAppendRight) & " 收 款..  " & StrAppendSpace(Format(CStr(dMoney), "###,###,###"), 26, StrAppendRight)
          ListNum = ListNum + 1
        End If
        receiveSum = receiveSum + dMoney
      Loop
      If receiveSum <> 0 Then
        ListBuf(ListNum) = "< 收款合計 >         " & StrAppendSpace(Format(CStr(receiveSum), "###,###,###"), 24, StrAppendRight)
        ListNum = ListNum + 1
        ListBuf(ListNum) = REPORT_ROW_SPAN
        ListNum = ListNum + 1
      End If
      ListBuf(ListNum) = REPORT_NULL_SPAN
      ListNum = ListNum + 1
      ListBuf(ListNum) = "至 " & StrAppendSpace(DC2PC(CDate(dateStop)), 9, StrAppendRight) & " 止合計結欠:  "
      If finalSum <> 0 Then
        ListBuf(ListNum) = ListBuf(ListNum) & StrAppendSpace(Format(CStr(finalSum), "###,###,###"), 15, StrAppendRight) & String(4, " ")
      Else
        ListBuf(ListNum) = ListBuf(ListNum) & StrAppendSpace("0", 15, StrAppendRight) & String(4, " ")
      End If
      ListNum = ListNum + 1
      NumSmallPage = Fix((ListNum + NUM_LIST_PRE_PAGE - 1) / NUM_LIST_PRE_PAGE)
      For i = 0 To NumSmallPage - 1
        If toggleList = 0 Then
          ''CurrentList = ListNum Mod NUM_LIST_PRE_PAGE
          If i = NumSmallPage - 1 Then
            CurrentList = ListNum Mod NUM_LIST_PRE_PAGE
          Else
            CurrentList = NUM_LIST_PRE_PAGE
          End If
          If CurrentList = 0 Then
            CurrentList = NUM_LIST_PRE_PAGE
          End If
          SmallPageBuf(0) = ListHeader(0) & StrAppendSpace("# " & CStr(i + 1), 20, StrAppendRight)
          For j = 1 To 2
            SmallPageBuf(j) = ListHeader(j)
          Next
          For j = 0 To CurrentList - 1
            SmallPageBuf(j + 3) = ListBuf(i * NUM_LIST_PRE_PAGE + j)
          Next
          '''''''''''(聯)
          'If i = NumSmallPage - 1 Then
            SmallPageBuf(CurrentList + 3) = REPORT_NULL_SPAN
            SmallPageBuf(CurrentList + 3 + 1) = REPORT_FINAL_SPAN
            CurrentList = CurrentList + 2
          'End If
          'For j = CurrentList To NUM_LIST_PRE_PAGE - 1
          For j = CurrentList To NUM_LIST_PRE_PAGE + 7
            SmallPageBuf(j + 3) = REPORT_NULL_SPAN
          Next
          If TotalPage < 100 Then
            If (SmallPageCnt Mod 6) = 0 Then
              TotalPage = TotalPage + 1
              PageBuf(TotalPage) = ""
            End If
          End If
          SmallPageCnt = SmallPageCnt + 1
          toggleList = 1
        Else
          'CurrentList = ListNum Mod NUM_LIST_PRE_PAGE
          If i = NumSmallPage - 1 Then
            CurrentList = ListNum Mod NUM_LIST_PRE_PAGE
          Else
            CurrentList = NUM_LIST_PRE_PAGE
          End If
          If CurrentList = 0 Then
            CurrentList = NUM_LIST_PRE_PAGE
          End If
          SmallPageBuf(0) = SmallPageBuf(0) & REPORT_COLUMN_SPAN & ListHeader(0) & StrAppendSpace("# " & CStr(i + 1), 20, StrAppendRight)
          For j = 1 To 2
            SmallPageBuf(j) = SmallPageBuf(j) & REPORT_COLUMN_SPAN & ListHeader(j)
          Next
          For j = 0 To CurrentList - 1
            SmallPageBuf(j + 3) = SmallPageBuf(j + 3) & REPORT_COLUMN_SPAN & ListBuf(i * NUM_LIST_PRE_PAGE + j)
          Next
          '''''''''''(聯)
          'If i = NumSmallPage - 1 Then
            SmallPageBuf(CurrentList + 3) = SmallPageBuf(CurrentList + 3) & REPORT_COLUMN_SPAN & REPORT_NULL_SPAN
            SmallPageBuf(CurrentList + 3 + 1) = SmallPageBuf(CurrentList + 3 + 1) & REPORT_COLUMN_SPAN & REPORT_FINAL_SPAN
            CurrentList = CurrentList + 2
          'End If
          'For j = CurrentList To NUM_LIST_PRE_PAGE - 1
          For j = CurrentList To NUM_LIST_PRE_PAGE + 8
            SmallPageBuf(j + 3) = SmallPageBuf(j + 3) & REPORT_COLUMN_SPAN & REPORT_NULL_SPAN
          Next
          If TotalPage < 100 Then
            If (SmallPageCnt Mod 6) = 0 Then
              TotalPage = TotalPage + 1
              PageBuf(TotalPage) = ""
            End If
            If (((SmallPageCnt Mod 6) = 4) Or ((SmallPageCnt Mod 6) = 5)) Then
              For j = 0 To NUM_LIST_PRE_PAGE + 4
                PageBuf(TotalPage) = PageBuf(TotalPage) & SmallPageBuf(j) & vbCrLf
              Next
            Else
              For j = 0 To 25 '27 'NUM_LIST_PRE_PAGE + 5
                PageBuf(TotalPage) = PageBuf(TotalPage) & SmallPageBuf(j) & vbCrLf
              Next
            End If
          End If
          SmallPageCnt = SmallPageCnt + 1
          toggleList = 0
        End If
      Next
    End If 'checked?
  Next ' loop k
  If toggleList = 1 Then
    If TotalPage < 100 Then
      If (SmallPageCnt Mod 6) = 0 Then
        TotalPage = TotalPage + 1
        PageBuf(TotalPage) = ""
      End If
      If (((SmallPageCnt Mod 6) = 4) Or ((SmallPageCnt Mod 6) = 5)) Then
        For j = 0 To NUM_LIST_PRE_PAGE + 4
          PageBuf(TotalPage) = PageBuf(TotalPage) & SmallPageBuf(j) & vbCrLf
        Next
      Else
        For j = 0 To 25 ' 27 'NUM_LIST_PRE_PAGE + 5
          PageBuf(TotalPage) = PageBuf(TotalPage) & SmallPageBuf(j) & vbCrLf
        Next
      End If
    End If
  End If
End Sub

Private Sub cmdPage_Click(Index As Integer)

  If Index = 0 Then
    CurrentPage = 1
  ElseIf Index = 1 Then
    CurrentPage = CurrentPage - 1
  ElseIf Index = 2 Then
    CurrentPage = CurrentPage + 1
  ElseIf Index = 3 Then
    CurrentPage = TotalPage
  End If
  If CurrentPage = 1 Then
    cmdPage(0).Enabled = False
    cmdPage(1).Enabled = False
  Else
    cmdPage(0).Enabled = True
    cmdPage(1).Enabled = True
  End If
  If CurrentPage = TotalPage Then
    cmdPage(2).Enabled = False
    cmdPage(3).Enabled = False
  Else
    cmdPage(2).Enabled = True
    cmdPage(3).Enabled = True
  End If
  Preview.Text = PageBuf(CurrentPage)
  lblPage = CurrentPage & " / " & TotalPage
End Sub


Private Sub Page2List(PageIndex As Integer)
  Dim Text1 As String
  Dim Text2 As String
  Dim Text3 As String
  Dim PrevPos As Integer
  Dim CurrPos As Integer
  Dim LenText As Integer
  Dim LenText2 As Integer
  Dim AsciiSize As Integer
  Dim i As Byte
  
  NumList = 0
  CurrPos = 1
  PrevPos = 1
  Text1 = PageBuf(PageIndex)
  LenText = Len(Text1)
  LenText2 = LenB(StrConv(Text1, vbFromUnicode))
  While CurrPos < LenText
    CurrPos = InStr(CurrPos, Text1, vbCrLf)
    If CurrPos <> 0 Then
      If CurrPos > PrevPos Then
        ListBufPage(NumList) = Mid(Text1, PrevPos, CurrPos - PrevPos)
        NumList = NumList + 1
      End If
      CurrPos = CurrPos + 2
      PrevPos = CurrPos
    Else
      LenText = 0
    End If
  Wend
  
  NumListBuf14 = 0
  '寶號
  For i = 1 To NumList
    CurrPos = 1
    PrevPos = 1
    Text1 = ListBufPage(i - 1)
    Text2 = ""
    LenText = Len(Text1)
    While CurrPos < LenText
      CurrPos = InStr(CurrPos, Text1, "寶 號 : ")
      If CurrPos <> 0 Then
        CurrPos = CurrPos + 6
        Text2 = Text2 & Mid(Text1, PrevPos, CurrPos - PrevPos)
        Text3 = Mid(Text1, CurrPos, 12)
        AsciiSize = LenB(StrConv(Text3, vbFromUnicode))
        CurrPos = CurrPos + 12
        ListBuf14pt(NumListBuf14) = Text3
        Text2 = Text2 & String(AsciiSize, " ")
        PrevPos = CurrPos
        ListBuf14ptLine(NumListBuf14) = i
        If CurrPos > 50 Then
          ListBuf14ptType(NumListBuf14) = 2
        Else
          ListBuf14ptType(NumListBuf14) = 1
        End If
        NumListBuf14 = NumListBuf14 + 1
      Else
        Text2 = Text2 & Mid(Text1, PrevPos, LenText - PrevPos + 1)
        LenText = 0
      End If
    Wend
    ListBufPage(i - 1) = Text2
  Next
  For i = 1 To NumList
    CurrPos = 1
    PrevPos = 1
    Text1 = ListBufPage(i - 1)
    Text2 = ""
    LenText = Len(Text1)
    While CurrPos < LenText
      CurrPos = InStr(CurrPos, Text1, "止合計結欠:")
      If CurrPos <> 0 Then
        CurrPos = CurrPos + 15
        Text2 = Text2 & Mid(Text1, PrevPos, CurrPos - PrevPos)
        Text3 = Mid(Text1, CurrPos, 8)
        AsciiSize = LenB(StrConv(Text3, vbFromUnicode))
        CurrPos = CurrPos + 8
        ListBuf14pt(NumListBuf14) = Text3
        Text2 = Text2 & String(AsciiSize, " ")
        PrevPos = CurrPos
        ListBuf14ptLine(NumListBuf14) = i
        If CurrPos > 50 Then
          ListBuf14ptType(NumListBuf14) = 4
        Else
          ListBuf14ptType(NumListBuf14) = 3
        End If
        NumListBuf14 = NumListBuf14 + 1
      Else
        Text2 = Text2 & Mid(Text1, PrevPos, LenText - PrevPos + 1)
        LenText = 0
      End If
    Wend
    ListBufPage(i - 1) = Text2
  Next
End Sub
Private Sub cmdPrint_Click()
  Dim i, j, k As Integer
  
  On Error GoTo ErrHandlerPrintReceipt
  Dim nTextLength      As Long
  Dim nNextCharPos     As Long
  Dim PrintData As String
  
  PrintDialog.CancelError = True
  PrintDialog.Flags = cdlPDReturnDC + cdlPDPageNums + cdlPDNoSelection + cdlPDDisablePrintToFile + cdlPDAllPages + cdlPDCollate
  PrintDialog.ShowPrinter
  Call SetPrintPage(40)
  Printer.FontName = "細明體"
  'Printer.FontBold = True
  Printer.Print Space(1)
  Printer.ScaleMode = vbTwips
  For i = 1 To TotalPage
    Page2List (i)
    For j = 1 To NumList
      Printer.FontSize = 11
      Printer.CurrentX = PRINTER_X_OFFSET
      Printer.CurrentY = PRINTER_Y_OFFSET + j * 220
      Printer.Print ListBufPage(j - 1)
      For k = 1 To NumListBuf14
        If j = ListBuf14ptLine(k - 1) Then
          Printer.FontSize = 14
          Printer.CurrentY = (ListBuf14ptLine(k - 1) * 220) - 50
          If ListBuf14ptType(k - 1) = 1 Then
            Printer.CurrentX = PRINTER_X_OFFSET + PRINTER_MONEY_X1_OFFSET
          ElseIf ListBuf14ptType(k - 1) = 2 Then
            Printer.CurrentX = PRINTER_X_OFFSET + PRINTER_MONEY_X1_OFFSET + PRINTER_X2_OFFSET
          ElseIf ListBuf14ptType(k - 1) = 3 Then
            Printer.CurrentX = PRINTER_X_OFFSET + PRINTER_MONEY_X2_OFFSET
          ElseIf ListBuf14ptType(k - 1) = 4 Then
            Printer.CurrentX = PRINTER_X_OFFSET + PRINTER_MONEY_X2_OFFSET + PRINTER_X2_OFFSET
          Else
          End If
          Printer.Print ListBuf14pt(k - 1)
        End If
      Next
    Next
    'For j = 1 To NumListBuf14
    '  'If ListBuf14ptLine(j - 1) > 5 Then
    '  '  Printer.CurrentY = 50 + (ListBuf14ptLine(j - 1) + 1) * 200
    '  'Else
    '    Printer.CurrentY = 50 + ListBuf14ptLine(j - 1) * 200
    '  'End If
    '  If ListBuf14ptType(j - 1) = 1 Then
    '    Printer.CurrentX = PRINTER_X_OFFSET + PRINTER_MONEY_X1_OFFSET
    '  ElseIf ListBuf14ptType(j - 1) = 2 Then
    '    Printer.CurrentX = PRINTER_X_OFFSET + PRINTER_MONEY_X1_OFFSET + PRINTER_X2_OFFSET
    '  ElseIf ListBuf14ptType(j - 1) = 3 Then
    '    Printer.CurrentX = PRINTER_X_OFFSET + PRINTER_MONEY_X2_OFFSET
    '  ElseIf ListBuf14ptType(j - 1) = 4 Then
    '    Printer.CurrentX = PRINTER_X_OFFSET + PRINTER_MONEY_X2_OFFSET + PRINTER_X2_OFFSET
    '  Else
    '  End If
    '  Printer.Print ListBuf14pt(j - 1)
    'Next
    If i <> TotalPage Then
      Printer.NewPage
    End If
  Next
  Printer.EndDoc
  Screen.MousePointer = 0
ErrHandlerPrintReceipt:
End Sub

'Private Sub DateChk_Click()
'  If DateChk.Value = 1 Then
'    checkDate(0).Enabled = True
'    checkDate(1).Enabled = True
'  Else
'    checkDate(0).Enabled = False
'    checkDate(1).Enabled = False
'  End If
'End Sub

Private Sub Form_Load()
  TimerMode = 0
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  UpdateFlag = 0
  TotalList = 0
  StartId = "000"
  StopId = "999"
  TotalPage = 0
  CurrentPage = 0
  checkDate(0) = CalendarValue
  checkDate(1) = CalendarValue
  StartDate = CalendarValue
  StopDate = CalendarValue
  dbDataBase1.ConnectionString = database_string
  dbDataBase2.ConnectionString = database_string
  dbDataBase3.ConnectionString = database_string
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub

Private Sub listReport_ItemCheck(ByVal item As MSComctlLib.ListItem)
  UpdateFlag = 1
End Sub

Private Sub TabStrip1_Click()
  Dim i As Integer
  If TabStrip1.SelectedItem.Index = 1 Then
    Frame2.Visible = True
    Frame3.Visible = False
    Frame4.Visible = True
    listReport.Visible = True
    Preview.Visible = False
    cmdPrint.Visible = False
    lblPage.Visible = False
    For i = 0 To 3
      cmdPage(i).Visible = False
    Next
  Else
    Frame2.Visible = False
    Frame3.Visible = True
    Frame4.Visible = False
    listReport.Visible = False
    Preview.Visible = True
    cmdPrint.Enabled = False
    cmdPrint.Visible = True
    lblPage.Visible = True
    For i = 0 To 3
      cmdPage(i).Enabled = False
      cmdPage(i).Visible = True
    Next
    Call Print_Receipt
  End If
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  If TimerMode = 1 Then
    Call AutoFun
  ElseIf TimerMode = 2 Then
    Call AddFun
  ElseIf TimerMode = 3 And UpdateFlag = 1 Then
  'ElseIf TimerMode = 3 Then
    Call ReceiptFun
    If TotalPage <> 0 Then
      cmdPrint.Enabled = True
      If TotalPage <> 1 Then
        cmdPage(2).Enabled = True
        cmdPage(3).Enabled = True
      End If
      CurrentPage = 1
      Preview.Text = PageBuf(CurrentPage)
      lblPage = CurrentPage & " / " & TotalPage
    End If
  Else
    If TotalPage <> 0 Then
      If TotalPage <> 1 Then
        cmdPage(2).Enabled = True
        cmdPage(3).Enabled = True
      End If
      Preview.Text = PageBuf(CurrentPage)
      lblPage = CurrentPage & " / " & TotalPage
    End If
  End If
  
  Call ProcDone
End Sub
