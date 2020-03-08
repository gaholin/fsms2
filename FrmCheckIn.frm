VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmCheckIn 
   BorderStyle     =   1  '單線固定
   Caption         =   "交易/收款過帳作業"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "FrmCheckIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "FrmCheckIn.frx":0E42
   ScaleHeight     =   7785
   ScaleWidth      =   11910
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   8640
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame CheckWindow 
      Height          =   2415
      Left            =   3240
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Label3 
         Caption         =   "過帳中請稍後"
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
         Left            =   1200
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc dbDataBase3 
      Height          =   330
      Left            =   4800
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   2640
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSDataGridLib.DataGrid listDataBase 
      Height          =   4095
      Left            =   600
      TabIndex        =   6
      Top             =   3360
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc dbDataBase1 
      Height          =   330
      Left            =   600
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.Frame Frame1 
      Caption         =   "設定範圍"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   600
      TabIndex        =   7
      Top             =   840
      Width           =   7815
      Begin VB.OptionButton optSort 
         Caption         =   "日期優先"
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
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   1560
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optSort 
         Caption         =   "客戶編號"
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
         Index           =   0
         Left            =   2880
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "列印"
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
         Left            =   6120
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton CheckInButton 
         Caption         =   "過帳"
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
         Left            =   6120
         TabIndex        =   5
         Top             =   840
         Width           =   1335
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
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   1
         Top             =   360
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
         Index           =   1
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin MSComCtl2.DTPicker checkDate 
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   3
         Top             =   960
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
         CurrentDate     =   36327
      End
      Begin MSComCtl2.DTPicker checkDate 
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   4
         Top             =   960
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
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "排列方式"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   975
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
         Left            =   3240
         TabIndex        =   9
         Top             =   1080
         Width           =   375
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
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3255
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   12938
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "交易過帳作業"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "收款過帳作業"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "已過帳－交易明細表"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "已過帳－收款明細表"
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
End
Attribute VB_Name = "FrmCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private start_cid As String
Private end_cid As String
Private cid_vaild As Boolean
Private date_vaild As Boolean
Private cid_date_criteria As String
Private sum_criteria As String
Private total_list As Integer
Private list_report(4096) As String
'Deal Start Span
Private Const REPORT_START_SPAN1 = " "
'Receive Start Span
Private Const REPORT_START_SPAN2 = "     "
Private Const PREORT_RECEIVE_HEADER1 = "============================================="
Private Const PREORT_RECEIVE_HEADER2 = "|   客      戶   |  日    期  | 收 款 金 額 |"
Private Const PREORT_RECEIVE_HEADER3 = "============================================="
Private Const REPORT_RECEIVE_COLUMN_SPAN = "        "
Private Const REPORT_RECEIVE_ROW_SPAN = "---------------------------------------------"
Private Const PREORT_DEAL_HEADER1 = "============================================================================================================"
Private Const PREORT_DEAL_HEADER2 = "|  客     戶  | 交易日期 | 魚 名 代 號 | 重  量 | 單 價 |稅 率| 傳 票 |籠 子|持分| 金   額 |其 他| 合   計 |"
Private Const PREORT_DEAL_HEADER3 = "============================================================================================================"
Private Const PREORT_DEAL_ROW_SPAN = "------------------------------------------------------------------------------------------------------------"

Private Sub set_listDeal()
'客戶編號 , 客戶姓名, 交易日期
'識別碼 , 魚貨代號, 魚貨名稱, 單位,
'重量 , 單價, 稅別, 稅率,
'傳票 , 籠子, 持分, 金額,
'其他, 合計 "
  On Error GoTo ErrListDeal
  listDataBase.Columns.item(0).Width = 1000 '客戶編號
  listDataBase.Columns.item(1).Width = 1000 '客戶姓名
  listDataBase.Columns.item(2).Width = 1000 '交易日期
  listDataBase.Columns.item(3).Width = 12 '識別碼
  listDataBase.Columns.item(4).Width = 1000 '魚貨代號
  listDataBase.Columns.item(5).Width = 1000 '魚貨名稱
  listDataBase.Columns.item(6).Width = 12 '稅率
  listDataBase.Columns.item(7).Width = 800 '重量
  listDataBase.Columns.item(8).Width = 800 '單價
  listDataBase.Columns.item(9).Width = 800 '稅別
  listDataBase.Columns.item(10).Width = 12  '稅率
  listDataBase.Columns.item(11).Width = 12  '傳票
  listDataBase.Columns.item(12).Width = 12  '籠子
  listDataBase.Columns.item(13).Width = 12  '持分
  listDataBase.Columns.item(14).Width = 1000 '金額
  listDataBase.Columns.item(15).Width = 700 '其他
  listDataBase.Columns.item(16).Width = 1000 '合計
ErrListDeal:
End Sub
Private Sub set_listReceive()
  On Error GoTo ErrListReceive
  listDataBase.Columns.item(0).Width = 1600
  listDataBase.Columns.item(1).Width = 2000
  listDataBase.Columns.item(2).Width = 2000
  listDataBase.Columns.item(3).Width = 12
  listDataBase.Columns.item(4).Width = 2000
ErrListReceive:
End Sub
Private Sub update_disp()
  Dim tmp As String
  Dim item As String
  Dim firstDate As String
  Dim listCnt As Integer
  Dim listDate As Date
  Dim SumWeight As Double
  Dim TotalWeight As Double
  Dim SumSummons As Double
  Dim TotalSummons As Double
  Dim SumBasket As Long
  Dim TotalBasket As Long
  Dim SumMoney As Long
  Dim TotalMoney As Long
  Dim SumOther As Long
  Dim TotalOther As Long
  Dim SumSum As Long
  Dim TotalSum As Long
  Dim SumCount As Integer
  Dim TotalCount As Integer
  'On Error GoTo ErrorProc
  
  Call get_criteria
  Call get_sum_criteria
  If CheckMode = 0 Then
    Call dispDeal
    Call dispChkInDealSum
    Call set_listDeal
  ElseIf CheckMode = 1 Then
    Call dispReceive
    Call dispChkInReceiveSum
    Call set_listReceive
  ElseIf CheckMode = 2 Then
    Call dispChkInDeal
    Call set_listDeal
  Else
    Call dispChkInReceive
    Call set_listReceive
  End If
  total_list = 0
  If dbDataBase1.Recordset.BOF = True And dbDataBase1.Recordset.EOF = True Then
    CheckInButton.Enabled = False
    cmdPrint.Enabled = False
  Else
    CheckInButton.Enabled = True
    cmdPrint.Enabled = True
    dbDataBase1.Recordset.MoveFirst
    ' Deal Report
    If CheckMode = 0 Or CheckMode = 2 Then
      firstDate = dbDataBase1.Recordset.Fields("交易日期")
      SumMoney = 0
      TotalMoney = 0
      SumWeight = 0
      TotalWeight = 0
      SumSummons = 0
      TotalSummons = 0
      SumBasket = 0
      TotalBasket = 0
      SumOther = 0
      TotalOther = 0
      SumSum = 0
      TotalSum = 0
      SumCount = 0
      TotalCount = 0
      listCnt = 0
      Do Until dbDataBase1.Recordset.EOF
        item = dbDataBase1.Recordset.Fields("客戶編號")
        tmp = StrAppendSpace(item, 4, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("客戶姓名")
        tmp = tmp & StrAppendSpace(item, 9, StrAppendLeft) & " "
        item = dbDataBase1.Recordset.Fields("交易日期")
        listDate = CDate(item)
        'tmp = tmp & " " & String(3 - Len(Year(listDate) - 1911), " ") & Year(listDate) - 1911
        tmp = tmp & " " & String(3 - Len(Year(listDate) - 1911), " ") & Year(listDate) - 1911
        tmp = tmp & "/" & String(2 - Len(Month(listDate)), "0") & Month(listDate)
        tmp = tmp & "/" & String(2 - Len(Day(listDate)), "0") & Day(listDate) & " "
        
        If (item <> firstDate) And (optSort(1).value = True) Then
          firstDate = item
          If total_list < 1023 Then
            If (listCnt Mod 5) <> 0 Then
              list_report(total_list) = PREORT_DEAL_ROW_SPAN
              total_list = total_list + 1
            End If
          item = "  小 計 :        共 " & SumCount & " 筆"
            list_report(total_list) = StrAppendSpace(item, 40, StrAppendLeft) & StrAppendSpace(StrFraction(CStr(SumWeight), 2), 8, StrAppendRight) & " "
            list_report(total_list) = list_report(total_list) & StrAppendSpace(StrFraction(CStr(SumSummons), 1), 21, StrAppendRight) & " "
            list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(SumBasket, "#,####"), 5, StrAppendRight) & " "
            list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(SumMoney, "#,###,###"), 14, StrAppendRight) & " "
            list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(SumOther, "##,###"), 5, StrAppendRight) & " "
            list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(SumSum, "#,###,###"), 9, StrAppendRight)
            total_list = total_list + 1
            list_report(total_list) = PREORT_DEAL_ROW_SPAN
            total_list = total_list + 1
            listCnt = 0
          End If
          SumWeight = 0
          SumSummons = 0
          SumBasket = 0
          SumMoney = 0
          SumOther = 0
          SumSum = 0
          SumCount = 0
        End If
        item = dbDataBase1.Recordset.Fields("魚貨代號")
        tmp = tmp & StrAppendSpace(item, 4, StrAppendLeft) & " "
        item = dbDataBase1.Recordset.Fields("魚貨名稱")
        tmp = tmp & StrAppendSpace(item, 8, StrAppendLeft) & " "
        item = dbDataBase1.Recordset.Fields("重量")
        SumWeight = SumWeight + CDbl(item)
        TotalWeight = TotalWeight + CDbl(item)
        tmp = tmp & StrAppendSpace(StrFraction(item, 2), 8, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("單價")
        tmp = tmp & StrAppendSpace(StrFraction(item, 1), 7, StrAppendRight) & " "
        item = CStr(CDbl(dbDataBase1.Recordset.Fields("稅率")) - 1)
        
        tmp = tmp & StrAppendSpace(StrFraction(item, 3), 5, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("傳票")
        SumSummons = SumSummons + CDbl(item)
        TotalSummons = TotalSummons + CLng(item)
        tmp = tmp & StrAppendSpace(StrFraction(item, 1), 7, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("籠子")
        SumBasket = SumBasket + CLng(item)
        TotalBasket = TotalBasket + CLng(item)
        tmp = tmp & StrAppendSpace(CStr(Format(CLng(item), "#,###")), 5, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("持分")
        tmp = tmp & StrAppendSpace(item, 4, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("金額")
        SumMoney = SumMoney + CLng(item)
        TotalMoney = TotalMoney + CLng(item)
        tmp = tmp & StrAppendSpace(CStr(Format(CLng(item), "#,###,###")), 9, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("其他")
        SumOther = SumOther + CLng(item)
        TotalOther = TotalOther + CLng(item)
        tmp = tmp & StrAppendSpace(CStr(Format(CLng(item), "##,###")), 5, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("合計")
        SumSum = SumSum + CLng(item)
        TotalSum = TotalSum + CLng(item)
        tmp = tmp & StrAppendSpace(CStr(Format(CLng(item), "#,###,###")), 9, StrAppendRight)
        SumCount = SumCount + 1
        TotalCount = TotalCount + 1
        If total_list < 1023 Then
          list_report(total_list) = tmp
          total_list = total_list + 1
          listCnt = listCnt + 1
          If (listCnt Mod 5) = 0 Then
            list_report(total_list) = PREORT_DEAL_ROW_SPAN
            total_list = total_list + 1
          End If
        End If
        dbDataBase1.Recordset.MoveNext
      Loop
      If total_list < 4094 Then
        If (listCnt Mod 5) <> 0 Then
          list_report(total_list) = PREORT_DEAL_ROW_SPAN
          total_list = total_list + 1
        End If
        If optSort(1).value = True Then
          item = "  小 計 :        共 " & SumCount & " 筆"
          list_report(total_list) = StrAppendSpace(item, 40, StrAppendLeft) & "" & StrAppendSpace(StrFraction(CStr(SumWeight), 2), 8, StrAppendRight) & " "
          list_report(total_list) = list_report(total_list) & StrAppendSpace(StrFraction(CStr(SumSummons), 1), 21, StrAppendRight) & " "
          list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(SumBasket, "#,####"), 5, StrAppendRight) & " "
          list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(SumMoney, "#,###,###"), 14, StrAppendRight) & " "
          list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(SumOther, "##,###"), 5, StrAppendRight) & " "
          list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(SumSum, "#,###,###"), 9, StrAppendRight)
          total_list = total_list + 1
          list_report(total_list) = PREORT_DEAL_ROW_SPAN
          total_list = total_list + 1
        End If
        item = "  合 計 :        共 " & TotalCount & " 筆"
        list_report(total_list) = StrAppendSpace(item, 40, StrAppendLeft) & StrAppendSpace(StrFraction(CStr(TotalWeight), 2), 8, StrAppendRight) & " "
        list_report(total_list) = list_report(total_list) & StrAppendSpace(StrFraction(CStr(TotalSummons), 1), 21, StrAppendRight) & " "
        list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(TotalBasket, "#,####"), 5, StrAppendRight) & " "
        list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(TotalMoney, "#,###,###"), 14, StrAppendRight) & " "
        list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(TotalOther, "##,###"), 5, StrAppendRight) & " "
        list_report(total_list) = list_report(total_list) & StrAppendSpace(Format(TotalSum, "#,###,###"), 9, StrAppendRight)
        total_list = total_list + 1
      End If
    ' Receive Report
    ElseIf CheckMode = 1 Or CheckMode = 3 Then
      firstDate = dbDataBase1.Recordset.Fields("收款日期")
      SumMoney = 0
      TotalMoney = 0
      listCnt = 0
      Do Until dbDataBase1.Recordset.EOF
        item = dbDataBase1.Recordset.Fields("客戶編號")
        tmp = StrAppendSpace(item, 5, StrAppendRight) & "  "
        item = dbDataBase1.Recordset.Fields("客戶姓名")
        tmp = tmp & StrAppendSpace(item, 10, StrAppendLeft) & "  "
        item = dbDataBase1.Recordset.Fields("收款日期")
        listDate = CDate(item)
        'tmp = tmp & " " & String(3 - Len(Year(listDate) - 1911), " ") & Year(listDate) - 1911
        tmp = tmp & " " & String(3 - Len(Year(listDate) - 1911), " ") & Year(listDate) - 1911
        tmp = tmp & "/" & String(2 - Len(Month(listDate)), "0") & Month(listDate)
        tmp = tmp & "/" & String(2 - Len(Day(listDate)), "0") & Day(listDate) & " "
        If (item <> firstDate) And (optSort(1).value = True) Then
          firstDate = item
          If total_list < 1023 Then
            If (listCnt Mod 5) <> 0 Then
              list_report(total_list) = REPORT_RECEIVE_ROW_SPAN
              total_list = total_list + 1
            End If
            list_report(total_list) = String(22, " ") & "小 計 : " & StrAppendSpace(Format(SumMoney, "###,###,###"), 13, StrAppendRight) & "  "
            total_list = total_list + 1
            list_report(total_list) = REPORT_RECEIVE_ROW_SPAN
            total_list = total_list + 1
            listCnt = 0
          End If
          SumMoney = 0
        End If
        item = Format(dbDataBase1.Recordset.Fields("收款金額"), "###,###,###")
        SumMoney = SumMoney + CLng(item)
        tmp = tmp & StrAppendSpace(item, 13, StrAppendRight)
        If total_list < 1023 Then
          list_report(total_list) = tmp & "  "
          total_list = total_list + 1
          listCnt = listCnt + 1
          If (listCnt Mod 5) = 0 Then
            list_report(total_list) = REPORT_RECEIVE_ROW_SPAN
            total_list = total_list + 1
          End If
          TotalMoney = TotalMoney + CLng(dbDataBase1.Recordset.Fields("收款金額"))
        End If
        dbDataBase1.Recordset.MoveNext
      Loop
      If total_list < 4094 Then
        If (listCnt Mod 5) <> 0 Then
          list_report(total_list) = REPORT_RECEIVE_ROW_SPAN
          total_list = total_list + 1
        End If
        If optSort(1).value = True Then
          list_report(total_list) = String(22, " ") & "小 計 : " & StrAppendSpace(Format(SumMoney, "###,###,###"), 13, StrAppendRight) & "  "
          total_list = total_list + 1
          list_report(total_list) = REPORT_RECEIVE_ROW_SPAN
          total_list = total_list + 1
        End If
        list_report(total_list) = String(22, " ") & "合 計 : " & StrAppendSpace(Format(TotalMoney, "###,###,###"), 13, StrAppendRight) & "  "
        total_list = total_list + 1
      End If
      If total_list > 210 Then
        total_list = total_list + 1
      End If
    End If
    dbDataBase1.Recordset.MoveFirst
  End If
ErrorProc:
End Sub
Private Sub get_criteria()
  Dim cmd As String
  cmd = ""
  If CheckMode = 0 Then
    cmd = cmd & "WHERE (((交易資料表.客戶編號) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((交易資料表.交易日期) "
      cmd = cmd & "Between #" & checkDate(0) & "# And #" & checkDate(1) & "#) "
    'End If
    cmd = cmd & ") "
  ElseIf CheckMode = 1 Then
    cmd = cmd & "WHERE (((收款資料表.客戶編號) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((收款資料表.收款日期) "
      cmd = cmd & "Between #" & checkDate(0) & "# And #" & checkDate(1) & "#) "
    'End If
    cmd = cmd & ") "
  ElseIf CheckMode = 2 Then
    cmd = cmd & "WHERE (((過帳交易資料表.客戶編號) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((過帳交易資料表.交易日期) "
      cmd = cmd & "Between #" & checkDate(0) & "# And #" & checkDate(1) & "#) "
    'End If
    cmd = cmd & ") "
  ElseIf CheckMode = 3 Then
    cmd = cmd & "WHERE (((過帳收款資料表.客戶編號) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((過帳收款資料表.收款日期) "
      cmd = cmd & "Between #" & checkDate(0) & "# And #" & checkDate(1) & "#) "
    'End If
    cmd = cmd & ") "
  End If
  cid_date_criteria = cmd
End Sub

Private Sub get_sum_criteria()
  Dim cmd As String
  cmd = ""
  If CheckMode = 0 Then
    cmd = cmd & "HAVING (((交易資料表.客戶編號) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((交易資料表.交易日期) "
      cmd = cmd & "Between #" & checkDate(0) & "# And #" & checkDate(1) & "#) "
    'End If
    cmd = cmd & ") "
  ElseIf CheckMode = 1 Then
    cmd = cmd & "HAVING (((收款資料表.客戶編號) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((收款資料表.收款日期) "
      cmd = cmd & "Between #" & checkDate(0) & "# And #" & checkDate(1) & "#) "
    'End If
    cmd = cmd & ") "
  End If
  sum_criteria = cmd
End Sub

Private Sub dispDeal()
  Dim cmd As String
  cmd = "SELECT 交易資料表.客戶編號, 客戶資料表.客戶姓名, 交易資料表.交易日期, "
  cmd = cmd & "交易資料表.識別碼, 交易資料表.魚貨代號, 魚貨資料表.魚貨名稱, 交易資料表.單位, "
  cmd = cmd & "交易資料表.重量, 交易資料表.單價, 交易資料表.稅別, 稅率資料表.稅率, "
  cmd = cmd & "稅率資料表.傳票, 稅率資料表.籠子, 交易資料表.持分, 交易資料表.金額, "
  cmd = cmd & "交易資料表.其他, 交易資料表.合計 "
  cmd = cmd & "FROM (稅率資料表 INNER JOIN (魚貨資料表 INNER JOIN 交易資料表 ON "
  cmd = cmd & "(魚貨資料表.魚貨代號 = 交易資料表.魚貨代號) AND (魚貨資料表.魚貨代號 "
  cmd = cmd & "= 交易資料表.魚貨代號)) ON (稅率資料表.識別碼 = 交易資料表.稅別) AND "
  cmd = cmd & "(稅率資料表.識別碼 = 交易資料表.稅別)) INNER JOIN 客戶資料表 ON "
  cmd = cmd & "(交易資料表.客戶編號 = 客戶資料表.客戶編號) AND (交易資料表.客戶編號 "
  cmd = cmd & "= 客戶資料表.客戶編號) "
  cmd = cmd & cid_date_criteria
  'cmd = cmd & "ORDER BY 交易資料表.客戶編號, 交易資料表.交易日期, 交易資料表.識別碼;"
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY 交易資料表.客戶編號, 交易資料表.交易日期, 交易資料表.識別碼;"
  Else
  cmd = cmd & "ORDER BY 交易資料表.交易日期, 交易資料表.識別碼;"
  End If
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh
  Call set_listDeal
End Sub

Private Sub dispReceive()
  Dim cmd As String
  cmd = "SELECT 收款資料表.客戶編號, 客戶資料表.客戶姓名, 收款資料表.收款日期, "
  cmd = cmd & "收款資料表.識別碼, 收款資料表.收款金額 "
  cmd = cmd & "FROM 客戶資料表 INNER JOIN 收款資料表 ON 客戶資料表.客戶編號 = 收款資料表.客戶編號 "
  cmd = cmd & cid_date_criteria
  'cmd = cmd & "ORDER BY 收款資料表.客戶編號, 收款資料表.收款日期, 收款資料表.識別碼;"
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY 收款資料表.客戶編號, 收款資料表.收款日期, 收款資料表.識別碼;"
  Else
  cmd = cmd & "ORDER BY 收款資料表.收款日期, 收款資料表.識別碼;"
  End If
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh
End Sub

Private Sub dispChkInDeal()
  Dim cmd As String
  'cmd = "SELECT 過帳交易資料表.客戶編號, 客戶資料表.客戶姓名, 過帳交易資料表.交易日期, 魚貨資料表.魚貨代號, "
  'cmd = cmd & "過帳交易資料表.識別碼, 魚貨資料表.魚貨名稱, 過帳交易資料表.單位, "
  'cmd = cmd & "過帳交易資料表.重量, 過帳交易資料表.單價, 過帳交易資料表.稅別, 過帳交易資料表.持分, "
  'cmd = cmd & "過帳交易資料表.金額, 過帳交易資料表.其他, 過帳交易資料表.合計 "
  
  cmd = "SELECT 過帳交易資料表.客戶編號, 客戶資料表.客戶姓名, 過帳交易資料表.交易日期, "
  cmd = cmd & "過帳交易資料表.識別碼, 魚貨資料表.魚貨代號, 魚貨資料表.魚貨名稱, 過帳交易資料表.單位, "
  cmd = cmd & "過帳交易資料表.重量, 過帳交易資料表.單價, 過帳交易資料表.稅別, 稅率資料表.稅率, "
  cmd = cmd & "稅率資料表.傳票, 稅率資料表.籠子, 過帳交易資料表.持分, 過帳交易資料表.金額, "
  cmd = cmd & "過帳交易資料表.其他, 過帳交易資料表.合計 "
  cmd = cmd & "FROM 稅率資料表 INNER JOIN ((過帳交易資料表 INNER JOIN 客戶資料表 ON "
  cmd = cmd & "過帳交易資料表.客戶編號 = 客戶資料表.客戶編號) INNER JOIN 魚貨資料表 ON "
  cmd = cmd & "過帳交易資料表.魚貨代號 = 魚貨資料表.魚貨代號) ON 稅率資料表.識別碼 = 過帳交易資料表.稅別 "
  cmd = cmd & cid_date_criteria
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY 過帳交易資料表.客戶編號, 過帳交易資料表.交易日期, 過帳交易資料表.識別碼;"
  Else
  cmd = cmd & "ORDER BY 過帳交易資料表.交易日期, 過帳交易資料表.識別碼;"
  End If
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh

End Sub
Private Sub dispChkInReceive()
  Dim cmd As String
  cmd = "SELECT 過帳收款資料表.客戶編號, 客戶資料表.客戶姓名, 過帳收款資料表.收款日期, "
  cmd = cmd & "過帳收款資料表.識別碼, 過帳收款資料表.收款金額 "
  cmd = cmd & "FROM 過帳收款資料表 INNER JOIN 客戶資料表 ON 過帳收款資料表.客戶編號 = 客戶資料表.客戶編號 "
  cmd = cmd & cid_date_criteria
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY 過帳收款資料表.客戶編號, 過帳收款資料表.收款日期, 過帳收款資料表.識別碼;"
  Else
  cmd = cmd & "ORDER BY 過帳收款資料表.收款日期, 過帳收款資料表.識別碼;"
  End If
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh
End Sub

Private Sub dispChkInDealSum()
  Dim cmd As String
  cmd = "SELECT 交易資料表.客戶編號, 交易資料表.交易日期, Sum(交易資料表.金額) AS 金額之總計 "
  cmd = cmd & "FROM (稅率資料表 INNER JOIN (魚貨資料表 INNER JOIN 交易資料表 ON "
  cmd = cmd & "(魚貨資料表.魚貨代號 = 交易資料表.魚貨代號) AND (魚貨資料表.魚貨代號 = "
  cmd = cmd & "交易資料表.魚貨代號)) ON (稅率資料表.識別碼 = 交易資料表.稅別) AND "
  cmd = cmd & "(稅率資料表.識別碼 = 交易資料表.稅別)) INNER JOIN 客戶資料表 ON "
  cmd = cmd & "(交易資料表.客戶編號 = 客戶資料表.客戶編號) AND (交易資料表.客戶編號 = 客戶資料表.客戶編號) "
  cmd = cmd & "GROUP BY 交易資料表.客戶編號, 交易資料表.交易日期 "
  cmd = cmd & sum_criteria
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY 交易資料表.客戶編號, 交易資料表.交易日期;"
  Else
  cmd = cmd & "ORDER BY 交易資料表.交易日期, 交易資料表.客戶編號;"
  End If
  dbDataBase2.CommandType = adCmdText
  dbDataBase2.RecordSource = cmd
  dbDataBase2.Refresh
End Sub
Private Sub dispChkInReceiveSum()
  Dim cmd As String
  cmd = "SELECT 收款資料表.客戶編號, 收款資料表.收款日期, Sum(收款資料表.收款金額) AS 收款金額之總計 "
  cmd = cmd & "FROM 客戶資料表 INNER JOIN 收款資料表 ON 客戶資料表.客戶編號=收款資料表.客戶編號 "
  cmd = cmd & "GROUP BY 收款資料表.客戶編號, 收款資料表.收款日期 "
  cmd = cmd & sum_criteria
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY 收款資料表.客戶編號, 收款資料表.收款日期;"
  Else
  cmd = cmd & "ORDER BY 收款資料表.收款日期, 收款資料表.客戶編號;"
  End If
  dbDataBase2.CommandType = adCmdText
  dbDataBase2.RecordSource = cmd
  dbDataBase2.Refresh
End Sub

Private Sub CheckDealFunc()
  Dim cmd As String
  CheckWindow.Visible = True
  FrmCheckIn.Enabled = False
  'dbDataBase3.CommandType = adCmdTable
  'dbDataBase3.RecordSource = "收銀資料表"
  'dbDataBase3.Refresh
  'dbDataBase2.Recordset.MoveFirst
  'Do Until dbDataBase2.Recordset.EOF
  '  dbDataBase3.Recordset.AddNew
  '  dbDataBase3.Recordset.Fields("客戶編號") = dbDataBase2.Recordset.Fields("客戶編號")
  '  dbDataBase3.Recordset.Fields("交易日期") = dbDataBase2.Recordset.Fields("交易日期")
  '  dbDataBase3.Recordset.Fields("交易模式") = False
  '  dbDataBase3.Recordset.Fields("交易金額") = dbDataBase2.Recordset.Fields("金額之總計")
  '  dbDataBase3.Recordset.Update
  '  dbDataBase2.Recordset.MoveNext
  'Loop
  ' 新增過帳交易資料表
  dbDataBase3.CommandType = adCmdTable
  dbDataBase3.RecordSource = "過帳交易資料表"
  dbDataBase3.Refresh
  dbDataBase1.Recordset.MoveFirst
  Do Until dbDataBase1.Recordset.EOF
    dbDataBase3.Recordset.AddNew
    dbDataBase3.Recordset.Fields("客戶編號") = dbDataBase1.Recordset.Fields("客戶編號")
    dbDataBase3.Recordset.Fields("交易日期") = dbDataBase1.Recordset.Fields("交易日期")
    'dbDataBase3.Recordset.Fields("傳票編號") = dbDataBase1.Recordset.Fields("傳票編號")
    dbDataBase3.Recordset.Fields("魚貨代號") = dbDataBase1.Recordset.Fields("魚貨代號")
    dbDataBase3.Recordset.Fields("單位") = dbDataBase1.Recordset.Fields("單位")
    dbDataBase3.Recordset.Fields("重量") = dbDataBase1.Recordset.Fields("重量")
    dbDataBase3.Recordset.Fields("單價") = dbDataBase1.Recordset.Fields("單價")
    dbDataBase3.Recordset.Fields("稅別") = dbDataBase1.Recordset.Fields("稅別")
    dbDataBase3.Recordset.Fields("持分") = dbDataBase1.Recordset.Fields("持分")
    dbDataBase3.Recordset.Fields("金額") = dbDataBase1.Recordset.Fields("金額")
    dbDataBase3.Recordset.Fields("其他") = dbDataBase1.Recordset.Fields("其他")
    dbDataBase3.Recordset.Fields("合計") = dbDataBase1.Recordset.Fields("合計")
    dbDataBase3.Recordset.Update
    dbDataBase1.Recordset.MoveNext
  Loop
  ' 移除交易資料表
  dbDataBase3.CommandType = adCmdTable
  dbDataBase3.RecordSource = "交易資料表"
  dbDataBase3.Refresh
  dbDataBase1.Recordset.MoveFirst
  Do Until dbDataBase1.Recordset.EOF
    cmd = "識別碼=" & dbDataBase1.Recordset.Fields("識別碼")
    dbDataBase3.Recordset.MoveFirst
    dbDataBase3.Recordset.Find cmd
    dbDataBase3.Recordset.Update
    dbDataBase3.Recordset.Delete
    dbDataBase3.Recordset.Update
    dbDataBase1.Recordset.MoveNext
  Loop
  Timer1.Enabled = True
End Sub

Private Sub CheckReceiveFunc()
  Dim cmd As String
  CheckWindow.Visible = True
  FrmCheckIn.Enabled = False
  'dbDataBase3.CommandType = adCmdTable
  'dbDataBase3.RecordSource = "收銀資料表"
  'dbDataBase3.Refresh
  'dbDataBase2.Recordset.MoveFirst
  'Do Until dbDataBase2.Recordset.EOF
  '  dbDataBase3.Recordset.AddNew
  '  dbDataBase3.Recordset.Fields("客戶編號") = dbDataBase2.Recordset.Fields("客戶編號")
  '  dbDataBase3.Recordset.Fields("交易日期") = dbDataBase2.Recordset.Fields("收款日期")
  '  dbDataBase3.Recordset.Fields("交易模式") = True
  '  dbDataBase3.Recordset.Fields("交易金額") = -dbDataBase2.Recordset.Fields("收款金額之總計")
  '  dbDataBase3.Recordset.Update
  '  dbDataBase2.Recordset.MoveNext
  'Loop
  ' 新增過帳交易資料表
  dbDataBase3.CommandType = adCmdTable
  dbDataBase3.RecordSource = "過帳收款資料表"
  dbDataBase3.Refresh
  dbDataBase1.Recordset.MoveFirst
  Do Until dbDataBase1.Recordset.EOF
    dbDataBase3.Recordset.AddNew
    dbDataBase3.Recordset.Fields("客戶編號") = dbDataBase1.Recordset.Fields("客戶編號")
    dbDataBase3.Recordset.Fields("收款日期") = dbDataBase1.Recordset.Fields("收款日期")
    dbDataBase3.Recordset.Fields("收款金額") = dbDataBase1.Recordset.Fields("收款金額")
    dbDataBase3.Recordset.Update
    dbDataBase1.Recordset.MoveNext
  Loop
  ' 移除交易資料表
  dbDataBase3.CommandType = adCmdTable
  dbDataBase3.RecordSource = "收款資料表"
  dbDataBase3.Refresh
  dbDataBase1.Recordset.MoveFirst
  Do Until dbDataBase1.Recordset.EOF
    cmd = "識別碼=" & dbDataBase1.Recordset.Fields("識別碼")
    dbDataBase3.Recordset.MoveFirst
    dbDataBase3.Recordset.Find cmd
    dbDataBase3.Recordset.Update
    dbDataBase3.Recordset.Delete
    dbDataBase3.Recordset.Update
    dbDataBase1.Recordset.MoveNext
  Loop
  Timer1.Enabled = True
End Sub

Private Sub checkDate_Change(Index As Integer)
  If checkDate(0) > checkDate(1) Then
    Beep
    checkDate(1) = checkDate(0)
  End If
  date_vaild = True
  Call update_disp
End Sub

Private Sub cid_Change(Index As Integer)
  If cid(0) = "" And cid(1) = "" Then
    start_cid = "000"
    end_cid = "999"
  Else
    If cid(0) = "" Then
    start_cid = "000"
    Else
      start_cid = String(3 - Len(cid(0)), "0") & cid(0)
    End If
    If cid(1) = "" Then
      end_cid = start_cid
    Else
      end_cid = String(3 - Len(cid(1)), "0") & cid(1)
    End If
  End If
  If start_cid > end_cid Then
    Beep
    cid(1).SetFocus
    cid_vaild = False
  Else
    cid_vaild = True
    Call update_disp
  End If
End Sub

Private Sub cid_LostFocus(Index As Integer)
  If Index = 0 Then
    If cid(0) <> "" Then
      cid(0) = String(3 - Len(cid(0)), "0") & cid(0)
    End If
  ElseIf cid(1) <> "" Then
    cid(1) = String(3 - Len(cid(1)), "0") & cid(1)
  End If
End Sub

Private Sub PrintDeal()
  Dim i As Integer
  Dim now_date As Date
  Dim now_day As String
  Dim now_time As String
  Dim CurrPage As Integer
  Dim TotalPage As Integer
  Dim ListPage As Integer
  Dim ListPage2 As Integer
  Dim CurrList As Integer
  
  now_time = Format(Now, "yyyymmdd hh:mm:ss")
  now_date = Now
  'now_day = String(3 - Len(Year(now_date) - 1911), " ") & Year(now_date) - 1911
  now_day = Year(now_date) - 1911
  now_day = now_day & "/" & String(2 - Len(Month(now_date)), "0") & Month(now_date)
  now_day = now_day & "/" & String(2 - Len(Day(now_date)), "0") & Day(now_date)
  now_day = now_day & Mid(now_time, 9, 8)
  now_day = now_day & String(18 - Len(now_day), " ")
  
  ListPage = 10 * 6
  TotalPage = Fix((total_list + ListPage - 1) / ListPage)
  For CurrPage = 1 To TotalPage
    Text1.Text = vbCrLf & vbCrLf
    If CurrPage <> TotalPage Then
      CurrList = ListPage
    Else
      CurrList = total_list Mod ListPage
      If CurrList = 0 And total_list <> 0 Then
        CurrList = ListPage
      End If
    End If
    Text1.Text = Text1.Text & String(38, " ")
    If CheckMode = 0 Then
      Text1.Text = Text1.Text & REPORT_START_SPAN1 & "** 未 過 帳 - 交 易 明 細 表 **" & vbCrLf
    ElseIf CheckMode = 2 Then
      Text1.Text = Text1.Text & REPORT_START_SPAN1 & "** 已 過 帳 - 交 易 明 細 表 **" & vbCrLf
    End If
    Text1.Text = Text1.Text & REPORT_START_SPAN1 & "製表依據: "
    If optSort(0).value = True Then
      Text1.Text = Text1.Text & "依編號別" & vbCrLf
    Else
      Text1.Text = Text1.Text & "依日期別" & vbCrLf
    End If
    Text1.Text = Text1.Text & REPORT_START_SPAN1 & "製表日期: " & now_day & vbCrLf
    'If DateChk.Value = 1 Then
      Text1.Text = Text1.Text & REPORT_START_SPAN1 & "日期區間: " & DC2PC(CDate(checkDate(0))) & " 至 " & DC2PC(CDate(checkDate(1))) & vbCrLf
    'End If
    Text1.Text = Text1.Text & REPORT_START_SPAN1 & "表單頁次: " & CurrPage & vbCrLf
    Text1.Text = Text1.Text & REPORT_START_SPAN1 & PREORT_DEAL_HEADER1 & vbCrLf
    Text1.Text = Text1.Text & REPORT_START_SPAN1 & PREORT_DEAL_HEADER2 & vbCrLf
    Text1.Text = Text1.Text & REPORT_START_SPAN1 & PREORT_DEAL_HEADER3 & vbCrLf
    For i = 0 To CurrList - 1
      Text1.Text = Text1.Text & REPORT_START_SPAN1 & list_report((CurrPage - 1) * ListPage + i) & vbCrLf
    Next
    CurrList = CurrList + 9
    'For i = CurrList To (10 * 6) + 9
    '  Text1.Text = Text1.Text & vbCrLf
    'Next
    Printer.Print Text1
    If CurrPage <> TotalPage Then
      Printer.NewPage
    End If
  Next
  Printer.EndDoc

End Sub

Private Sub PrintReceive()
  Dim i As Integer
  Dim now_date As Date
  Dim now_day As String
  Dim now_time As String
  Dim CurrPage As Integer
  Dim TotalPage As Integer
  Dim ListPage As Integer
  Dim ListPage2 As Integer
  Dim CurrList As Integer
  
  now_time = Format(Now, "yyyymmdd hh:mm:ss")
  now_date = Now
  'now_day = String(3 - Len(Year(now_date) - 1911), " ") & Year(now_date) - 1911
  now_day = Year(now_date) - 1911
  now_day = now_day & "/" & String(2 - Len(Month(now_date)), "0") & Month(now_date)
  now_day = now_day & "/" & String(2 - Len(Day(now_date)), "0") & Day(now_date)
  now_day = now_day & Mid(now_time, 9, 8)
  now_day = now_day & String(18 - Len(now_day), " ")
  
  ListPage = 10 * 6
  ListPage2 = 10 * 6 * 2
  TotalPage = Fix((total_list + ListPage2 - 1) / ListPage2)
  For CurrPage = 1 To TotalPage
    Text1.Text = vbCrLf & vbCrLf
    If CurrPage <> TotalPage Then
      CurrList = ListPage2
    Else
      ' last page
      CurrList = total_list Mod ListPage2
      If CurrList = 0 Then
        CurrList = ListPage2
      End If
    End If
    'If CurrList > ListPage Then
      Text1.Text = Text1.Text & String(34, " ")
    'Else
    '  Text1.Text = Text1.Text & String(8, " ")
    'End If
    If CheckMode = 1 Then
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & "** 未 過 帳 - 收 款 明 細 表 **" & vbCrLf
    ElseIf CheckMode = 3 Then
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & "** 已 過 帳 - 收 款 明 細 表 **" & vbCrLf
    End If
    Text1.Text = Text1.Text & REPORT_START_SPAN2 & "製表依據: "
    If optSort(0).value = True Then
      Text1.Text = Text1.Text & "依編號別" & vbCrLf
    Else
      Text1.Text = Text1.Text & "依日期別" & vbCrLf
    End If
    Text1.Text = Text1.Text & REPORT_START_SPAN2 & "製表日期: " & now_day & vbCrLf
    'If DateChk.Value = 1 Then
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & "日期區間: " & DC2PC(CDate(checkDate(0))) & " 至 " & DC2PC(CDate(checkDate(1))) & vbCrLf
    'End If
    Text1.Text = Text1.Text & REPORT_START_SPAN2 & "表單頁次: " & CurrPage & vbCrLf
    If CurrList > ListPage Then
      '雙排
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & PREORT_RECEIVE_HEADER1 & REPORT_RECEIVE_COLUMN_SPAN & PREORT_RECEIVE_HEADER1 & vbCrLf
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & PREORT_RECEIVE_HEADER2 & REPORT_RECEIVE_COLUMN_SPAN & PREORT_RECEIVE_HEADER2 & vbCrLf
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & PREORT_RECEIVE_HEADER3 & REPORT_RECEIVE_COLUMN_SPAN & PREORT_RECEIVE_HEADER3 & vbCrLf
      For i = 0 To ListPage - 1
        If i < (CurrList Mod ListPage) Then
          Text1.Text = Text1.Text & REPORT_START_SPAN2 & list_report((CurrPage - 1) * ListPage2 + i) & REPORT_RECEIVE_COLUMN_SPAN & list_report((CurrPage - 1) * ListPage2 + ListPage + i) & vbCrLf
        Else
          Text1.Text = Text1.Text & REPORT_START_SPAN2 & list_report((CurrPage - 1) * ListPage2 + i) & REPORT_RECEIVE_COLUMN_SPAN & vbCrLf
        End If
      Next
    Else '單排
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & PREORT_RECEIVE_HEADER1 & vbCrLf
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & PREORT_RECEIVE_HEADER2 & vbCrLf
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & PREORT_RECEIVE_HEADER3 & vbCrLf
      For i = 0 To CurrList - 1
        Text1.Text = Text1.Text & REPORT_START_SPAN2 & list_report((CurrPage - 1) * ListPage2 + i) & vbCrLf
      Next
      'CurrList = CurrList + 9
    End If
    'For i = CurrList To (10 * 6) + 9
    '  Text1.Text = Text1.Text & vbCrLf
    'Next
    Printer.Print Text1
    If CurrPage <> TotalPage Then
      Printer.NewPage
    End If
  Next
  Printer.EndDoc
End Sub

Private Sub cmdPrint_Click()
  Call SetPrintPage(1)
  PrintDialog.CancelError = True
  On Error GoTo ErrHandlerPrintCheck
  PrintDialog.Flags = cdlPDReturnDC + cdlPDPageNums + cdlPDNoSelection + cdlPDDisablePrintToFile + cdlPDAllPages + cdlPDCollate
  PrintDialog.ShowPrinter
  Call SetPrintPage(1)
  Printer.FontName = "細明體"
  Printer.FontSize = 10
  If CheckMode = 1 Or CheckMode = 3 Then
    Call PrintReceive
  Else
    Call PrintDeal
  End If
  'Call PrintRTF(Text1, 800, 700, 400, 300)
  Screen.MousePointer = 0
ErrHandlerPrintCheck:
End Sub

'Private Sub DateChk_Click()
'  If DateChk.Value = 1 Then
'    date_vaild = True
'    Call update_disp
'  Else
'    If checkDate(0) > checkDate(1) Then
'      date_vaild = False
'      Beep
'      checkDate(1).SetFocus
'    Else
'      date_vaild = True
'      Call update_disp
'    End If
'  End If
'End Sub

Private Sub CheckInButton_Click()
  If CheckMode = 0 Then
    Call CheckDealFunc
  ElseIf CheckMode = 1 Then
    Call CheckReceiveFunc
  ElseIf CheckMode = 2 Then
  ElseIf CheckMode = 3 Then
  End If
End Sub

Private Sub listDataBase_Scroll(Cancel As Integer)
  listDataBase.Refresh
End Sub

Private Sub optSort_Click(Index As Integer)
  Call update_disp
End Sub

Private Sub TabStrip1_Click()
  If TabStrip1.Tabs.item(1).Selected = True Then
    CheckMode = 0
    CheckInButton.Visible = True
  ElseIf TabStrip1.Tabs.item(2).Selected = True Then
    CheckMode = 1
    CheckInButton.Visible = True
  ElseIf TabStrip1.Tabs.item(3).Selected = True Then
    CheckMode = 2
    CheckInButton.Visible = False
  Else
    CheckMode = 3
    CheckInButton.Visible = False
  End If
  
  If cid_vaild = False Then
    Beep
    'cid(1).SetFocus
  ElseIf date_vaild = False Then
    Beep
    checkDate(1).SetFocus
  Else
    Call update_disp
  End If
End Sub

Private Sub Timer1_Timer()
  CheckWindow.Visible = False
  Call update_disp
  Timer1.Enabled = False
  FrmCheckIn.Enabled = True
End Sub

Private Sub Form_Load()
  Dim cmd As String
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  dbDataBase1.ConnectionString = database_string
  dbDataBase2.ConnectionString = database_string
  dbDataBase3.ConnectionString = database_string
  start_cid = "000"
  end_cid = "999"
  cid_vaild = True
  date_vaild = True
  If CheckMode = 0 Then
    TabStrip1.Tabs.item(1).Selected = True
  ElseIf CheckMode = 1 Then
    TabStrip1.Tabs.item(2).Selected = True
  ElseIf CheckMode = 2 Then
    TabStrip1.Tabs.item(3).Selected = True
  Else
    TabStrip1.Tabs.item(4).Selected = True
  End If
  checkDate(0) = DateValue(Now)
  checkDate(1) = DateValue(Now)
  cmd = "SELECT 交易資料表.識別碼, 交易資料表.客戶編號, 客戶資料表.客戶姓名, 交易資料表.交易日期, 交易資料表.魚貨代號, "
  cmd = cmd & "魚貨資料表.魚貨名稱, 交易資料表.單位, 交易資料表.重量, 交易資料表.單價, 交易資料表.稅別, 交易資料表.金額, 交易資料表.其他, 交易資料表.合計 "
  cmd = cmd & "FROM (稅率資料表 INNER JOIN (魚貨資料表 INNER JOIN 交易資料表 ON (魚貨資料表.魚貨代號 = 交易資料表.魚貨代號) "
  cmd = cmd & "AND (魚貨資料表.魚貨代號 = 交易資料表.魚貨代號)) ON (稅率資料表.識別碼 = 交易資料表.稅別) AND (稅率資料表.識別碼 = 交易資料表.稅別)) "
  cmd = cmd & "INNER JOIN 客戶資料表 ON (交易資料表.客戶編號 = 客戶資料表.客戶編號) AND (交易資料表.客戶編號 = 客戶資料表.客戶編號) "
  cmd = cmd & "ORDER BY 交易資料表.客戶編號, 交易資料表.交易日期;"
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh
  cmd = "SELECT 交易資料表.客戶編號, 交易資料表.交易日期, Sum(交易資料表.金額) AS 金額之總計 "
  cmd = cmd & "FROM (稅率資料表 INNER JOIN (魚貨資料表 INNER JOIN 交易資料表 ON (魚貨資料表.魚貨代號 "
  cmd = cmd & "= 交易資料表.魚貨代號) AND (魚貨資料表.魚貨代號 = 交易資料表.魚貨代號)) ON (稅率資料表.識別碼 "
  cmd = cmd & "= 交易資料表.稅別) AND (稅率資料表.識別碼 = 交易資料表.稅別)) INNER JOIN 客戶資料表 ON "
  cmd = cmd & "(交易資料表.客戶編號 = 客戶資料表.客戶編號) AND (交易資料表.客戶編號 = 客戶資料表.客戶編號) "
  cmd = cmd & "GROUP BY 交易資料表.客戶編號, 交易資料表.交易日期 "
  cmd = cmd & "ORDER BY 交易資料表.交易日期, 交易資料表.客戶編號;"
  dbDataBase2.CommandType = adCmdText
  dbDataBase2.RecordSource = cmd
  dbDataBase2.Refresh
  dbDataBase3.CommandType = adCmdTable
  dbDataBase3.RecordSource = "過帳交易資料表"
  dbDataBase3.Refresh
  
  Set listDataBase.DataSource = dbDataBase1
  Call update_disp
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub

