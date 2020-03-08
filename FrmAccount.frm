VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmAccount 
   BorderStyle     =   1  '單線固定
   Caption         =   "營業彙總表"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   8.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11880
   Begin VB.CommandButton cmdAccount 
      Caption         =   "結算"
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
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc dbDataBase1 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin MSDataGridLib.DataGrid listBusSummary 
      Height          =   6375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   11245
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
   Begin MSAdodcLib.Adodc dbDataBase2 
      Height          =   375
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin VB.Label lblDate 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "結算日期:"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "FrmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private procDate As String
Private Sub set_listBusSummary()
  listBusSummary.Columns.item(0).Width = 1400
  listBusSummary.Columns.item(1).Width = 1400
  listBusSummary.Columns.item(2).Width = 1400
  listBusSummary.Columns.item(3).Width = 1400
  listBusSummary.Columns.item(4).Width = 1400
  listBusSummary.Columns.item(5).Width = 1400
End Sub

Private Sub cmdAccount_Click()
  Dim cid As String
  Dim cMoney As String
  Dim DateStr As String
  Dim CurrDate As String
  Dim CmpY1 As Integer
  Dim CmpM1 As Integer
  Dim CmpY2 As Integer
  Dim CmpM2 As Integer
  CmpY1 = Year(procDate)
  CmpM1 = Month(procDate)
  Do Until dbDataBase2.Recordset.EOF
    CurrDate = dbDataBase2.Recordset.Fields("結算日期")
    CmpY2 = Year(CurrDate)
    CmpM2 = Month(CurrDate)
    If CmpY1 = CmpY2 And CmpM1 = CmpM2 Then
      dbDataBase2.Recordset.Delete
    End If
    dbDataBase2.Recordset.MoveNext
  Loop
  
  If dbDataBase1.Recordset.EOF = False Or dbDataBase1.Recordset.BOF = False Then
    dbDataBase1.Recordset.MoveFirst
  End If
  
  Do Until dbDataBase1.Recordset.EOF
    cid = dbDataBase1.Recordset.Fields("客戶編號")
    cMoney = dbDataBase1.Recordset.Fields("應收餘額")
    With dbDataBase2.Recordset
    .AddNew
    .Fields("客戶編號") = cid
    .Fields("結算日期") = procDate
    .Fields("結算金額") = cMoney
    .Update
    End With
    dbDataBase1.Recordset.MoveNext
  Loop
  
  If vYY < 1900 Then
    DateStr = CStr(vYY + 1911) & "/" & vMM & "/1"
  Else
    DateStr = vYY & "/" & vMM & "/1"
  End If
  dbDataBase2.CommandType = adCmdTable
  dbDataBase2.RecordSource = "系統資料表"
  dbDataBase2.Refresh
  dbDataBase2.Recordset.Fields("前次結算日期") = DateStr
  dbDataBase2.Recordset.Update
  Unload Me
End Sub

Private Sub Form_Load()
  Dim cmd As String
  Dim date_str As String
  Dim date_start As Date
  Dim date_stop As Date
  Dim cid As String
  Dim cName As String
  Dim cMoney As String
  Dim cList As String
  Dim currentYear As String
  Dim now_str As Date
  Dim tmpY As Integer
  Dim tmpM As Integer
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  If vMM = 1 Then
    tmpY = vYY - 1
    tmpM = 12
  Else
    tmpY = vYY
    tmpM = vMM - 1
  End If
  Me.Caption = tmpY & "年" & tmpM & "月份營業結算"
  If vYY < 1900 Then
    date_str = (tmpY + 1911) & "/" & tmpM & "/1"
    procDate = (vYY + 1911) & "/" & vMM & "/1"
  Else
    date_str = tmpY & "/" & tmpM & "/1"
    procDate = vYY & "/" & vMM & "/1"
  End If
  now_str = Now
  currentYear = Year(now_str) - 1911
  date_start = date_str
  date_stop = DateAdd("m", 1, date_str)
  date_stop = DateAdd("d", -1, date_stop)
  lblDate = DC2PC(CDate(date_str)) & " 至 " & DC2PC(CDate(date_stop))
  dbDataBase1.ConnectionString = database_string
  dbDataBase2.ConnectionString = database_string
  cmd = "SELECT 營業彙總表.客戶編號, 客戶資料表.客戶姓名, Sum(營業彙總表.結算金額) "
  cmd = cmd & "AS 期初餘額, Sum(營業彙總表.交易金額) AS 本期銷售, Sum(營業彙總表.收款金額) "
  cmd = cmd & "AS 本期收款, [期初餘額]+[本期銷售]-[本期收款] AS 應收餘額 "
  cmd = cmd & "FROM 客戶資料表 INNER JOIN 營業彙總表 ON 客戶資料表.客戶編號 = 營業彙總表.客戶編號 "
  cmd = cmd & "WHERE (((營業彙總表.結算日期) Between #" & date_start & "# And #" & date_stop & "#)) "
  cmd = cmd & "GROUP BY 營業彙總表.客戶編號, 客戶資料表.客戶姓名 "
  cmd = cmd & "HAVING ((營業彙總表.客戶編號) Between '" & "000" & "' And '" & "999" & "')"
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh
  dbDataBase2.CommandType = adCmdTable
  dbDataBase2.RecordSource = "結算資料表"
  dbDataBase2.Refresh
  Set listBusSummary.DataSource = dbDataBase1
  Call set_listBusSummary
  If dbDataBase1.Recordset.EOF = False Then
    cmdAccount.Enabled = True
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub

