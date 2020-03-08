VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmReceive 
   BorderStyle     =   1  '單線固定
   Caption         =   "收款資料登錄及修改"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10545
   Icon            =   "FrmReceive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   10545
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame DeleteWindow 
      Height          =   2415
      Left            =   3840
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Label12 
         Caption         =   "資料刪除中...請稍候"
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
         Left            =   1080
         TabIndex        =   25
         Top             =   1080
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1935
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   4215
      Begin VB.TextBox rMoney 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   18
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox cId 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton ReceiveCmd 
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
         Left            =   2880
         TabIndex        =   16
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "收款金額"
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
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "收款日期"
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
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
      Begin VB.Label rDate 
         BorderStyle     =   1  '單線固定
         Caption         =   "Label12"
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
         Left            =   1320
         TabIndex        =   21
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label cName 
         BorderStyle     =   1  '單線固定
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
         Left            =   2640
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc dbDataBase5 
      Height          =   330
      Left            =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dbDataBase5"
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame debugReceive 
      Caption         =   "Debug-Receive"
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox receiveItem 
         Height          =   270
         Index           =   4
         Left            =   4200
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox receiveItem 
         Height          =   270
         Index           =   2
         Left            =   2040
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox receiveItem 
         Height          =   270
         Index           =   3
         Left            =   3120
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox custId 
         Height          =   270
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox custName 
         Height          =   270
         Left            =   1080
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox receiveItem 
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox receiveItem 
         Height          =   270
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "操作模式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton ReceiveMode 
         Caption         =   "刪除"
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
         Index           =   3
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton ReceiveMode 
         Caption         =   "修改"
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
         Index           =   2
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton ReceiveMode 
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
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton ReceiveMode 
         Caption         =   "查詢"
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
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid listReceive 
      Height          =   7095
      Left            =   4800
      TabIndex        =   4
      Top             =   360
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   12515
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
   Begin MSDataGridLib.DataGrid listCustom 
      Height          =   4215
      Left            =   480
      TabIndex        =   5
      Top             =   3240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7435
      _Version        =   393216
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
   Begin MSAdodcLib.Adodc dbDataBase1 
      Height          =   330
      Left            =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "客戶查詢"
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
   Begin MSAdodcLib.Adodc dbDataBase4 
      Height          =   330
      Left            =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dbDataBase4"
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
End
Attribute VB_Name = "FrmReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private updateCId As Boolean
Private updateCName As Boolean
Private start_flag As Boolean
Private receiveItemMode As Boolean
Private ACDate As String
Private fSkip As Boolean

Private Sub set_listReceive()
  listReceive.Columns.item(0).Width = 1400
  listReceive.Columns.item(1).Width = 1400
  listReceive.Columns.item(2).Width = 1400
  listReceive.Columns.item(3).Width = 12
  listReceive.Columns.item(4).Width = 1400
End Sub
Private Sub set_listCustom()
  listCustom.Columns.item(0).Width = 12
  listCustom.Columns.item(1).Width = 1400
  listCustom.Columns.item(2).Width = 1400
  listCustom.Columns.item(3).Width = 12
End Sub
Private Sub locked_edit()
  rMoney.Locked = True
  ReceiveCmd.Visible = False
End Sub
Private Sub unlocked_edit()
  rMoney.Locked = False
  ReceiveCmd.Visible = True
End Sub
Private Sub clear_edit()
  cid = ""
  rMoney = ""
End Sub
Private Sub displayAllReceive()
  Dim cmd As String
  Dim flag As Boolean
  cmd = "SELECT 收款資料表.客戶編號, 客戶資料表.客戶姓名, "
  cmd = cmd & "收款資料表.收款日期, 收款資料表.識別碼, 收款資料表.收款金額 "
  cmd = cmd & "FROM 客戶資料表 INNER JOIN 收款資料表 ON "
  cmd = cmd & "客戶資料表.客戶編號 = 收款資料表.客戶編號 "
  'cmd = cmd & "ORDER BY 收款資料表.客戶編號, 收款資料表.收款日期, 收款資料表.識別碼;"
  cmd = cmd & "ORDER BY 收款資料表.識別碼;"
  dbDataBase4.CommandType = adCmdText
  dbDataBase4.RecordSource = cmd
  'dbDataBase4.Recordset.MoveFirst
  flag = IsEmpty(dbDataBase4.Recordset)
  If flag = False Then
    dbDataBase4.Refresh
  End If
  Call set_listReceive
End Sub
Private Sub displayPartReceive()
  Call displayAllReceive
  'Dim cmd As String
  'Dim cid_str As String
  'Dim flag As Boolean
  'cid_str = appendstr(cid, 3)
  'cmd = "SELECT 收款資料表.客戶編號, 客戶資料表.客戶姓名, "
  'cmd = cmd & "收款資料表.收款日期, 收款資料表.識別碼, 收款資料表.收款金額 "
  'cmd = cmd & "FROM 客戶資料表 INNER JOIN 收款資料表 ON "
  'cmd = cmd & "客戶資料表.客戶編號 = 收款資料表.客戶編號 "
  'cmd = cmd & "WHERE (((客戶資料表.客戶編號)= '" & cid_str & "')) "
  'cmd = cmd & "ORDER BY 收款資料表.客戶編號, 收款資料表.收款日期, 收款資料表.識別碼;"
  'dbDataBase4.CommandType = adCmdText
  'dbDataBase4.RecordSource = cmd
  ''dbDataBase4.Recordset.MoveFirst
  'flag = IsEmpty(dbDataBase4.Recordset)
  'If flag = False Then
  '  dbDataBase4.Refresh
  'End If
  ''ReceiveCmd.SetFocus
  'Call set_listReceive
End Sub
Private Function checkInputVaild() As Boolean
  If cid = "" Then
    Beep
    cid.SelStart = 0
    cid.SelLength = Len(cid)
    cid.SetFocus
    checkInputVaild = False
  ElseIf cName = "" Then
    Beep
    cid.SelStart = 0
    cid.SelLength = Len(cid)
    cid.SetFocus
    checkInputVaild = False
  ElseIf IsNumeric(rMoney) = False Then
    Beep
    rMoney.SelStart = 0
    rMoney.SelLength = Len(rMoney)
    rMoney.SetFocus
    checkInputVaild = False
  Else
    checkInputVaild = True
  End If
End Function

Private Sub cId_KeyPress(KeyAscii As Integer)
  If cid <> "" Then
    If KeyAscii = vbKeyReturn Then
      rMoney.SetFocus
      KeyAscii = 0
    ElseIf ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 8)) = False Then
      KeyAscii = 0
      Beep
    ElseIf KeyAscii <> 8 Then
      If Len(cid) = 3 And cid.SelLength = 0 Then
        Beep
        KeyAscii = 0
      End If
    End If
  End If
End Sub

Private Sub cId_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim cmd As String
  Dim flag As Boolean
  flag = False
  cmd = cid
  If cmd <> "" Then
    cmd = appendstr(cid, 3)
    cmd = "客戶編號 = '" & cmd & "'"
    If dbDataBase1.Recordset.EOF = True And dbDataBase1.Recordset.BOF = True Then
      flag = True
      cName = ""
    Else
      dbDataBase1.Recordset.MoveFirst
      dbDataBase1.Recordset.Find cmd
      If dbDataBase1.Recordset.EOF = True Then
        flag = True
        cName = ""
      Else
        cName = custName
      End If
    End If
  End If
  'If flag Then
  '  cId.SelStart = 0
  '  cId.SelLength = Len(cId)
  '  Beep
  'Else
  If cid = "" Or flag Then
  '  Call displayAllReceive
    cName = ""
  Else
  '  Call displayPartReceive
  End If
End Sub

Private Sub cid_LostFocus()
  Dim cmd As String
  Dim flag As Integer
  On Error GoTo ErrHandlerCId1
  flag = 0
  cmd = cid
  If cmd <> "" Then
    cmd = appendstr(cid, 3)
    cmd = "客戶編號 = '" & cmd & "'"
    dbDataBase1.Recordset.MoveFirst
    dbDataBase1.Recordset.Find cmd
    If dbDataBase1.Recordset.EOF = True Then
      flag = 1
    End If
  End If
  If flag Then
    cid.SelStart = 0
    cid.SelLength = Len(cid)
    Beep
  ElseIf cid <> "" Then
    cid = appendstr(cid, 3)
    cName = custName
  End If
ErrHandlerCId1:
End Sub

Private Sub listCustom_Click()
  updateCId = True
  updateCName = True
  cid = custId
  cName = custName
End Sub

Private Sub ReceiveCmd_Click()
  Dim cmd As String
  Dim select_index As String
  Dim del_database As Integer
  Dim vaild As Boolean
  select_index = receiveItem(0).Text
  cmd = "識別碼=" & select_index
  If ReceiveMode(1).value = True Then ' 新增
    receiveItemMode = False
    vaild = checkInputVaild
    If vaild Then
      If dbDataBase5.Recordset.EOF = False Then
        dbDataBase5.Recordset.MoveLast
        dbDataBase5.Recordset.Update
      End If
      With dbDataBase5.Recordset
      .AddNew
      .Fields("客戶編號") = cid
      .Fields("收款日期") = ACDate
      .Fields("收款金額") = rMoney
      .Update
      End With
      Call displayAllReceive
      If cmd <> "識別碼=" Then
        dbDataBase4.Recordset.MoveFirst
        dbDataBase4.Recordset.Find cmd
        If dbDataBase4.Recordset.EOF = True Then
          MsgBox ("程式錯誤")
        End If
        dbDataBase4.Recordset.MoveLast
      End If
      fSkip = True
      cid = ""
      cid.SetFocus
      rMoney = ""
      cName = ""
    End If
    receiveItemMode = True
  ElseIf ReceiveMode(2).value = True Then ' 修改
    vaild = checkInputVaild
    receiveItemMode = False
    If vaild Then
      dbDataBase4.Recordset.MoveFirst
      dbDataBase4.Recordset.Find cmd
      If dbDataBase4.Recordset.EOF = True Then
        MsgBox ("程式錯誤")
      End If
      With dbDataBase4.Recordset
      .Fields("客戶編號") = cid
      '.Fields("收款日期") = ACDate  ' 日期不修改
      .Fields("收款金額") = rMoney
      .Update
      End With
      dbDataBase4.Refresh
      dbDataBase4.Recordset.MoveFirst
      dbDataBase4.Recordset.Find cmd
      If dbDataBase4.Recordset.EOF = True Then
        MsgBox ("程式錯誤")
      End If
      If cid = "" Then
        Call displayAllReceive
      Else
        Call displayPartReceive
      End If
      dbDataBase4.Recordset.MoveFirst
      dbDataBase4.Recordset.Find cmd
      If dbDataBase4.Recordset.EOF = True Then
        MsgBox ("程式錯誤")
      End If
      rMoney.SelStart = 0
      rMoney.SelLength = Len(rMoney)
      cid.SetFocus
    End If
    receiveItemMode = True
  ElseIf ReceiveMode(3).value = True Then ' 刪除
    If select_index <> "" Then
      dbDataBase4.Recordset.MoveFirst
      dbDataBase4.Recordset.Find cmd
      If dbDataBase4.Recordset.EOF = False Then
        del_database = MsgBox("確定是否刪除?", vbYesNo, "刪除登錄資料")
        If del_database = vbYes Then
          dbDataBase5.Refresh
          dbDataBase5.Recordset.MoveFirst
          dbDataBase5.Recordset.Find cmd
          dbDataBase5.Recordset.Delete
          dbDataBase5.Recordset.Update
          DeleteWindow.Visible = True
          FrmReceive.Enabled = False
          Timer1.Enabled = True
        End If
      End If
      Call displayAllReceive
    End If
  End If
End Sub

Private Sub receiveItem_Change(Index As Integer)
  If receiveItemMode = True And dbDataBase4.Recordset.BOF = False And dbDataBase4.Recordset.EOF = False Then
    If ReceiveMode(1).value = False Then
      If Index = 1 Then
        cid = receiveItem(Index)
      ElseIf Index = 3 Then
        'rDate = receiveItem(Index)
        rDate = DC2PC(receiveItem(Index))
      ElseIf Index = 4 Then
        rMoney = receiveItem(Index)
      End If
    End If
  End If
End Sub

Private Sub ReceiveMode_Click(Index As Integer)
  If ReceiveMode(0).value = True Then
    Call locked_edit
    cid.Enabled = True
    rMoney.Enabled = True
    Frame1.Enabled = False
    cid = receiveItem(1)
    cName = receiveItem(2)
    If receiveItem(3) = "" Then
      rDate = ""
    Else
      rDate = DC2PC(receiveItem(3))
    End If
    rMoney = receiveItem(4)
  ElseIf ReceiveMode(1).value = True Then
    fSkip = True
    start_flag = True
    Call unlocked_edit
    Call clear_edit
    Frame1.Enabled = True
    cid.Enabled = True
    rMoney.Enabled = True
    ReceiveCmd.Caption = "新增"
    rDate = vYY & "/" & vMM & "/" & vDD
    ACDate = CStr(vYY + 1911) & "/" & vMM & "/" & vDD
    cid.SetFocus
  ElseIf ReceiveMode(2).value = True Then
    Call unlocked_edit
    Frame1.Enabled = True
    ReceiveCmd.Caption = "修改"
    cid.Enabled = True
    rMoney.Enabled = True
    cid = receiveItem(1)
    cName = receiveItem(2)
    If receiveItem(3) = "" Then
      rDate = ""
    Else
      rDate = DC2PC(receiveItem(3))
    End If
    rMoney = receiveItem(4)
  Else
    Call locked_edit
    Frame1.Enabled = True
    ReceiveCmd.Visible = True
    ReceiveCmd.Caption = "刪除"
    cid.Enabled = False
    rMoney.Enabled = False
    cid = receiveItem(1)
    cName = receiveItem(2)
    If receiveItem(3) = "" Then
      rDate = ""
    Else
      rDate = DC2PC(receiveItem(3))
    End If
    rMoney = receiveItem(4)
    Call displayAllReceive
  End If
End Sub

Private Sub rMoney_Change()
 If rMoney = "-" Then
 ElseIf IsNumeric(rMoney) = False And fSkip = False Then
   If rMoney <> "" Then
    Beep
    rMoney = ""
    rMoney.SelStart = 0
    rMoney.SelLength = Len(rMoney)
  End If
 End If
 fSkip = False
End Sub

Private Sub rMoney_KeyPress(KeyAscii As Integer)
  If rMoney <> "" Then
    If KeyAscii = vbKeyReturn Then
      ReceiveCmd.SetFocus
      KeyAscii = 0
    ElseIf ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 8)) = False Then
      KeyAscii = 0
      Beep
    End If
  End If
End Sub

Private Sub Timer1_Timer()
  FrmReceive.Enabled = True
  If cName = "" Then
    Call displayAllReceive
  Else
    Call displayPartReceive
  End If
  If dbDataBase4.Recordset.BOF = False And dbDataBase4.Recordset.EOF = False Then
    dbDataBase4.Recordset.MoveFirst
  End If
  Timer1.Enabled = False
  DeleteWindow.Visible = False
  ReceiveCmd.SetFocus
End Sub

Private Sub Form_Load()
  Dim cmd As String
  fSkip = False
  receiveItemMode = True
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  start_flag = False
  rDate = vYY & "/" & vMM & "/" & vDD
  dbDataBase1.ConnectionString = database_string
  dbDataBase1.CommandType = adCmdTable
  dbDataBase1.RecordSource = "客戶資料表"
  dbDataBase1.Refresh
  cmd = "SELECT 收款資料表.客戶編號, 客戶資料表.客戶姓名, 收款資料表.收款日期, 收款資料表.識別碼, 收款資料表.收款金額 "
  cmd = cmd & "FROM 客戶資料表 INNER JOIN 收款資料表 ON 客戶資料表.客戶編號=收款資料表.客戶編號 "
  cmd = cmd & "ORDER BY 收款資料表.客戶編號, 收款資料表.收款日期, 收款資料表.識別碼;"
  dbDataBase4.ConnectionString = database_string
  dbDataBase4.CommandType = adCmdText
  dbDataBase4.RecordSource = cmd
  dbDataBase4.Refresh
  dbDataBase5.ConnectionString = database_string
  dbDataBase5.CommandType = adCmdTable
  dbDataBase5.RecordSource = "收款資料表"
  dbDataBase5.Refresh
  
  Set listCustom.DataSource = dbDataBase1
  Set listReceive.DataSource = dbDataBase4
  Set custId.DataSource = dbDataBase1
  custId.DataField = "客戶編號"
  Set custName.DataSource = dbDataBase1
  custName.DataField = "客戶姓名"

  Set receiveItem(0).DataSource = dbDataBase4
  receiveItem(0).DataField = "識別碼"
  Set receiveItem(1).DataSource = dbDataBase4
  receiveItem(1).DataField = "客戶編號"
  Set receiveItem(2).DataSource = dbDataBase4
  receiveItem(2).DataField = "客戶姓名"
  Set receiveItem(3).DataSource = dbDataBase4
  receiveItem(3).DataField = "收款日期"
  Set receiveItem(4).DataSource = dbDataBase4
  receiveItem(4).DataField = "收款金額"
  
  Call set_listReceive
  Call set_listCustom
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub
