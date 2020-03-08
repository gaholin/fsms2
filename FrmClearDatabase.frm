VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmClearDateBase 
   BorderStyle     =   1  '單線固定
   Caption         =   "清除歷史資料"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   Icon            =   "FrmClearDatabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "FrmClearDatabase.frx":0E42
   ScaleHeight     =   5055
   ScaleWidth      =   9495
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   480
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   2640
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Label Label1 
         Caption         =   "資料庫刪除中...請稍後"
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
         Left            =   720
         TabIndex        =   12
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.CheckBox ClearDataBase 
      Caption         =   "清除系統資料"
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
      Index           =   8
      Left            =   5640
      TabIndex        =   10
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CheckBox ClearDataBase 
      Caption         =   "清除稅率資料"
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
      Index           =   7
      Left            =   840
      TabIndex        =   9
      Top             =   1680
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc dbDataBase1 
      Height          =   330
      Left            =   360
      Top             =   120
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
   Begin VB.CommandButton ClearCannel 
      Caption         =   "取消"
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
      Left            =   5160
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton ClearEnter 
      Caption         =   "確定"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CheckBox ClearDataBase 
      Caption         =   "清除前月結餘資料"
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
      Index           =   6
      Left            =   3240
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CheckBox ClearDataBase 
      Caption         =   "清除過帳收款資料"
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
      Index           =   5
      Left            =   5640
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CheckBox ClearDataBase 
      Caption         =   "清除過帳交易資料"
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
      Index           =   4
      Left            =   5640
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox ClearDataBase 
      Caption         =   "清除收款資料"
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
      Index           =   3
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CheckBox ClearDataBase 
      Caption         =   "清除交易資料"
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
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox ClearDataBase 
      Caption         =   "清除魚貨資料"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CheckBox ClearDataBase 
      Caption         =   "清除客戶資料"
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
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "FrmClearDateBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearCannel_Click()
  Unload Me
End Sub

Private Sub ClearDataBase_Click(Index As Integer)
  Dim i As Integer
  Dim flag As Integer
  flag = 0
  For i = 0 To 8
    If ClearDataBase(i).Value = 1 Then
      flag = 1
    End If
  Next
  If flag = 1 Then
    ClearEnter.Enabled = True
  Else
    ClearEnter.Enabled = False
  End If
End Sub

Private Sub ClearFunc()
  Dim TmpStr As String
  dbDataBase1.ConnectionString = database_string
  dbDataBase1.CommandType = adCmdTable
  If ClearDataBase(0).Value = 1 Then
    dbDataBase1.RecordSource = "客戶資料表"
    dbDataBase1.Refresh
    Do Until dbDataBase1.Recordset.EOF
      dbDataBase1.Recordset.Delete
      dbDataBase1.Recordset.MoveNext
    Loop
  End If
  If ClearDataBase(1).Value = 1 Then
    dbDataBase1.RecordSource = "魚貨資料表"
    dbDataBase1.Refresh
    Do Until dbDataBase1.Recordset.EOF
      dbDataBase1.Recordset.Delete
      dbDataBase1.Recordset.MoveNext
    Loop
  End If
  If ClearDataBase(2).Value = 1 Then
    dbDataBase1.RecordSource = "交易資料表"
    dbDataBase1.Refresh
    Do Until dbDataBase1.Recordset.EOF
      dbDataBase1.Recordset.Delete
      dbDataBase1.Recordset.MoveNext
    Loop
  End If
  If ClearDataBase(3).Value = 1 Then
    dbDataBase1.RecordSource = "收款資料表"
    dbDataBase1.Refresh
    Do Until dbDataBase1.Recordset.EOF
      dbDataBase1.Recordset.Delete
      dbDataBase1.Recordset.MoveNext
    Loop
  End If
  If ClearDataBase(4).Value = 1 Then
    dbDataBase1.RecordSource = "過帳交易資料表"
    dbDataBase1.Refresh
    Do Until dbDataBase1.Recordset.EOF
      dbDataBase1.Recordset.Delete
      dbDataBase1.Recordset.MoveNext
    Loop
  End If
  If ClearDataBase(5).Value = 1 Then
    dbDataBase1.RecordSource = "過帳收款資料表"
    dbDataBase1.Refresh
    Do Until dbDataBase1.Recordset.EOF
      dbDataBase1.Recordset.Delete
      dbDataBase1.Recordset.MoveNext
    Loop
  End If
  If ClearDataBase(6).Value = 1 Then
    dbDataBase1.RecordSource = "結算資料表"
    dbDataBase1.Refresh
    Do Until dbDataBase1.Recordset.EOF
      dbDataBase1.Recordset.Delete
      dbDataBase1.Recordset.MoveNext
    Loop
  End If
  If ClearDataBase(7).Value = 1 Then
    dbDataBase1.RecordSource = "稅率資料表"
    dbDataBase1.Refresh
    Do Until dbDataBase1.Recordset.EOF
      dbDataBase1.Recordset.Fields("稅率") = "1.00"
      dbDataBase1.Recordset.Fields("傳票") = "0"
      dbDataBase1.Recordset.Fields("籠子") = "0"
      dbDataBase1.Recordset.MoveNext
    Loop
  End If
  If ClearDataBase(8).Value = 1 Then
    If vYY < 1990 Then
      TmpStr = CStr(vYY + 1911) & "/" & vMM & "/" & 1
    Else
      TmpStr = vYY & "/" & vMM & "/" & 1
    End If
    dbDataBase1.RecordSource = "系統資料表"
    dbDataBase1.Refresh
    Do Until dbDataBase1.Recordset.EOF
      dbDataBase1.Recordset.Fields("密碼") = "1234"
      dbDataBase1.Recordset.Fields("前次工作日期") = TmpStr
      dbDataBase1.Recordset.Fields("前次結算日期") = TmpStr
      dbDataBase1.Recordset.MoveNext
    Loop
  End If
  Unload Me
End Sub

Private Sub ClearEnter_Click()
  Dim i As Integer
  For i = 0 To 8
    ClearDataBase(i).Enabled = False
  Next
  ClearEnter.Enabled = False
  ClearCannel.Enabled = False
  Frame1.Visible = True
  Timer1.Enabled = True
End Sub

Private Sub Form_Load()
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200

End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True

End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Call ClearFunc
End Sub
