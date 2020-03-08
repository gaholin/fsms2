VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCondition 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶漁獲條件搜尋"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11040
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame1 
      Caption         =   "條件輸入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3375
      Begin VB.OptionButton Option1 
         Caption         =   "漁獲查詢"
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
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "客戶查詢"
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
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtCondition 
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
         Left            =   1560
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Find_Database 
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
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "客戶名稱"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid listCustom 
      Height          =   7095
      Left            =   4080
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
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
   Begin MSAdodcLib.Adodc dbDataBase1 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Caption         =   "Adodc1"
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
Attribute VB_Name = "FrmCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OptMode As Integer
Private Sub set_listFish()
  listCustom.Columns.item(0).Width = 1600
  listCustom.Columns.item(1).Width = 1600
  listCustom.Columns.item(2).Width = 1600
End Sub
Private Sub Database_Update()
  Dim cmd As String
  If OptMode = 0 Then
    cmd = "SELECT 客戶資料表.客戶編號, 客戶資料表.客戶姓名, 客戶資料表.性別 From 客戶資料表 WHERE (((客戶資料表.客戶姓名) Like ""%" & txtCondition & "%""));"
    dbDataBase1.ConnectionString = database_string
    dbDataBase1.CommandType = adCmdText
    dbDataBase1.RecordSource = cmd
  Else
    cmd = "SELECT 魚貨資料表.魚貨代號, 魚貨資料表.魚貨名稱, 魚貨資料表.魚貨單位 From 魚貨資料表 WHERE (((魚貨資料表.魚貨名稱) Like ""%" & txtCondition & "%""));"
    dbDataBase1.ConnectionString = database_string
    dbDataBase1.CommandType = adCmdText
    dbDataBase1.RecordSource = cmd
  End If
  dbDataBase1.Refresh
  Set listCustom.DataSource = dbDataBase1
  Call set_listFish
End Sub


Private Sub Find_Database_Click()
  Call Database_Update
End Sub

Private Sub Form_Load()
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  OptMode = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
  If Option1.item(0).value = True Then
    Label1.Caption = "客戶名稱"
    OptMode = 0
  Else
    Label1.Caption = "漁獲名稱"
    OptMode = 1
  End If
End Sub
