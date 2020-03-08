VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmPostingDeal 
   BackColor       =   &H80000004&
   Caption         =   "交易資料登錄"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   9225
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox fishName 
         DataField       =   "魚貨名稱"
         DataSource      =   "dbFish"
         Height          =   270
         Left            =   1080
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox fishId 
         DataField       =   "魚貨代號"
         DataSource      =   "dbFish"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox custName 
         DataField       =   "客戶姓名"
         DataSource      =   "dbCustom"
         Height          =   270
         Left            =   1080
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox custId 
         DataField       =   "客戶編號"
         DataSource      =   "dbCustom"
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox fishUnit 
         DataField       =   "魚貨單位"
         DataSource      =   "dbFish"
         Height          =   270
         Left            =   3120
         TabIndex        =   8
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox dealDate 
         DataField       =   "交易日期"
         DataSource      =   "dbDeal"
         Height          =   270
         Left            =   2040
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TaxBasket 
         DataField       =   "籠子"
         DataSource      =   "dbTax"
         Height          =   270
         Left            =   2040
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TaxSummons 
         DataField       =   "傳票"
         DataSource      =   "dbTax"
         Height          =   270
         Left            =   1080
         TabIndex        =   5
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TaxRate 
         DataField       =   "稅率"
         DataSource      =   "dbTax"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
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
      MaxLength       =   3
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc dbDeal 
      Height          =   330
      Left            =   7560
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "交易資料表"
      Caption         =   "dbDeal"
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
   Begin MSAdodcLib.Adodc dbFish 
      Height          =   330
      Left            =   5760
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
      RecordSource    =   "魚貨查詢"
      Caption         =   "dbFish"
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
   Begin MSAdodcLib.Adodc dbCustom 
      Height          =   330
      Left            =   3840
      Top             =   120
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
      Caption         =   "dbCustom"
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
   Begin MSAdodcLib.Adodc dbFindDeal 
      Height          =   330
      Left            =   3840
      Top             =   480
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"FrmPostingDeal.frx":0000
      Caption         =   "dbDeal"
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
   Begin MSAdodcLib.Adodc dbTax 
      Height          =   330
      Left            =   5760
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "稅率資料表"
      Caption         =   "dbTax"
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPostingDeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
