VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmDeal 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�����Ƶn���έק�"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "FrmDeal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   11910
   Begin VB.Frame DeleteWindow 
      Height          =   2415
      Left            =   4320
      TabIndex        =   57
      Top             =   2400
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Label12 
         Caption         =   "��ƧR����...�еy��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   58
         Top             =   1080
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc dbDataBase5 
      Height          =   375
      Left            =   13320
      Top             =   240
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Caption         =   "dbDataBase5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "�ާ@�Ҧ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   56
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton DealMode 
         Caption         =   "�d��"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton DealMode 
         Caption         =   "�s�W"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton DealMode 
         Caption         =   "�ק�"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton DealMode 
         Caption         =   "�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   14640
      Top             =   600
   End
   Begin VB.Frame debugInfo 
      Caption         =   "Debug-Info"
      Height          =   1095
      Left            =   600
      TabIndex        =   46
      Top             =   7800
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox TaxRate 
         Height          =   270
         Left            =   120
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TaxSummons 
         Height          =   270
         Left            =   1080
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TaxBasket 
         Height          =   270
         Left            =   2040
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox fishUnit 
         Height          =   270
         Left            =   2040
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox custId 
         Height          =   270
         Left            =   120
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox custName 
         Height          =   270
         Left            =   1080
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox fishId 
         Height          =   270
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox fishName 
         Height          =   270
         Left            =   1080
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame debugDeal 
      Caption         =   "Debug-Deal"
      Height          =   1095
      Left            =   5280
      TabIndex        =   33
      Top             =   7920
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   10
         Left            =   960
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   13
         Left            =   3480
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   12
         Left            =   2640
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   11
         Left            =   1800
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   9
         Left            =   120
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "dbFindDeal"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   8
         Left            =   3480
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   7
         Left            =   2640
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "dbFindDeal"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   6
         Left            =   1800
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "dbFindDeal"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   5
         Left            =   960
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   4
         Left            =   120
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   3
         Left            =   2640
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   2
         Left            =   1800
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   1
         Left            =   960
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox DealItem 
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox cId 
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame FraDeal 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   17
      Top             =   960
      Width           =   11415
      Begin VB.ComboBox fUnit 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmDeal.frx":0E42
         Left            =   3120
         List            =   "FrmDeal.frx":0E4F
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "����"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox fId 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         MaxLength       =   4
         TabIndex        =   5
         ToolTipText     =   "���W�^��Ȥ�s��"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox fWeight 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3960
         MaxLength       =   8
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox fMoney 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4920
         MaxLength       =   8
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox dTax 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmDeal.frx":0E63
         Left            =   5880
         List            =   "FrmDeal.frx":0E7F
         TabIndex        =   9
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox dDivide 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton DealCmd 
         Caption         =   "�s�W"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10200
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox dOther 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8160
         MaxLength       =   3
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label dSum 
         BorderStyle     =   1  '��u�T�w
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8880
         TabIndex        =   59
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "�X�p"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9240
         TabIndex        =   32
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "��   �f   �N   ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "���q"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   28
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "�|�O"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   26
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "���B"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
      Begin VB.Label fName 
         BorderStyle     =   1  '��u�T�w
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   840
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.Label dMoney 
         BorderStyle     =   1  '��u�T�w
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7200
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "��L"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label dDate 
         BorderStyle     =   1  '��u�T�w
         Caption         =   "Label12"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2040
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid listDeal 
      Height          =   2295
      Left            =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
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
   Begin VB.Frame FraCustom 
      Caption         =   "�Ȥ�/���f/�|�v�d��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   16
      Top             =   4680
      Width           =   11415
      Begin MSDataGridLib.DataGrid listCustom 
         Height          =   2535
         Left            =   240
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
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
      Begin MSDataGridLib.DataGrid listFish 
         Height          =   2535
         Left            =   3360
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
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
      Begin MSDataGridLib.DataGrid listTax 
         Height          =   2535
         Left            =   7440
         TabIndex        =   61
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
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
   End
   Begin MSAdodcLib.Adodc dbDataBase2 
      Height          =   330
      Left            =   11400
      Top             =   240
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
      Caption         =   "dbDataBase2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
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
      Left            =   9480
      Top             =   240
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
         Name            =   "�s�ө���"
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
      Left            =   11280
      Top             =   600
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dbDataBase4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc dbDataBase3 
      Height          =   330
      Left            =   9480
      Top             =   600
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
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc dbDataBase6 
      Height          =   375
      Left            =   13320
      Top             =   600
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Caption         =   "dbDataBase6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "�Ȥ�s��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label cName 
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   30
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "FrmDeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private updateCId As Boolean
Private updateCName As Boolean
Private updateFId As Boolean
Private updateFName As Boolean
Private updateFUnit As Boolean
Private start_flag As Boolean
Private dealItemMode As Boolean
Private OtherMoney As Long
Private ACDate As String

Private Sub set_listDeal()
  listDeal.Columns.item(0).Width = 12
  listDeal.Columns.item(1).Width = 1000  '�Ȥ�s��
  listDeal.Columns.item(2).Width = 1000  '�Ȥ�m�W
  listDeal.Columns.item(3).Width = 1000  '������
  listDeal.Columns.item(4).Width = 12    '���f�N��
  listDeal.Columns.item(5).Width = 1100  '���f�W��
  listDeal.Columns.item(6).Width = 800  '���
  listDeal.Columns.item(7).Width = 800   '���q
  listDeal.Columns.item(8).Width = 900  '���
  listDeal.Columns.item(9).Width = 600   '�|�O
  listDeal.Columns.item(10).Width = 600  '����
  listDeal.Columns.item(11).Width = 1000 '���B
  listDeal.Columns.item(12).Width = 900 '��L
  listDeal.Columns.item(13).Width = 1100 '�X�p
End Sub

Private Sub set_listFish()
  listFish.Columns.item(0).Width = 1000
  listFish.Columns.item(1).Width = 1100
  listFish.Columns.item(2).Width = 1100
End Sub
Private Sub set_listCustom()
  listCustom.Columns.item(0).Width = 1000
  listCustom.Columns.item(1).Width = 1100
End Sub
Private Sub set_listTax()
  listTax.Columns.item(0).Width = 1000
  listTax.Columns.item(1).Width = 800
  listTax.Columns.item(2).Width = 800
  listTax.Columns.item(3).Width = 800
End Sub

Private Sub update_money()
  Dim value As Long
  If fMoney <> "" And fWeight <> "" Then
    value = Int(CDbl(fWeight) * CDbl(fMoney) * CDbl(TaxRate) + 0.5)
    dMoney = CLng(value + TaxSummons + TaxBasket)
  Else
    dMoney = "0"
  End If
  dSum = CLng(dMoney) + OtherMoney
End Sub

Private Sub locked_edit()
  'dId.Locked = True
  fId.Locked = True
  fUnit.Locked = True
  fWeight.Locked = True
  fMoney.Locked = True
  dTax.Locked = True
  dDivide.Locked = True
  dOther.Locked = True
  DealCmd.Visible = False
  listFish.Enabled = False
End Sub
Private Sub unlocked_edit()
  'dId.Locked = False
  fId.Locked = False
  fUnit.Locked = False
  fWeight.Locked = False
  fMoney.Locked = False
  dTax.Locked = False
  dDivide.Locked = False
  dOther.Locked = False
  DealCmd.Visible = True
  listFish.Enabled = True
  listFish.ReBind
  Call set_listFish
End Sub
Private Sub clear_edit()
  'dId = ""
  fId = ""
  fName = ""
  fUnit = "����"
  fWeight = ""
  fMoney = ""
  dTax = "1"
  Call dTax_DataBaseUpdate
  dDivide = ""
  dOther = ""
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
  ElseIf fId = "" Then
    Beep
    fId.SelStart = 0
    fId.SelLength = Len(fId)
    fId.SetFocus
    checkInputVaild = False
  ElseIf fName = "" Then
    Beep
    fId.SelStart = 0
    fId.SelLength = Len(fId)
    fId.SetFocus
    checkInputVaild = False
  ElseIf IsNumeric(fWeight) = False Then
    Beep
    fWeight.SelStart = 0
    fWeight.SelLength = Len(fWeight)
    fWeight.SetFocus
    checkInputVaild = False
  ElseIf IsNumeric(fMoney) = False Then
    Beep
    fMoney.SelStart = 0
    fMoney.SelLength = Len(fMoney)
    fMoney.SetFocus
    checkInputVaild = False
  'ElseIf dDivide = "" Then
  '  Beep
  '  dDivide.SelStart = 0
  '  dDivide.SelLength = Len(dDivide)
  '  dDivide.SetFocus
  '  checkInputVaild = False
  ElseIf dOther = "" Then
    Beep
    dOther.SelLength = Len(dOther)
    dOther.SelStart = 0
    dOther.SetFocus
    checkInputVaild = False
  ElseIf dDivide <> "" Then
    checkInputVaild = True
    If dDivide > 255 Then
      dDivide = 255
      dDivide.SetFocus
      Beep
      checkInputVaild = False
    End If
  Else
    checkInputVaild = True
  End If
End Function
Private Sub displayAllDeal()
  Dim cmd As String
  Dim flag As Boolean
  cmd = "SELECT �����ƪ�.�ѧO�X, �����ƪ�.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W, "
  cmd = cmd & "�����ƪ�.������, �����ƪ�.���f�N��, ���f��ƪ�.���f�W��, "
  cmd = cmd & "�����ƪ�.���, �����ƪ�.���q, �����ƪ�.���, �����ƪ�.����, "
  cmd = cmd & "�����ƪ�.�|�O, �����ƪ�.���B, �����ƪ�.��L, �����ƪ�.�X�p "
  cmd = cmd & "FROM (�|�v��ƪ� INNER JOIN (���f��ƪ� INNER JOIN "
  cmd = cmd & "�����ƪ� ON (���f��ƪ�.���f�N�� = �����ƪ�.���f�N��) "
  cmd = cmd & "AND (���f��ƪ�.���f�N�� = �����ƪ�.���f�N��)) ON "
  cmd = cmd & "(�|�v��ƪ�.�ѧO�X = �����ƪ�.�|�O) AND (�|�v��ƪ�.�ѧO�X = "
  cmd = cmd & "�����ƪ�.�|�O)) INNER JOIN �Ȥ��ƪ� ON (�����ƪ�.�Ȥ�s�� = "
  cmd = cmd & "�Ȥ��ƪ�.�Ȥ�s��) AND (�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) "
  cmd = cmd & "ORDER BY �����ƪ�.�Ȥ�s��, �����ƪ�.������, �����ƪ�.�ѧO�X;"
  dbDataBase4.CommandType = adCmdText
  dbDataBase4.RecordSource = cmd
  'dbDataBase4.Recordset.MoveFirst
  flag = IsEmpty(dbDataBase4.Recordset)
  If flag = False Then
    dbDataBase4.Refresh
  End If
  Call set_listDeal
End Sub
Private Sub displayPartDeal()
  Dim cmd As String
  Dim cid_str As String
  Dim flag As Boolean
  cid_str = appendstr(cid, 3)
  cmd = "SELECT �����ƪ�.�ѧO�X, �����ƪ�.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W, "
  cmd = cmd & "�����ƪ�.������, �����ƪ�.���f�N��, ���f��ƪ�.���f�W��, "
  cmd = cmd & "�����ƪ�.���, �����ƪ�.���q, �����ƪ�.���, �����ƪ�.����, "
  cmd = cmd & "�����ƪ�.�|�O, �����ƪ�.���B, �����ƪ�.��L, �����ƪ�.�X�p "
  cmd = cmd & "FROM (�|�v��ƪ� INNER JOIN (���f��ƪ� INNER JOIN "
  cmd = cmd & "�����ƪ� ON (���f��ƪ�.���f�N�� = �����ƪ�.���f�N��) "
  cmd = cmd & "AND (���f��ƪ�.���f�N�� = �����ƪ�.���f�N��)) ON "
  cmd = cmd & "(�|�v��ƪ�.�ѧO�X = �����ƪ�.�|�O) AND (�|�v��ƪ�.�ѧO�X = "
  cmd = cmd & "�����ƪ�.�|�O)) INNER JOIN �Ȥ��ƪ� ON (�����ƪ�.�Ȥ�s�� = "
  cmd = cmd & "�Ȥ��ƪ�.�Ȥ�s��) AND (�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) "
  cmd = cmd & "WHERE (((�����ƪ�.�Ȥ�s��)='" & cid_str & "'))"
  cmd = cmd & "ORDER BY �����ƪ�.�Ȥ�s��, �����ƪ�.������, �����ƪ�.�ѧO�X;"
  dbDataBase4.CommandType = adCmdText
  dbDataBase4.RecordSource = cmd
  dbDataBase4.Refresh
  Call set_listDeal
End Sub

Private Sub cid_Change()
  Dim cmd As String
  Dim flag As Boolean
  Dim RecordsetExist As Boolean
  flag = False
  cmd = cid
  If cmd <> "" Then
    cmd = appendstr(cid, 3)
    cmd = "�Ȥ�s�� = '" & cmd & "'"
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
  If cid = "" Then
    Call displayAllDeal
    cName = ""
  'ElseIf flag = True Then
  '  Call displayPartDeal
  Else
    Call displayPartDeal
  End If
End Sub

Private Sub cid_LostFocus()
  Dim cmd As String
  Dim flag As Boolean
  flag = False
  cmd = cid
  If cmd <> "" Then
    cmd = appendstr(cid, 3)
    cmd = "�Ȥ�s�� = '" & cmd & "'"
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
  If flag Then
    cid.SelStart = 0
    cid.SelLength = Len(cid)
    Beep
  ElseIf cid <> "" Then
    cid = appendstr(cid, 3)
    cName = custName
  End If
End Sub

Private Sub cId_KeyPress(KeyAscii As Integer)
  'If ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
  '    (KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9) Or (KeyAscii = 8)) = False Then
  If KeyAscii = vbKeyReturn And cid <> "" Then
    KeyAscii = 0
    fId.SelStart = 0
    fId.SelLength = Len(fId)
    fId.SetFocus
  ElseIf ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 8)) = False Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub DealCmd_Click()
  Dim cmd As String
  Dim select_index As String
  Dim del_database As Integer
  Dim vaild As Boolean
  
  select_index = DealItem(0).Text
  cmd = "�ѧO�X=" & select_index
  dealItemMode = False
  If dDivide = "" Then
    dDivide = 1
  End If
  If dOther = "" Then
    dOther = "0"
  End If
  If DealMode(1).value = True Then ' �s�W
    vaild = checkInputVaild
    If vaild Then
      'dealItemMode = False
      With dbDataBase4.Recordset
      .AddNew
      .Fields("�Ȥ�s��") = cid
      .Fields("������") = ACDate
      '.Fields("�ǲ��s��") = dId
      .Fields("���f�N��") = fId
      .Fields("���") = fUnit
      .Fields("���q") = fWeight
      .Fields("���") = fMoney
      .Fields("����") = dDivide
      .Fields("�|�O") = dTax
      .Fields("���B") = dMoney
      .Fields("��L") = dOther
      .Fields("�X�p") = dSum
      .Update
      End With
      Call clear_edit
      'Call update_money
      dMoney = CStr(CLng(TaxSummons) + CLng(TaxBasket))
      dSum = CStr(CLng(TaxSummons) + CLng(TaxBasket))
      'dealItemMode = True
      Call displayPartDeal
      If cmd <> "�ѧO�X=" Then
        dbDataBase4.Recordset.MoveFirst
        dbDataBase4.Recordset.Find cmd
        If dbDataBase4.Recordset.EOF = True Then
          MsgBox ("�{�����~")
        End If
      End If
      dbDataBase5.Refresh
      fId.SelStart = 0
      fId.SelLength = Len(fId)
      fId.SetFocus
    End If
    dealItemMode = True
  ElseIf DealMode(2).value = True Then ' �ק�
    vaild = checkInputVaild
    If vaild Then
      dbDataBase4.Recordset.MoveFirst
      dbDataBase4.Recordset.Find cmd
      If dbDataBase4.Recordset.EOF = True Then
        MsgBox ("�{�����~")
      End If
      'dealItemMode = False
      With dbDataBase4.Recordset
      .Fields("�Ȥ�s��") = cid
      '.Fields("������") = ACDate  '������ק�
      '.Fields("�ǲ��s��") = dId
      .Fields("���f�N��") = fId
      .Fields("���") = fUnit
      .Fields("���q") = fWeight
      .Fields("���") = fMoney
      .Fields("�|�O") = dTax
      .Fields("����") = dDivide
      .Fields("���B") = dMoney
      .Fields("��L") = dOther
      .Fields("�X�p") = dSum
      .Update
      End With
      'dealItemMode = True
      dbDataBase4.Refresh
      dbDataBase5.Refresh
      If cid = "" Then
        Call displayAllDeal
      Else
        Call displayPartDeal
      End If
      dbDataBase4.Recordset.MoveFirst
      dbDataBase4.Recordset.Find cmd
      If dbDataBase4.Recordset.EOF = True Then
        MsgBox ("�{�����~")
      End If
      fId.SelStart = 0
      fId.SelLength = Len(fId)
      fId.SetFocus
    End If
    dealItemMode = True
  ElseIf DealMode(3).value = True Then ' �R��
    If select_index <> "" Then
      dbDataBase4.Recordset.MoveFirst
      dbDataBase4.Recordset.Find cmd
      If dbDataBase4.Recordset.EOF = False Then
        del_database = MsgBox("�T�w�O�_�R��?", vbYesNo, "�R���n�����")
        If del_database = vbYes Then
          dealItemMode = True
          dbDataBase5.Refresh
          dbDataBase5.Recordset.MoveFirst
          dbDataBase5.Recordset.Find cmd
          dbDataBase5.Recordset.Delete
          dbDataBase5.Recordset.Update
          DeleteWindow.Visible = True
          Timer1.Enabled = True
        End If
      End If
    End If
  End If
End Sub


Private Sub DealItem_Change(Index As Integer)
  If dealItemMode = True And dbDataBase4.Recordset.BOF = False And dbDataBase4.Recordset.EOF = False Then
    If DealMode(1).value = False Then
      If Index = 3 Then
        dSum = DealItem(Index)
      ElseIf Index = 4 Then
        fId = DealItem(Index)
      ElseIf Index = 5 Then
        fName = DealItem(Index)
      ElseIf Index = 6 Then
        fUnit = DealItem(Index)
      ElseIf Index = 7 Then
        fWeight = DealItem(Index)
      ElseIf Index = 8 Then
        fMoney = DealItem(Index)
      ElseIf Index = 9 Then
        dTax.Text = DealItem(Index)
        Call dTax_DataBaseUpdate
      ElseIf Index = 10 Then
        dDivide = DealItem(Index)
      ElseIf Index = 11 Then
        dMoney = DealItem(Index)
      ElseIf Index = 12 Then
        dOther = DealItem(Index)
      ElseIf Index = 13 And DealMode(1).value = False Then
        dDate = DC2PC(CDate(DealItem(Index)))
      End If
    End If
  End If
End Sub

Private Sub DealMode_Click(Index As Integer)
  If DealMode(0).value = True Then
    Call locked_edit
  ElseIf DealMode(1).value = True Then
    start_flag = True
    Call unlocked_edit
    Call clear_edit
    DealCmd.Caption = "�s�W"
    dDate = vYY & "/" & vMM & "/" & vDD
    ACDate = CStr(vYY + 1911) & "/" & vMM & "/" & vDD
    fId.SetFocus
  ' �ק�
  ElseIf DealMode(2).value = True Then
    Call unlocked_edit
    fId = DealItem(4)
    fName = DealItem(5)
    DealCmd.Caption = "�ק�"
              
    dSum = DealItem(3)
    fId = DealItem(4)
    fName = DealItem(5)
    fUnit = DealItem(6)
    fWeight = DealItem(7)
    fMoney = DealItem(8)
    dTax.Text = DealItem(9)
    Call dTax_DataBaseUpdate
    dDivide = DealItem(10)
    dMoney = DealItem(11)
    dOther = DealItem(12)
    If DealItem(13) <> "" Then
      dDate = DC2PC(CDate(DealItem(13)))
    End If
  Else
    Call locked_edit
    DealCmd.Visible = True
    DealCmd.Caption = "�R��"
  End If
End Sub

'Private Sub dId_KeyPress(KeyAscii As Integer)
'  If ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
'      (KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9) Or (KeyAscii = 8)) = False Then
'    KeyAscii = 0
'    Beep
'  End If
'End Sub

Private Sub dMoney_Change()
  Dim tmp As Integer
  If dMoney <> "" Then
    dSum = CLng(dMoney) + OtherMoney
  Else
    dSum = OtherMoney
  End If
End Sub

Private Sub dOther_Change()
  If dOther = "" Then
    OtherMoney = 0
  ElseIf dOther = "-" Then
    OtherMoney = 0
  ElseIf IsNumeric(dOther) Then
    OtherMoney = CInt(dOther)
  Else
    OtherMoney = 0
  End If
  Call update_money
End Sub

Private Sub dOther_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    DealCmd.SetFocus
    KeyAscii = 0
  'End If
  'If dOther <> "" Then
    'If ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 45) Or _
    '    (KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9) Or (KeyAscii = 8)) = False Then
  ElseIf ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 45) Or (KeyAscii = 8)) = False Then
    KeyAscii = 0
    Beep
  End If
  'End If
End Sub

Private Sub dDivide_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    dOther.SetFocus
    KeyAscii = 0
  ElseIf ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 45) Or (KeyAscii = 8)) = False Then
  'If ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 45) Or _
      '(KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9) Or (KeyAscii = 8)) = False Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub dTax_DataBaseUpdate()
  Dim i As Integer
  dbDataBase3.Recordset.MoveFirst
  For i = 1 To dTax.Text - 1
    dbDataBase3.Recordset.MoveNext
  Next i
End Sub

Private Sub dTax_Click()
  Call dTax_DataBaseUpdate
  Call update_money
End Sub

Private Sub dTax_KeyPress(KeyAscii As Integer)
  Dim lenStr As Integer
  Dim selStrCnt As Integer
  Dim flag As Boolean
  If KeyAscii = vbKeyReturn Then
    dDivide.SetFocus
  End If
  selStrCnt = dTax.SelLength
  If KeyAscii = 8 Then
    selStrCnt = 1
  End If
  lenStr = Len(dTax) - selStrCnt
  If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
     (KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9) Or KeyAscii = 8 Then
    flag = False
  Else
    flag = True
  End If
  If lenStr = 1 Or flag Then
    KeyAscii = 0
  End If
End Sub

Private Sub dTax_LostFocus()
  Dim i As Integer
  If dTax.Text = "" Then
    dTax.Text = "1"
  ElseIf dTax.Text < 1 Or dTax.Text > 8 Then
    dTax.Text = 1
  End If
  dbDataBase3.Recordset.MoveFirst
  For i = 1 To dTax.Text - 1
    dbDataBase3.Recordset.MoveNext
  Next i
  Call update_money
End Sub


Private Sub fId_Change()
  Dim cmd As String
  On Error GoTo ErrHandlerFId1
  If start_flag = True Then
    cmd = "���f�N�� = '" & appendstr(fId.Text, 4) & "'"
    dbDataBase2.Recordset.MoveFirst
    dbDataBase2.Recordset.Find cmd
    fName = fishName
    fUnit = fishUnit
  End If
  'If dbDataBase2.Recordset.EOF = True Then
  'End If
  'If DealMode(0).Value = False And DealMode(3).Value = False Then
  '  fName = fishName
  'End If
ErrHandlerFId1:
End Sub

Private Sub fId_KeyPress(KeyAscii As Integer)
  If ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
      (KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9) Or (KeyAscii = 8)) = False Then
    If KeyAscii = vbKeyReturn And fId <> "" Then
      KeyAscii = 0
      fUnit.SetFocus
    Else
      fId.SetFocus
      KeyAscii = 0
      Beep
    End If
  End If
End Sub

Private Sub fId_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    cid.SelStart = 0
    cid.SelLength = Len(cid.Text)
    cid.SetFocus
  End If
End Sub

Private Sub fId_LostFocus()
  Dim cmd As String
  On Error GoTo ErrHandlerFId2
  If fId <> "" Then
    fId = appendstr(fId, 4)
    cmd = "���f�N�� = '" & fId.Text & "'"
    dbDataBase2.Recordset.MoveFirst
    dbDataBase2.Recordset.Find cmd
    If dbDataBase2.Recordset.EOF = True Then
      MsgBox ("�d�L���f�s��")
    End If
  End If
  start_flag = True
ErrHandlerFId2:
End Sub

Private Sub fishId_Change()
  If updateFId = True Then
    fId = fishId
    updateFId = False
  End If
End Sub

Private Sub fishName_Change()
  If updateFName = True Then
    fName = fishName
    updateFName = False
  End If
End Sub

Private Sub fishUnit_Change()
  If updateFUnit = True Then
    fUnit = fishUnit
    updateFUnit = False
  End If
End Sub

Private Sub fMoney_Change()
 If IsNumeric(fMoney) Then
   Call update_money
 ElseIf fMoney <> "" Then
   Beep
   fMoney.SelStart = 0
   fMoney.SelLength = Len(fMoney)
 End If
End Sub

Private Sub fMoney_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn And fMoney <> "" Then
    dTax.SetFocus
    KeyAscii = 0
  ElseIf IsNumeric(fMoney) And fMoney <> "" Then
    If fMoney <= -100000 Or fMoney >= 100000 Then
      If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
         (KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9) Then
      Beep
      KeyAscii = 0
      End If
    End If
  End If
End Sub

Private Sub fUnit_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    fWeight.SetFocus
  End If
End Sub

Private Sub fWeight_Change()
 If fWeight = "-" Then
 ElseIf IsNumeric(fWeight) Then
   Call update_money
 ElseIf fWeight <> "" Then
   Beep
   fWeight = ""
   fWeight.SelStart = 0
   fWeight.SelLength = Len(fWeight)
 End If
End Sub

Private Sub fWeight_KeyPress(KeyAscii As Integer)
  If IsNumeric(fWeight) Then
    If KeyAscii = vbKeyReturn Then
      fMoney.SetFocus
      KeyAscii = 0
    ElseIf fWeight <= -999 Or fWeight >= 999 Then
      If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
      Beep
      KeyAscii = 0
      End If
    End If
  End If
End Sub

Private Sub listCustom_Click()
  updateCId = True
  updateCName = True
  cid = custId
  cName = custName
End Sub

Private Sub listDeal_Change()
  Call displayPartDeal
End Sub

Private Sub listDeal_Scroll(Cancel As Integer)
  listFish.Refresh
End Sub

Private Sub listFish_Click()
  updateFId = True
  updateFName = True
  updateFUnit = True
  fId = fishId
  fName = fishName
  fUnit = fishUnit
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  dbDataBase4.Refresh
  If cName = "" Then
    Call displayAllDeal
  Else
    Call displayPartDeal
  End If
  If dbDataBase4.Recordset.BOF = False Or dbDataBase4.Recordset.EOF = False Then
    dbDataBase4.Recordset.MoveFirst
  End If
  fId.SetFocus
  DeleteWindow.Visible = False
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim cmd As String
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  start_flag = False
  dDate = vYY & "/" & vMM & "/" & vDD
  Call locked_edit
  cmd = "SELECT �Ȥ��ƪ�.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W "
  cmd = cmd & "From �Ȥ��ƪ� "
  cmd = cmd & "ORDER BY �Ȥ��ƪ�.�Ȥ�s��;"
  dbDataBase1.ConnectionString = database_string
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh
  
  cmd = "SELECT ���f��ƪ�.���f�N��, ���f��ƪ�.���f�W��, ���f��ƪ�.���f��� "
  cmd = cmd & "From ���f��ƪ� "
  cmd = cmd & "ORDER BY ���f��ƪ�.���f�N��;"

  dbDataBase2.ConnectionString = database_string
  dbDataBase2.CommandType = adCmdText
  dbDataBase2.RecordSource = cmd
  dbDataBase2.Refresh
  
  cmd = "SELECT �|�v��ƪ�.�ѧO�X, �|�v��ƪ�.�|�v, �|�v��ƪ�.�ǲ�, �|�v��ƪ�.Ţ�l "
  cmd = cmd & "From �|�v��ƪ� "
  cmd = cmd & "ORDER BY �|�v��ƪ�.�ѧO�X;"

  dbDataBase3.ConnectionString = database_string
  dbDataBase3.CommandType = adCmdText
  dbDataBase3.RecordSource = cmd
  dbDataBase3.Refresh
  
  cmd = "SELECT �����ƪ�.�ѧO�X, �����ƪ�.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W, �����ƪ�.������, �����ƪ�.���f�N��, ���f��ƪ�.���f�W��, �����ƪ�.���, �����ƪ�.���q, �����ƪ�.���, �����ƪ�.�|�O, �����ƪ�.����, �����ƪ�.���B, �����ƪ�.��L, �����ƪ�.�X�p "
  cmd = cmd & "FROM (�|�v��ƪ� INNER JOIN (���f��ƪ� INNER JOIN �����ƪ� ON (���f��ƪ�.���f�N�� = �����ƪ�.���f�N��) AND (���f��ƪ�.���f�N�� = �����ƪ�.���f�N��)) ON (�|�v��ƪ�.�ѧO�X = �����ƪ�.�|�O) AND (�|�v��ƪ�.�ѧO�X = �����ƪ�.�|�O)) INNER JOIN �Ȥ��ƪ� ON (�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) AND (�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) "
  cmd = cmd & "ORDER BY �����ƪ�.�Ȥ�s��, �����ƪ�.������, �����ƪ�.�ѧO�X;"
  dbDataBase4.ConnectionString = database_string
  dbDataBase4.CommandType = adCmdText
  dbDataBase4.RecordSource = cmd
  dbDataBase4.Refresh
  
  
  dbDataBase5.ConnectionString = database_string
  dbDataBase5.CommandType = adCmdTable
  dbDataBase5.RecordSource = "�����ƪ�" '"�����ƪ��`�M"
  dbDataBase5.Refresh
  
  
  cmd = "SELECT �|�v��ƪ�.�ѧO�X, �|�v��ƪ�.�|�v, �|�v��ƪ�.�ǲ�, �|�v��ƪ�.Ţ�l "
  cmd = cmd & "From �|�v��ƪ� "
  cmd = cmd & "ORDER BY �|�v��ƪ�.�ѧO�X;"
  dbDataBase6.ConnectionString = database_string
  dbDataBase6.CommandType = adCmdText
  dbDataBase6.RecordSource = cmd
  dbDataBase6.Refresh
  
  Set listCustom.DataSource = dbDataBase1
  Set listFish.DataSource = dbDataBase2
  Set listDeal.DataSource = dbDataBase4
  Set custId.DataSource = dbDataBase1
  Set listTax.DataSource = dbDataBase6
  Call set_listCustom
  Call set_listDeal
  Call set_listFish
  Call set_listTax
  Call update_money
  custId.DataField = "�Ȥ�s��"
  Set custName.DataSource = dbDataBase1
  custName.DataField = "�Ȥ�m�W"
  
  Set fishId.DataSource = dbDataBase2
  fishId.DataField = "���f�N��"
  Set fishName.DataSource = dbDataBase2
  fishName.DataField = "���f�W��"
  Set fishUnit.DataSource = dbDataBase2
  fishUnit.DataField = "���f���"
  
  Set TaxRate.DataSource = dbDataBase3
  TaxRate.DataField = "�|�v"
  Set TaxSummons.DataSource = dbDataBase3
  TaxSummons.DataField = "�ǲ�"
  Set TaxBasket.DataSource = dbDataBase3
  TaxBasket.DataField = "Ţ�l"
  
  Set DealItem(0).DataSource = dbDataBase4
  DealItem(0).DataField = "�ѧO�X"
  Set DealItem(1).DataSource = dbDataBase4
  DealItem(1).DataField = "�Ȥ�s��"
  Set DealItem(2).DataSource = dbDataBase4
  DealItem(2).DataField = "�Ȥ�m�W"
  Set DealItem(3).DataSource = dbDataBase4
  DealItem(3).DataField = "�X�p"
  Set DealItem(4).DataSource = dbDataBase4
  DealItem(4).DataField = "���f�N��"
  Set DealItem(5).DataSource = dbDataBase4
  DealItem(5).DataField = "���f�W��"
  Set DealItem(6).DataSource = dbDataBase4
  DealItem(6).DataField = "���"
  Set DealItem(7).DataSource = dbDataBase4
  DealItem(7).DataField = "���q"
  Set DealItem(8).DataSource = dbDataBase4
  DealItem(8).DataField = "���"
  Set DealItem(9).DataSource = dbDataBase4
  DealItem(9).DataField = "�|�O"
  Set DealItem(10).DataSource = dbDataBase4
  DealItem(10).DataField = "����"
  Set DealItem(11).DataSource = dbDataBase4
  DealItem(11).DataField = "���B"
  Set DealItem(12).DataSource = dbDataBase4
  DealItem(12).DataField = "��L"
  Set DealItem(13).DataSource = dbDataBase4
  DealItem(13).DataField = "������"
  
  updateCId = False
  updateCName = False
  updateFId = False
  updateFName = False
  updateFUnit = False
  dealItemMode = True
  OtherMoney = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub
