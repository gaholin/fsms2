VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmCustom 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶資料表"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13785
   Icon            =   "FrmCustom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   13785
   StartUpPosition =   3  '系統預設值
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
      Height          =   855
      Left            =   360
      TabIndex        =   28
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton PrintCustom 
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
         Left            =   4680
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton DealMode 
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
         Left            =   3480
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton DealMode 
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
         Left            =   2520
         TabIndex        =   31
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton DealMode 
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
         Left            =   1560
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton DealMode 
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
         Left            =   600
         TabIndex        =   29
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox PreviewCustom 
      Height          =   975
      Left            =   360
      TabIndex        =   27
      Top             =   7560
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Frame Frame2 
      Caption         =   "說明"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   360
      TabIndex        =   15
      Top             =   2640
      Width           =   6015
      Begin VB.Label Label16 
         Caption         =   "2) 點選[刪除]"
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
         Left            =   480
         TabIndex        =   26
         Top             =   3960
         Width           =   5055
      End
      Begin VB.Label Label15 
         Caption         =   "1) 請使用滑鼠點選欲刪除之客戶"
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
         Left            =   480
         TabIndex        =   25
         Top             =   3600
         Width           =   5055
      End
      Begin VB.Label Label14 
         Caption         =   "3. 刪除方式"
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
         TabIndex        =   24
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "4. 排列方式 (以新增順序排列)"
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
         Top             =   4440
         Width           =   5055
      End
      Begin VB.Label Label11 
         Caption         =   "3) 點選[變更]"
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
         Left            =   480
         TabIndex        =   22
         Top             =   2760
         Width           =   4215
      End
      Begin VB.Label Label10 
         Caption         =   "2) 輸入欲修改之客戶編號, 姓名, 及性別"
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
         Left            =   480
         TabIndex        =   21
         Top             =   2400
         Width           =   4215
      End
      Begin VB.Label Label9 
         Caption         =   "1) 請使用滑鼠點選欲修改之客戶"
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
         Left            =   480
         TabIndex        =   20
         Top             =   2040
         Width           =   5055
      End
      Begin VB.Label Label8 
         Caption         =   "2. 修改方式"
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
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "2) 點選[新增]"
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
         TabIndex        =   18
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "1) 輸入客戶編號, 姓名, 及性別"
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
         TabIndex        =   17
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "1. 新增方式"
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
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   6015
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
         Height          =   375
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox cName 
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
         Left            =   4200
         MaxLength       =   8
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton cSex 
         Caption         =   "男"
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
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton cSex 
         Caption         =   "女"
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
         Left            =   2400
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cusSave 
         Caption         =   "變更"
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
         Left            =   4320
         TabIndex        =   7
         Top             =   840
         Width           =   1095
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
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "客戶姓名"
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
         Left            =   3120
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "客戶性別"
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
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog PrintDialog 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame debugCustom 
      Caption         =   "Debug-Custom"
      Height          =   975
      Left            =   6480
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox CusId 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text3"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox CusName 
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox CusSex 
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox cusNo 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid listCustom 
      Height          =   7215
      Left            =   6600
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   12726
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
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc dbDataBase2 
      Height          =   375
      Left            =   1800
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
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
End
Attribute VB_Name = "FrmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const REPORT_START = "      "
Private Const REPORT_COLUMN_SPAN = "             "
Private Const REPORT_ROW_SPAN = "-------------------------"
Private Const REPORT_HEADER1 = "                ***       客 戶 資 料 表       ***"
Private Const REPORT_HEADER2 = "製表日期:"
Private Const REPORT_CUSTOM_HEADER1 = "========================="
Private Const REPORT_CUSTOM_HEADER2 = "|  客       戶  | 姓 別 |"
Private Const REPORT_NUM_LIST = 42
Private OptMode As Integer
Private SkipUpdate As Boolean

Private Sub set_listCustom()
  listCustom.Columns.item(0).Width = 12
  listCustom.Columns.item(1).Width = 1400
  listCustom.Columns.item(2).Width = 1400
  listCustom.Columns.item(3).Width = 1400
End Sub

Private Sub savebutton_enable()
  Dim sex As String
  If cSex(0).value = True Then
    sex = "男"
  Else
    sex = "女"
  End If
  'If cid <> CusId Or cName <> CusName Or CusSex <> sex Then
  '  cusSave.Enabled = True
  'Else
  '  cusSave.Enabled = False
  'End If
End Sub
  
Private Sub cid_Change()
  Call savebutton_enable
End Sub

Private Sub cId_KeyPress(KeyAscii As Integer)
  If ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 8)) = False Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub cName_Change()
  Call savebutton_enable
End Sub

Private Sub cSex_Click(Index As Integer)
  Call savebutton_enable
End Sub

Private Sub CusId_Change()
  If SkipUpdate = False Then
    cid = CusId
  End If
End Sub

Private Sub CusName_Change()
  If SkipUpdate = False Then
    cName = CusName
  End If
End Sub

Private Sub CusSex_Change()
  If SkipUpdate = False Then
    If CusSex = "女" Then
      cSex(1).value = True
    Else
      cSex(0).value = True
    End If
  End If
End Sub

Private Sub cusNo_Change()
  'If cusNo = "" Then
  '  cusSave.Caption = "新增"
  'Else
  '  cusSave.Caption = "變更"
  'End If
  'cusSave.Enabled = False
End Sub

Private Sub cusSave_Click()
  Dim cmd As String
  Dim del_database As Integer
  cmd = "識別碼=" & cusNo
  ' 變更
  If OptMode = 2 Then
    If cid <> "" And cName <> "" Then
      If cSex(0).value = True Then
        CusSex = "男"
      Else
        CusSex = "女"
      End If
      'Set CusSex.DataSource = dbDataBase1
      CusId = String(3 - Len(cid), "0") & cid
      CusName = cName
      'If dbDataBase1.Recordset.Fields("識別碼") = "" Then
      'If cusNo = "" Then
      '  With dbDataBase1.Recordset
      '  .Fields("客戶編號") = CusId
      '  .Fields("客戶姓名") = CusName
      '  .Fields("性別") = CusSex
      '  .Update
      '  End With
      '  'cusNo = dbDataBase1.Recordset.Fields("識別碼")
      '  dbDataBase1.Recordset.MovePrevious
      '  dbDataBase1.Recordset.MoveNext
      'Else
        With dbDataBase1.Recordset
        .Fields("客戶編號") = CusId
        .Fields("客戶姓名") = CusName
        .Fields("性別") = CusSex
        .Update
        End With
      'End If
    End If
  End If
  ' 新增
  If OptMode = 1 Then
    If cid <> "" And cName <> "" Then
      'Do Until dbDataBase1.Recordset.EOF = False
      '  dbDataBase1.Recordset.MoveNext
      'Loop
      SkipUpdate = True
      dbDataBase1.Recordset.MoveLast
      dbDataBase1.Recordset.Update
      dbDataBase1.Recordset.AddNew
      If cSex(0).value = True Then
        CusSex = "男"
      Else
        CusSex = "女"
      End If
      'Set CusSex.DataSource = dbDataBase1
      CusId = String(3 - Len(cid), "0") & cid
      CusName = cName
      dbDataBase1.Recordset.Fields("客戶編號") = CusId
      dbDataBase1.Recordset.Fields("客戶姓名") = cName
      dbDataBase1.Recordset.Fields("性別") = CusSex
      dbDataBase1.Recordset.Update
      dbDataBase1.Recordset.MoveFirst
      dbDataBase1.Recordset.MoveLast
      SkipUpdate = False
    End If
  End If
  If OptMode = 3 Then
    dbDataBase1.Recordset.MoveFirst
    dbDataBase1.Recordset.Find cmd
    If dbDataBase1.Recordset.EOF = False Then
      del_database = MsgBox("確定是否刪除?", vbYesNo, "刪除登錄資料")
      If del_database = vbYes Then
        dbDataBase1.Recordset.Delete
        dbDataBase1.Recordset.Update
        'Timer1.Enabled = True
      End If
    End If
  End If
End Sub

Private Sub DealMode_Click(Index As Integer)
  OptMode = Index
  If Index = 0 Then
    cusSave.Visible = False
    cName.Enabled = False
    cid.Enabled = False
    cSex(0).Enabled = False
    cSex(1).Enabled = False
  ElseIf Index = 1 Then
    cusSave.Visible = True
    cusSave.Caption = "新增"
    Frame1.Enabled = True
    cName.Enabled = True
    cid.Enabled = True
    cSex(0).Enabled = True
    cSex(1).Enabled = True
    cid = ""
    cName = ""
  ElseIf Index = 2 Then
    cusSave.Visible = True
    cusSave.Caption = "變更"
    Frame1.Enabled = True
    cName.Enabled = True
    cid.Enabled = True
    cSex(0).Enabled = True
    cSex(1).Enabled = True
    cid = CusId
    cName = CusName
    If CusSex = "女" Then
      cSex(1).value = True
    Else
      cSex(0).value = True
    End If
  Else
    cusSave.Visible = True
    cusSave.Caption = "刪除"
    Frame1.Enabled = True
    cName.Enabled = False
    cid.Enabled = False
    cSex(0).Enabled = False
    cSex(1).Enabled = False
  End If
End Sub


Private Sub Form_Load()
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  SkipUpdate = False
  If CusSex = "男" Then
    cSex(0).value = True
  Else
    cSex(1).value = True
  End If
  dbDataBase1.ConnectionString = database_string
  dbDataBase1.CommandType = adCmdTable
  dbDataBase1.RecordSource = "客戶資料表"
  dbDataBase1.Refresh
  Set listCustom.DataSource = dbDataBase1
  Set cusNo.DataSource = dbDataBase1
  cusNo.DataField = "識別碼"
  Set CusId.DataSource = dbDataBase1
  CusId.DataField = "客戶編號"
  Set CusName.DataSource = dbDataBase1
  CusName.DataField = "客戶姓名"
  Set CusSex.DataSource = dbDataBase1
  CusSex.DataField = "性別"
  Call set_listCustom
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub

Private Sub PrintCustom_Click()
  Dim cmd As String
  Dim CustomReport(1000) As String
  Dim CustomCnt As Integer
  Dim item1 As String
  Dim item2 As String
  Dim item3 As String
  Dim NumPage As Integer
  Dim CurrentPage As Integer
  Dim CurrentColumn As Integer
  Dim CurrentRow As Integer
  Dim i, j As Integer
  Dim currentlistNum As Integer
  Dim MakeTime As String
  On Error GoTo ErrHandlerPrintCustom
  If Year(Now) > 1900 Then
    MakeTime = CStr(Year(Now) - 1911) & "/" & Month(Now) & "/" & Day(Now)
  Else
    MakeTime = Year(Now) & "/" & Month(Now) & "/" & Day(Now)
  End If
  MakeTime = MakeTime & "  " & Format(TimeValue(Now), "hh:mm:ss")
  
  cmd = "SELECT 客戶資料表.客戶編號, 客戶資料表.客戶姓名, 客戶資料表.性別 "
  cmd = cmd & "FROM 客戶資料表 "
  cmd = cmd & "ORDER BY 客戶資料表.客戶編號;"
  dbDataBase2.ConnectionString = database_string
  dbDataBase2.CommandType = adCmdText
  dbDataBase2.RecordSource = cmd
  dbDataBase2.Refresh
  If dbDataBase2.Recordset.EOF <> True Then
    Call SetPrintPage(1)
    PrintDialog.CancelError = True
    PrintDialog.Flags = cdlPDReturnDC + cdlPDPageNums + cdlPDNoSelection + cdlPDDisablePrintToFile + cdlPDAllPages + cdlPDCollate
    PrintDialog.ShowPrinter
    Call SetPrintPage(1)
    Printer.FontName = "細明體"
    Printer.FontSize = 14
    CustomCnt = 0
    Do Until dbDataBase2.Recordset.EOF
      item1 = dbDataBase2.Recordset.Fields("客戶編號")
      item2 = dbDataBase2.Recordset.Fields("客戶姓名")
      item3 = dbDataBase2.Recordset.Fields("性別")
      
      CurrentPage = Fix(CustomCnt / REPORT_NUM_LIST / 2)
      CurrentColumn = Fix((CustomCnt Mod (REPORT_NUM_LIST * 2)) / REPORT_NUM_LIST)
      CurrentRow = CustomCnt Mod REPORT_NUM_LIST
      If CurrentColumn = 0 Then
        CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = StrAppendSpace(item1, 5, StrAppendRight) & "  "
        CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & StrAppendSpace(item2, 10, StrAppendLeft) & "  "
        CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & StrAppendSpace(item3, 6, StrAppendLeft)
      Else
        CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & REPORT_COLUMN_SPAN
        CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & StrAppendSpace(item1, 5, StrAppendRight) & "  "
        CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & StrAppendSpace(item2, 10, StrAppendLeft) & "  "
        CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & StrAppendSpace(item3, 6, StrAppendLeft)
      End If
      
      CustomCnt = CustomCnt + 1
      If CustomCnt Mod 6 = 5 Then
        CurrentPage = Fix(CustomCnt / REPORT_NUM_LIST / 2)
        CurrentColumn = Fix((CustomCnt Mod (REPORT_NUM_LIST * 2)) / REPORT_NUM_LIST)
        CurrentRow = CustomCnt Mod REPORT_NUM_LIST
        If CurrentColumn = 0 Then
          CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = REPORT_ROW_SPAN
        Else
          CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & REPORT_COLUMN_SPAN & REPORT_ROW_SPAN
        End If
        CustomCnt = CustomCnt + 1
      End If
      dbDataBase2.Recordset.MoveNext
    Loop
    If CustomCnt Mod 6 <> 0 Then
      CurrentPage = Fix(CustomCnt / REPORT_NUM_LIST / 2)
      CurrentColumn = Fix((CustomCnt Mod (REPORT_NUM_LIST * 2)) / REPORT_NUM_LIST)
      CurrentRow = CustomCnt Mod REPORT_NUM_LIST
      If CurrentColumn = 0 Then
        CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = REPORT_ROW_SPAN
      Else
        CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = CustomReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & REPORT_COLUMN_SPAN & REPORT_ROW_SPAN
      End If
      CustomCnt = CustomCnt + 1
    End If
    NumPage = Fix((CustomCnt + (REPORT_NUM_LIST * 2) - 1) / REPORT_NUM_LIST / 2)
    For i = 1 To NumPage
      PreviewCustom = vbCrLf
      PreviewCustom = PreviewCustom & REPORT_START & REPORT_HEADER1 & vbCrLf & vbCrLf
      PreviewCustom = PreviewCustom & REPORT_START & REPORT_HEADER2 & " " & StrAppendSpace(MakeTime, 20, StrAppendLeft) & StrAppendSpace(CStr("#" & i), 30, StrAppendRight) & vbCrLf
      If i <> NumPage Then
        currentlistNum = REPORT_NUM_LIST
        PreviewCustom = PreviewCustom & REPORT_START & REPORT_CUSTOM_HEADER1 & REPORT_COLUMN_SPAN & REPORT_CUSTOM_HEADER1 & vbCrLf
        PreviewCustom = PreviewCustom & REPORT_START & REPORT_CUSTOM_HEADER2 & REPORT_COLUMN_SPAN & REPORT_CUSTOM_HEADER2 & vbCrLf
        PreviewCustom = PreviewCustom & REPORT_START & REPORT_CUSTOM_HEADER1 & REPORT_COLUMN_SPAN & REPORT_CUSTOM_HEADER1 & vbCrLf
      Else
        currentlistNum = CustomCnt Mod (REPORT_NUM_LIST * 2)
        PreviewCustom = PreviewCustom & REPORT_START & REPORT_CUSTOM_HEADER1
        If Fix(currentlistNum / REPORT_NUM_LIST) = 1 Then
          PreviewCustom = PreviewCustom & REPORT_COLUMN_SPAN & REPORT_CUSTOM_HEADER1
        End If
        PreviewCustom = PreviewCustom & vbCrLf & REPORT_START & REPORT_CUSTOM_HEADER2
        If Fix(currentlistNum / REPORT_NUM_LIST) = 1 Then
          PreviewCustom = PreviewCustom & REPORT_COLUMN_SPAN & REPORT_CUSTOM_HEADER2
        End If
        PreviewCustom = PreviewCustom & vbCrLf & REPORT_START & REPORT_CUSTOM_HEADER1
        If Fix(currentlistNum / REPORT_NUM_LIST) = 1 Then
          PreviewCustom = PreviewCustom & REPORT_COLUMN_SPAN & REPORT_CUSTOM_HEADER1
        End If
        PreviewCustom = PreviewCustom & vbCrLf
        If currentlistNum > REPORT_NUM_LIST Then
          currentlistNum = REPORT_NUM_LIST
        End If
      End If
      For j = 0 To currentlistNum - 1
        PreviewCustom = PreviewCustom & REPORT_START & CustomReport((i - 1) * REPORT_NUM_LIST + j) & vbCrLf
      Next
      'For j = currentlistNum To REPORT_NUM_LIST + 2
      '  PreviewCustom = PreviewCustom & vbCrLf
      'Next
      Printer.Print PreviewCustom
      If i <> NumPage Then
        Printer.NewPage
      End If
    Next i
    Printer.EndDoc
    'Call PrintRTF(PreviewCustom, 400, 400, 400, 400)
    Screen.MousePointer = 0
  End If
ErrHandlerPrintCustom:
End Sub

