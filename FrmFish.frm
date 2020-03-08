VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmFish 
   BorderStyle     =   1  '單線固定
   Caption         =   "魚貨資料"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14700
   Icon            =   "FrmFish.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   14700
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
      Begin VB.CommandButton PrintFish 
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
         TabIndex        =   32
         Top             =   360
         Value           =   -1  'True
         Width           =   855
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
         TabIndex        =   31
         Top             =   360
         Width           =   855
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
         TabIndex        =   30
         Top             =   360
         Width           =   975
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
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox PreviewFish 
      Height          =   1335
      Left            =   240
      TabIndex        =   27
      Top             =   7680
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
         Left            =   600
         TabIndex        =   26
         Top             =   3840
         Width           =   5055
      End
      Begin VB.Label Label15 
         Caption         =   "1) 請使用滑鼠點選欲刪除之魚貨"
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
         TabIndex        =   25
         Top             =   3480
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
         Left            =   360
         TabIndex        =   24
         Top             =   3120
         Width           =   1215
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
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "1) 輸入漁貨代號, 名稱, 及單位"
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
         TabIndex        =   22
         Top             =   840
         Width           =   4215
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
         TabIndex        =   21
         Top             =   1200
         Width           =   4215
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
         Left            =   360
         TabIndex        =   20
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "1) 請使用滑鼠點選欲修改之魚貨"
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
         TabIndex        =   19
         Top             =   2040
         Width           =   5055
      End
      Begin VB.Label Label10 
         Caption         =   "2) 輸入欲修改之魚貨代號, 名稱, 及單位"
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
         Top             =   2400
         Width           =   4215
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
         Left            =   600
         TabIndex        =   17
         Top             =   2760
         Width           =   4215
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
         Left            =   360
         TabIndex        =   16
         Top             =   4320
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   360
      MouseIcon       =   "FrmFish.frx":0E42
      TabIndex        =   6
      Top             =   1080
      Width           =   6015
      Begin VB.CommandButton fishSave 
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
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton fUnit 
         Caption         =   "台斤"
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
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton fUnit 
         Caption         =   "公斤"
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
         Width           =   975
      End
      Begin VB.TextBox fName 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4200
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox fId 
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
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "魚貨單位"
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
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "魚貨名稱"
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
         Left            =   3000
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "魚貨代號"
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
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog PrintDialog 
      Left            =   360
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame debugFish 
      Caption         =   "Debug-Fish"
      Height          =   975
      Left            =   6480
      TabIndex        =   1
      Top             =   7680
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox fishNo 
         DataField       =   "識別碼"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox fishUnit 
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox fishName 
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox fishId 
         DataField       =   "魚貨代號"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Text3"
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid listFish 
      Bindings        =   "FrmFish.frx":1C84
      Height          =   7215
      Left            =   6720
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
      Left            =   -240
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
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
      Left            =   1680
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\works\fsms\fsms2\fsms.mdb;Persist Security Info=False"
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
End
Attribute VB_Name = "FrmFish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const REPORT_START = "         "
Private Const REPORT_COLUMN_SPAN = "           "
Private Const REPORT_ROW_SPAN = "-------------------------"
Private Const PREORT_HEADER1 = "                 ***       魚 貨 資 料 表       ***"
Private Const PREORT_HEADER2 = "製表日期:"
Private Const PREORT_FISH_HEADER1 = "========================="
Private Const PREORT_FISH_HEADER2 = "|  魚       貨  | 單 位 |"
Private Const REPORT_NUM_LIST = 42
Private OptMode As Integer
Private SkipUpdate As Boolean

Private Sub set_listFish()
  listFish.Columns.item(0).Width = 12
  listFish.Columns.item(1).Width = 1400
  listFish.Columns.item(2).Width = 1400
  listFish.Columns.item(3).Width = 1400
End Sub

Private Sub savebutton_enable()
  Dim unit As String
  If fUnit(0).Value = True Then
    unit = "公斤"
  Else
    unit = "台斤"
  End If
  'If fId <> fishId Or fName <> fishName Or fishUnit <> unit Then
  '  fishSave.Enabled = True
  'Else
  '  fishSave.Enabled = False
  'End If
End Sub
  
Private Sub DealMode_Click(Index As Integer)
  OptMode = Index
  If Index = 0 Then
    fishSave.Visible = False
    Frame1.Enabled = False
    fName.Enabled = True
    fId.Enabled = True
    fUnit(0).Enabled = True
    fUnit(1).Enabled = True
  ElseIf Index = 1 Then
    fishSave.Visible = True
    fishSave.Caption = "新增"
    Frame1.Enabled = True
    fName.Enabled = True
    fId.Enabled = True
    fUnit(0).Enabled = True
    fUnit(1).Enabled = True
    fId = ""
    fName = ""
  ElseIf Index = 2 Then
    fishSave.Visible = True
    fishSave.Caption = "變更"
    Frame1.Enabled = True
    fName.Enabled = True
    fId.Enabled = True
    fUnit(0).Enabled = True
    fUnit(1).Enabled = True
    fId = fishId
    fName = fishName
    If fishUnit.Text = "公斤" Then
      fUnit(0).Value = True
    Else
      fUnit(1).Value = True
    End If
  Else
    fishSave.Visible = True
    fishSave.Caption = "刪除"
    Frame1.Enabled = True
    fName.Enabled = False
    fId.Enabled = False
    fUnit(0).Enabled = False
    fUnit(1).Enabled = False
  End If
End Sub

Private Sub fId_Change()
  Call savebutton_enable
End Sub

Private Sub fId_KeyPress(KeyAscii As Integer)
  If ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = 8)) = False Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub fName_Change()
  Call savebutton_enable
End Sub
Private Sub fUnit_Click(Index As Integer)
  Call savebutton_enable
End Sub

Private Sub fishId_Change()
  If SkipUpdate = False Then
    fId = fishId
  End If
End Sub

Private Sub fishName_Change()
  If SkipUpdate = False Then
   fName = fishName
  End If
End Sub

Private Sub fishUnit_Change()
  If SkipUpdate = False Then
    If fishUnit.Text = "公斤" Then
      fUnit(0).Value = True
    Else
      fUnit(1).Value = True
    End If
  End If
End Sub

Private Sub fishNo_Change()
  'If fishNo = "" Then
  '  fishSave.Caption = "新增"
  'Else
  '  fishSave.Caption = "變更"
  'End If
  'fishSave.Enabled = False
End Sub

Private Sub fishSave_Click()
  Dim cmd As String
  Dim del_database As Integer
  cmd = "識別碼=" & fishNo
  ' 變更
  If OptMode = 2 Then
    If fId <> "" And fName <> "" Then
      If fUnit(0).Value = True Then
        fishUnit = "公斤"
      Else
        fishUnit = "台斤"
      End If
      fishId = String(4 - Len(fId), "0") & fId
      fishName = fName
      With dbDataBase1.Recordset
      .Fields("魚貨代號") = fId
      .Fields("魚貨名稱") = fName
      .Fields("魚貨單位") = fishUnit
      .Update
      End With
    End If
  End If
  ' 新增
  If OptMode = 1 Then
    If fId <> "" And fName <> "" Then
      SkipUpdate = True
      dbDataBase1.Recordset.MoveLast
      dbDataBase1.Recordset.Update
      'delay1ms (500)
      dbDataBase1.Recordset.AddNew
      'delay1ms (500)
      If fUnit(0).Value = True Then
        fishUnit = "公斤"
      Else
        fishUnit = "台斤"
      End If
      fishId = String(4 - Len(fId), "0") & fId
      fishName = fName
      With dbDataBase1.Recordset
      .Fields("魚貨代號") = fId
      .Fields("魚貨名稱") = fName
      .Fields("魚貨單位") = fishUnit
      .Update
      End With
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

Private Sub Form_Load()
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  If fishUnit.Text = "公斤" Then
    fUnit(0).Value = True
  Else
    fUnit(1).Value = True
  End If
  dbDataBase1.ConnectionString = database_string
  dbDataBase1.CommandType = adCmdTable
  dbDataBase1.RecordSource = "魚貨資料表"
  dbDataBase1.Refresh
  
  Set listFish.DataSource = dbDataBase1
  Set fishNo.DataSource = dbDataBase1
  fishNo.DataField = "識別碼"
  Set fishId.DataSource = dbDataBase1
  fishId.DataField = "魚貨代號"
  Set fishName.DataSource = dbDataBase1
  fishName.DataField = "魚貨名稱"
  Set fishUnit.DataSource = dbDataBase1
  fishUnit.DataField = "魚貨單位"
  Call set_listFish
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub

Private Sub PrintFish_Click()
  Dim cmd As String
  Dim FishReport(1000) As String
  Dim FishCnt As Integer
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
  On Error GoTo ErrHandlerPrintFish
  If Year(Now) > 1900 Then
    MakeTime = CStr(Year(Now) - 1911) & "/" & Month(Now) & "/" & Day(Now)
  Else
    MakeTime = Year(Now) & "/" & Month(Now) & "/" & Day(Now)
  End If
  MakeTime = MakeTime & "  " & Format(TimeValue(Now), "hh:mm:ss")
  
  cmd = "SELECT 魚貨資料表.魚貨代號, 魚貨資料表.魚貨名稱, 魚貨資料表.魚貨單位 "
  cmd = cmd & "FROM 魚貨資料表 "
  cmd = cmd & "ORDER BY 魚貨資料表.魚貨代號;"
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
    Printer.ForeColor = &H80000008
    FishCnt = 0
    Do Until dbDataBase2.Recordset.EOF
      item1 = dbDataBase2.Recordset.Fields("魚貨代號")
      item2 = dbDataBase2.Recordset.Fields("魚貨名稱")
      item3 = dbDataBase2.Recordset.Fields("魚貨單位")
      
      CurrentPage = Fix(FishCnt / REPORT_NUM_LIST / 2)
      CurrentColumn = Fix((FishCnt Mod (REPORT_NUM_LIST * 2)) / REPORT_NUM_LIST)
      CurrentRow = FishCnt Mod REPORT_NUM_LIST
      If CurrentColumn = 0 Then
        FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = StrAppendSpace(item1, 5, StrAppendRight) & "  "
        FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & StrAppendSpace(item2, 10, StrAppendLeft) & "  "
        FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & StrAppendSpace(item3, 6, StrAppendLeft)
      Else
        FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & REPORT_COLUMN_SPAN
        FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & StrAppendSpace(item1, 5, StrAppendRight) & "  "
        FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & StrAppendSpace(item2, 10, StrAppendLeft) & "  "
        FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & StrAppendSpace(item3, 6, StrAppendLeft)
      End If
      
      FishCnt = FishCnt + 1
      If FishCnt Mod 6 = 5 Then
        CurrentPage = Fix(FishCnt / REPORT_NUM_LIST / 2)
        CurrentColumn = Fix((FishCnt Mod (REPORT_NUM_LIST * 2)) / REPORT_NUM_LIST)
        CurrentRow = FishCnt Mod REPORT_NUM_LIST
        If CurrentColumn = 0 Then
          FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = REPORT_ROW_SPAN
        Else
          FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & REPORT_COLUMN_SPAN & REPORT_ROW_SPAN
        End If
        FishCnt = FishCnt + 1
      End If
      dbDataBase2.Recordset.MoveNext
    Loop
    If FishCnt Mod 6 <> 0 Then
      CurrentPage = Fix(FishCnt / REPORT_NUM_LIST / 2)
      CurrentColumn = Fix((FishCnt Mod (REPORT_NUM_LIST * 2)) / REPORT_NUM_LIST)
      CurrentRow = FishCnt Mod REPORT_NUM_LIST
      If CurrentColumn = 0 Then
        FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = REPORT_ROW_SPAN
      Else
        FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) = FishReport(CurrentPage * REPORT_NUM_LIST + CurrentRow) & REPORT_COLUMN_SPAN & REPORT_ROW_SPAN
      End If
      FishCnt = FishCnt + 1
    End If
    NumPage = Fix((FishCnt + (REPORT_NUM_LIST * 2) - 1) / REPORT_NUM_LIST / 2)
    For i = 1 To NumPage
      PreviewFish.Text = vbCrLf
      PreviewFish.Text = PreviewFish.Text & REPORT_START & PREORT_HEADER1 & vbCrLf & vbCrLf
      PreviewFish.Text = PreviewFish.Text & REPORT_START & PREORT_HEADER2 & " " & StrAppendSpace(MakeTime, 20, StrAppendLeft) & StrAppendSpace(CStr("#" & i), 30, StrAppendRight) & vbCrLf
      If i <> NumPage Then
        currentlistNum = REPORT_NUM_LIST
        PreviewFish.Text = PreviewFish.Text & REPORT_START & PREORT_FISH_HEADER1 & REPORT_COLUMN_SPAN & PREORT_FISH_HEADER1 & vbCrLf
        PreviewFish.Text = PreviewFish.Text & REPORT_START & PREORT_FISH_HEADER2 & REPORT_COLUMN_SPAN & PREORT_FISH_HEADER2 & vbCrLf
        PreviewFish.Text = PreviewFish.Text & REPORT_START & PREORT_FISH_HEADER1 & REPORT_COLUMN_SPAN & PREORT_FISH_HEADER1 & vbCrLf
      Else
        currentlistNum = FishCnt Mod (REPORT_NUM_LIST * 2)
        PreviewFish.Text = PreviewFish.Text & REPORT_START & PREORT_FISH_HEADER1
        If Fix(currentlistNum / REPORT_NUM_LIST) = 1 Then
          PreviewFish.Text = PreviewFish.Text & REPORT_COLUMN_SPAN & PREORT_FISH_HEADER1
        End If
        PreviewFish.Text = PreviewFish.Text & vbCrLf & REPORT_START & PREORT_FISH_HEADER2
        If Fix(currentlistNum / REPORT_NUM_LIST) = 1 Then
          PreviewFish.Text = PreviewFish.Text & REPORT_COLUMN_SPAN & PREORT_FISH_HEADER2
        End If
        PreviewFish.Text = PreviewFish.Text & vbCrLf & REPORT_START & PREORT_FISH_HEADER1
        If Fix(currentlistNum / REPORT_NUM_LIST) = 1 Then
          PreviewFish.Text = PreviewFish.Text & REPORT_COLUMN_SPAN & PREORT_FISH_HEADER1
        End If
        PreviewFish.Text = PreviewFish.Text & vbCrLf
        If currentlistNum > REPORT_NUM_LIST Then
          currentlistNum = REPORT_NUM_LIST
        End If
      End If
      For j = 0 To currentlistNum - 1
        PreviewFish.Text = PreviewFish.Text & REPORT_START & FishReport((i - 1) * REPORT_NUM_LIST + j) & vbCrLf
      Next
      'For j = currentlistNum To REPORT_NUM_LIST + 2
      '  PreviewFish.Text = PreviewFish.Text & vbCrLf
      'Next
      Printer.Print PreviewFish
      If i <> NumPage Then
        Printer.NewPage
      End If
    Next
    Printer.EndDoc
    'Call PrintRTF(PreviewFish, 400, 400, 400, 400)
    Screen.MousePointer = 0
  End If
ErrHandlerPrintFish:
End Sub
