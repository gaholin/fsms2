VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmCheckIn 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���/���ڹL�b�@�~"
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
   StartUpPosition =   3  '�t�ιw�]��
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
         Caption         =   "�L�b���еy��"
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
   Begin VB.Frame Frame1 
      Caption         =   "�]�w�d��"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Caption         =   "����u��"
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
         Left            =   1440
         TabIndex        =   15
         Top             =   1560
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optSort 
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
         Index           =   0
         Left            =   2880
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "�C�L"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Caption         =   "�L�b"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
            Name            =   "�s�ө���"
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
            Name            =   "�s�ө���"
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
            Name            =   "�s�ө���"
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
            Name            =   "�s�ө���"
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
         Caption         =   "����϶�"
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
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "�ƦC�覡"
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
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "��"
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
         Left            =   3240
         TabIndex        =   9
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "�Ȥ�϶�                         ��"
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
            Caption         =   "����L�b�@�~"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���ڹL�b�@�~"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�w�L�b�Х�����Ӫ�"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�w�L�b�Ц��ک��Ӫ�"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
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
Private Const PREORT_RECEIVE_HEADER2 = "|   ��      ��   |  ��    ��  | �� �� �� �B |"
Private Const PREORT_RECEIVE_HEADER3 = "============================================="
Private Const REPORT_RECEIVE_COLUMN_SPAN = "        "
Private Const REPORT_RECEIVE_ROW_SPAN = "---------------------------------------------"
Private Const PREORT_DEAL_HEADER1 = "============================================================================================================"
Private Const PREORT_DEAL_HEADER2 = "|  ��     ��  | ������ | �� �W �N �� | ��  �q | �� �� |�| �v| �� �� |Ţ �l|����| ��   �B |�� �L| �X   �p |"
Private Const PREORT_DEAL_HEADER3 = "============================================================================================================"
Private Const PREORT_DEAL_ROW_SPAN = "------------------------------------------------------------------------------------------------------------"

Private Sub set_listDeal()
'�Ȥ�s�� , �Ȥ�m�W, ������
'�ѧO�X , ���f�N��, ���f�W��, ���,
'���q , ���, �|�O, �|�v,
'�ǲ� , Ţ�l, ����, ���B,
'��L, �X�p "
  On Error GoTo ErrListDeal
  listDataBase.Columns.item(0).Width = 1000 '�Ȥ�s��
  listDataBase.Columns.item(1).Width = 1000 '�Ȥ�m�W
  listDataBase.Columns.item(2).Width = 1000 '������
  listDataBase.Columns.item(3).Width = 12 '�ѧO�X
  listDataBase.Columns.item(4).Width = 1000 '���f�N��
  listDataBase.Columns.item(5).Width = 1000 '���f�W��
  listDataBase.Columns.item(6).Width = 12 '�|�v
  listDataBase.Columns.item(7).Width = 800 '���q
  listDataBase.Columns.item(8).Width = 800 '���
  listDataBase.Columns.item(9).Width = 800 '�|�O
  listDataBase.Columns.item(10).Width = 12  '�|�v
  listDataBase.Columns.item(11).Width = 12  '�ǲ�
  listDataBase.Columns.item(12).Width = 12  'Ţ�l
  listDataBase.Columns.item(13).Width = 12  '����
  listDataBase.Columns.item(14).Width = 1000 '���B
  listDataBase.Columns.item(15).Width = 700 '��L
  listDataBase.Columns.item(16).Width = 1000 '�X�p
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
      firstDate = dbDataBase1.Recordset.Fields("������")
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
        item = dbDataBase1.Recordset.Fields("�Ȥ�s��")
        tmp = StrAppendSpace(item, 4, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("�Ȥ�m�W")
        tmp = tmp & StrAppendSpace(item, 9, StrAppendLeft) & " "
        item = dbDataBase1.Recordset.Fields("������")
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
          item = "  �p �p :        �@ " & SumCount & " ��"
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
        item = dbDataBase1.Recordset.Fields("���f�N��")
        tmp = tmp & StrAppendSpace(item, 4, StrAppendLeft) & " "
        item = dbDataBase1.Recordset.Fields("���f�W��")
        tmp = tmp & StrAppendSpace(item, 8, StrAppendLeft) & " "
        item = dbDataBase1.Recordset.Fields("���q")
        SumWeight = SumWeight + CDbl(item)
        TotalWeight = TotalWeight + CDbl(item)
        tmp = tmp & StrAppendSpace(StrFraction(item, 2), 8, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("���")
        tmp = tmp & StrAppendSpace(StrFraction(item, 1), 7, StrAppendRight) & " "
        item = CStr(CDbl(dbDataBase1.Recordset.Fields("�|�v")) - 1)
        
        tmp = tmp & StrAppendSpace(StrFraction(item, 3), 5, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("�ǲ�")
        SumSummons = SumSummons + CDbl(item)
        TotalSummons = TotalSummons + CLng(item)
        tmp = tmp & StrAppendSpace(StrFraction(item, 1), 7, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("Ţ�l")
        SumBasket = SumBasket + CLng(item)
        TotalBasket = TotalBasket + CLng(item)
        tmp = tmp & StrAppendSpace(CStr(Format(CLng(item), "#,###")), 5, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("����")
        tmp = tmp & StrAppendSpace(item, 4, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("���B")
        SumMoney = SumMoney + CLng(item)
        TotalMoney = TotalMoney + CLng(item)
        tmp = tmp & StrAppendSpace(CStr(Format(CLng(item), "#,###,###")), 9, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("��L")
        SumOther = SumOther + CLng(item)
        TotalOther = TotalOther + CLng(item)
        tmp = tmp & StrAppendSpace(CStr(Format(CLng(item), "##,###")), 5, StrAppendRight) & " "
        item = dbDataBase1.Recordset.Fields("�X�p")
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
          item = "  �p �p :        �@ " & SumCount & " ��"
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
        item = "  �X �p :        �@ " & TotalCount & " ��"
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
      firstDate = dbDataBase1.Recordset.Fields("���ڤ��")
      SumMoney = 0
      TotalMoney = 0
      listCnt = 0
      Do Until dbDataBase1.Recordset.EOF
        item = dbDataBase1.Recordset.Fields("�Ȥ�s��")
        tmp = StrAppendSpace(item, 5, StrAppendRight) & "  "
        item = dbDataBase1.Recordset.Fields("�Ȥ�m�W")
        tmp = tmp & StrAppendSpace(item, 10, StrAppendLeft) & "  "
        item = dbDataBase1.Recordset.Fields("���ڤ��")
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
            list_report(total_list) = String(22, " ") & "�p �p : " & StrAppendSpace(Format(SumMoney, "###,###,###"), 13, StrAppendRight) & "  "
            total_list = total_list + 1
            list_report(total_list) = REPORT_RECEIVE_ROW_SPAN
            total_list = total_list + 1
            listCnt = 0
          End If
          SumMoney = 0
        End If
        item = Format(dbDataBase1.Recordset.Fields("���ڪ��B"), "###,###,###")
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
          TotalMoney = TotalMoney + CLng(dbDataBase1.Recordset.Fields("���ڪ��B"))
        End If
        dbDataBase1.Recordset.MoveNext
      Loop
      If total_list < 4094 Then
        If (listCnt Mod 5) <> 0 Then
          list_report(total_list) = REPORT_RECEIVE_ROW_SPAN
          total_list = total_list + 1
        End If
        If optSort(1).value = True Then
          list_report(total_list) = String(22, " ") & "�p �p : " & StrAppendSpace(Format(SumMoney, "###,###,###"), 13, StrAppendRight) & "  "
          total_list = total_list + 1
          list_report(total_list) = REPORT_RECEIVE_ROW_SPAN
          total_list = total_list + 1
        End If
        list_report(total_list) = String(22, " ") & "�X �p : " & StrAppendSpace(Format(TotalMoney, "###,###,###"), 13, StrAppendRight) & "  "
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
    cmd = cmd & "WHERE (((�����ƪ�.�Ȥ�s��) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((�����ƪ�.������) "
      cmd = cmd & "Between #" & checkDate(0) & "# And #" & checkDate(1) & "#) "
    'End If
    cmd = cmd & ") "
  ElseIf CheckMode = 1 Then
    cmd = cmd & "WHERE (((���ڸ�ƪ�.�Ȥ�s��) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((���ڸ�ƪ�.���ڤ��) "
      cmd = cmd & "Between #" & checkDate(0) & "# And #" & checkDate(1) & "#) "
    'End If
    cmd = cmd & ") "
  ElseIf CheckMode = 2 Then
    cmd = cmd & "WHERE (((�L�b�����ƪ�.�Ȥ�s��) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((�L�b�����ƪ�.������) "
      cmd = cmd & "Between #" & checkDate(0) & "# And #" & checkDate(1) & "#) "
    'End If
    cmd = cmd & ") "
  ElseIf CheckMode = 3 Then
    cmd = cmd & "WHERE (((�L�b���ڸ�ƪ�.�Ȥ�s��) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((�L�b���ڸ�ƪ�.���ڤ��) "
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
    cmd = cmd & "HAVING (((�����ƪ�.�Ȥ�s��) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((�����ƪ�.������) "
      cmd = cmd & "Between #" & checkDate(0) & "# And #" & checkDate(1) & "#) "
    'End If
    cmd = cmd & ") "
  ElseIf CheckMode = 1 Then
    cmd = cmd & "HAVING (((���ڸ�ƪ�.�Ȥ�s��) "
    cmd = cmd & "Between '" & start_cid & "' And '" & end_cid & "') "
    'If DateChk.Value = 1 Then
      cmd = cmd & "AND ((���ڸ�ƪ�.���ڤ��) "
      cmd = cmd & "Between #" & checkDate(0) & "# And #" & checkDate(1) & "#) "
    'End If
    cmd = cmd & ") "
  End If
  sum_criteria = cmd
End Sub

Private Sub dispDeal()
  Dim cmd As String
  cmd = "SELECT �����ƪ�.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W, �����ƪ�.������, "
  cmd = cmd & "�����ƪ�.�ѧO�X, �����ƪ�.���f�N��, ���f��ƪ�.���f�W��, �����ƪ�.���, "
  cmd = cmd & "�����ƪ�.���q, �����ƪ�.���, �����ƪ�.�|�O, �|�v��ƪ�.�|�v, "
  cmd = cmd & "�|�v��ƪ�.�ǲ�, �|�v��ƪ�.Ţ�l, �����ƪ�.����, �����ƪ�.���B, "
  cmd = cmd & "�����ƪ�.��L, �����ƪ�.�X�p "
  cmd = cmd & "FROM (�|�v��ƪ� INNER JOIN (���f��ƪ� INNER JOIN �����ƪ� ON "
  cmd = cmd & "(���f��ƪ�.���f�N�� = �����ƪ�.���f�N��) AND (���f��ƪ�.���f�N�� "
  cmd = cmd & "= �����ƪ�.���f�N��)) ON (�|�v��ƪ�.�ѧO�X = �����ƪ�.�|�O) AND "
  cmd = cmd & "(�|�v��ƪ�.�ѧO�X = �����ƪ�.�|�O)) INNER JOIN �Ȥ��ƪ� ON "
  cmd = cmd & "(�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) AND (�����ƪ�.�Ȥ�s�� "
  cmd = cmd & "= �Ȥ��ƪ�.�Ȥ�s��) "
  cmd = cmd & cid_date_criteria
  'cmd = cmd & "ORDER BY �����ƪ�.�Ȥ�s��, �����ƪ�.������, �����ƪ�.�ѧO�X;"
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY �����ƪ�.�Ȥ�s��, �����ƪ�.������, �����ƪ�.�ѧO�X;"
  Else
  cmd = cmd & "ORDER BY �����ƪ�.������, �����ƪ�.�ѧO�X;"
  End If
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh
  Call set_listDeal
End Sub

Private Sub dispReceive()
  Dim cmd As String
  cmd = "SELECT ���ڸ�ƪ�.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W, ���ڸ�ƪ�.���ڤ��, "
  cmd = cmd & "���ڸ�ƪ�.�ѧO�X, ���ڸ�ƪ�.���ڪ��B "
  cmd = cmd & "FROM �Ȥ��ƪ� INNER JOIN ���ڸ�ƪ� ON �Ȥ��ƪ�.�Ȥ�s�� = ���ڸ�ƪ�.�Ȥ�s�� "
  cmd = cmd & cid_date_criteria
  'cmd = cmd & "ORDER BY ���ڸ�ƪ�.�Ȥ�s��, ���ڸ�ƪ�.���ڤ��, ���ڸ�ƪ�.�ѧO�X;"
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY ���ڸ�ƪ�.�Ȥ�s��, ���ڸ�ƪ�.���ڤ��, ���ڸ�ƪ�.�ѧO�X;"
  Else
  cmd = cmd & "ORDER BY ���ڸ�ƪ�.���ڤ��, ���ڸ�ƪ�.�ѧO�X;"
  End If
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh
End Sub

Private Sub dispChkInDeal()
  Dim cmd As String
  'cmd = "SELECT �L�b�����ƪ�.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W, �L�b�����ƪ�.������, ���f��ƪ�.���f�N��, "
  'cmd = cmd & "�L�b�����ƪ�.�ѧO�X, ���f��ƪ�.���f�W��, �L�b�����ƪ�.���, "
  'cmd = cmd & "�L�b�����ƪ�.���q, �L�b�����ƪ�.���, �L�b�����ƪ�.�|�O, �L�b�����ƪ�.����, "
  'cmd = cmd & "�L�b�����ƪ�.���B, �L�b�����ƪ�.��L, �L�b�����ƪ�.�X�p "
  
  cmd = "SELECT �L�b�����ƪ�.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W, �L�b�����ƪ�.������, "
  cmd = cmd & "�L�b�����ƪ�.�ѧO�X, ���f��ƪ�.���f�N��, ���f��ƪ�.���f�W��, �L�b�����ƪ�.���, "
  cmd = cmd & "�L�b�����ƪ�.���q, �L�b�����ƪ�.���, �L�b�����ƪ�.�|�O, �|�v��ƪ�.�|�v, "
  cmd = cmd & "�|�v��ƪ�.�ǲ�, �|�v��ƪ�.Ţ�l, �L�b�����ƪ�.����, �L�b�����ƪ�.���B, "
  cmd = cmd & "�L�b�����ƪ�.��L, �L�b�����ƪ�.�X�p "
  cmd = cmd & "FROM �|�v��ƪ� INNER JOIN ((�L�b�����ƪ� INNER JOIN �Ȥ��ƪ� ON "
  cmd = cmd & "�L�b�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) INNER JOIN ���f��ƪ� ON "
  cmd = cmd & "�L�b�����ƪ�.���f�N�� = ���f��ƪ�.���f�N��) ON �|�v��ƪ�.�ѧO�X = �L�b�����ƪ�.�|�O "
  cmd = cmd & cid_date_criteria
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY �L�b�����ƪ�.�Ȥ�s��, �L�b�����ƪ�.������, �L�b�����ƪ�.�ѧO�X;"
  Else
  cmd = cmd & "ORDER BY �L�b�����ƪ�.������, �L�b�����ƪ�.�ѧO�X;"
  End If
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh

End Sub
Private Sub dispChkInReceive()
  Dim cmd As String
  cmd = "SELECT �L�b���ڸ�ƪ�.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W, �L�b���ڸ�ƪ�.���ڤ��, "
  cmd = cmd & "�L�b���ڸ�ƪ�.�ѧO�X, �L�b���ڸ�ƪ�.���ڪ��B "
  cmd = cmd & "FROM �L�b���ڸ�ƪ� INNER JOIN �Ȥ��ƪ� ON �L�b���ڸ�ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s�� "
  cmd = cmd & cid_date_criteria
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY �L�b���ڸ�ƪ�.�Ȥ�s��, �L�b���ڸ�ƪ�.���ڤ��, �L�b���ڸ�ƪ�.�ѧO�X;"
  Else
  cmd = cmd & "ORDER BY �L�b���ڸ�ƪ�.���ڤ��, �L�b���ڸ�ƪ�.�ѧO�X;"
  End If
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh
End Sub

Private Sub dispChkInDealSum()
  Dim cmd As String
  cmd = "SELECT �����ƪ�.�Ȥ�s��, �����ƪ�.������, Sum(�����ƪ�.���B) AS ���B���`�p "
  cmd = cmd & "FROM (�|�v��ƪ� INNER JOIN (���f��ƪ� INNER JOIN �����ƪ� ON "
  cmd = cmd & "(���f��ƪ�.���f�N�� = �����ƪ�.���f�N��) AND (���f��ƪ�.���f�N�� = "
  cmd = cmd & "�����ƪ�.���f�N��)) ON (�|�v��ƪ�.�ѧO�X = �����ƪ�.�|�O) AND "
  cmd = cmd & "(�|�v��ƪ�.�ѧO�X = �����ƪ�.�|�O)) INNER JOIN �Ȥ��ƪ� ON "
  cmd = cmd & "(�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) AND (�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) "
  cmd = cmd & "GROUP BY �����ƪ�.�Ȥ�s��, �����ƪ�.������ "
  cmd = cmd & sum_criteria
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY �����ƪ�.�Ȥ�s��, �����ƪ�.������;"
  Else
  cmd = cmd & "ORDER BY �����ƪ�.������, �����ƪ�.�Ȥ�s��;"
  End If
  dbDataBase2.CommandType = adCmdText
  dbDataBase2.RecordSource = cmd
  dbDataBase2.Refresh
End Sub
Private Sub dispChkInReceiveSum()
  Dim cmd As String
  cmd = "SELECT ���ڸ�ƪ�.�Ȥ�s��, ���ڸ�ƪ�.���ڤ��, Sum(���ڸ�ƪ�.���ڪ��B) AS ���ڪ��B���`�p "
  cmd = cmd & "FROM �Ȥ��ƪ� INNER JOIN ���ڸ�ƪ� ON �Ȥ��ƪ�.�Ȥ�s��=���ڸ�ƪ�.�Ȥ�s�� "
  cmd = cmd & "GROUP BY ���ڸ�ƪ�.�Ȥ�s��, ���ڸ�ƪ�.���ڤ�� "
  cmd = cmd & sum_criteria
  If optSort(0).value = True Then
  cmd = cmd & "ORDER BY ���ڸ�ƪ�.�Ȥ�s��, ���ڸ�ƪ�.���ڤ��;"
  Else
  cmd = cmd & "ORDER BY ���ڸ�ƪ�.���ڤ��, ���ڸ�ƪ�.�Ȥ�s��;"
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
  'dbDataBase3.RecordSource = "���ȸ�ƪ�"
  'dbDataBase3.Refresh
  'dbDataBase2.Recordset.MoveFirst
  'Do Until dbDataBase2.Recordset.EOF
  '  dbDataBase3.Recordset.AddNew
  '  dbDataBase3.Recordset.Fields("�Ȥ�s��") = dbDataBase2.Recordset.Fields("�Ȥ�s��")
  '  dbDataBase3.Recordset.Fields("������") = dbDataBase2.Recordset.Fields("������")
  '  dbDataBase3.Recordset.Fields("����Ҧ�") = False
  '  dbDataBase3.Recordset.Fields("������B") = dbDataBase2.Recordset.Fields("���B���`�p")
  '  dbDataBase3.Recordset.Update
  '  dbDataBase2.Recordset.MoveNext
  'Loop
  ' �s�W�L�b�����ƪ�
  dbDataBase3.CommandType = adCmdTable
  dbDataBase3.RecordSource = "�L�b�����ƪ�"
  dbDataBase3.Refresh
  dbDataBase1.Recordset.MoveFirst
  Do Until dbDataBase1.Recordset.EOF
    dbDataBase3.Recordset.AddNew
    dbDataBase3.Recordset.Fields("�Ȥ�s��") = dbDataBase1.Recordset.Fields("�Ȥ�s��")
    dbDataBase3.Recordset.Fields("������") = dbDataBase1.Recordset.Fields("������")
    'dbDataBase3.Recordset.Fields("�ǲ��s��") = dbDataBase1.Recordset.Fields("�ǲ��s��")
    dbDataBase3.Recordset.Fields("���f�N��") = dbDataBase1.Recordset.Fields("���f�N��")
    dbDataBase3.Recordset.Fields("���") = dbDataBase1.Recordset.Fields("���")
    dbDataBase3.Recordset.Fields("���q") = dbDataBase1.Recordset.Fields("���q")
    dbDataBase3.Recordset.Fields("���") = dbDataBase1.Recordset.Fields("���")
    dbDataBase3.Recordset.Fields("�|�O") = dbDataBase1.Recordset.Fields("�|�O")
    dbDataBase3.Recordset.Fields("����") = dbDataBase1.Recordset.Fields("����")
    dbDataBase3.Recordset.Fields("���B") = dbDataBase1.Recordset.Fields("���B")
    dbDataBase3.Recordset.Fields("��L") = dbDataBase1.Recordset.Fields("��L")
    dbDataBase3.Recordset.Fields("�X�p") = dbDataBase1.Recordset.Fields("�X�p")
    dbDataBase3.Recordset.Update
    dbDataBase1.Recordset.MoveNext
  Loop
  ' ���������ƪ�
  dbDataBase3.CommandType = adCmdTable
  dbDataBase3.RecordSource = "�����ƪ�"
  dbDataBase3.Refresh
  dbDataBase1.Recordset.MoveFirst
  Do Until dbDataBase1.Recordset.EOF
    cmd = "�ѧO�X=" & dbDataBase1.Recordset.Fields("�ѧO�X")
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
  'dbDataBase3.RecordSource = "���ȸ�ƪ�"
  'dbDataBase3.Refresh
  'dbDataBase2.Recordset.MoveFirst
  'Do Until dbDataBase2.Recordset.EOF
  '  dbDataBase3.Recordset.AddNew
  '  dbDataBase3.Recordset.Fields("�Ȥ�s��") = dbDataBase2.Recordset.Fields("�Ȥ�s��")
  '  dbDataBase3.Recordset.Fields("������") = dbDataBase2.Recordset.Fields("���ڤ��")
  '  dbDataBase3.Recordset.Fields("����Ҧ�") = True
  '  dbDataBase3.Recordset.Fields("������B") = -dbDataBase2.Recordset.Fields("���ڪ��B���`�p")
  '  dbDataBase3.Recordset.Update
  '  dbDataBase2.Recordset.MoveNext
  'Loop
  ' �s�W�L�b�����ƪ�
  dbDataBase3.CommandType = adCmdTable
  dbDataBase3.RecordSource = "�L�b���ڸ�ƪ�"
  dbDataBase3.Refresh
  dbDataBase1.Recordset.MoveFirst
  Do Until dbDataBase1.Recordset.EOF
    dbDataBase3.Recordset.AddNew
    dbDataBase3.Recordset.Fields("�Ȥ�s��") = dbDataBase1.Recordset.Fields("�Ȥ�s��")
    dbDataBase3.Recordset.Fields("���ڤ��") = dbDataBase1.Recordset.Fields("���ڤ��")
    dbDataBase3.Recordset.Fields("���ڪ��B") = dbDataBase1.Recordset.Fields("���ڪ��B")
    dbDataBase3.Recordset.Update
    dbDataBase1.Recordset.MoveNext
  Loop
  ' ���������ƪ�
  dbDataBase3.CommandType = adCmdTable
  dbDataBase3.RecordSource = "���ڸ�ƪ�"
  dbDataBase3.Refresh
  dbDataBase1.Recordset.MoveFirst
  Do Until dbDataBase1.Recordset.EOF
    cmd = "�ѧO�X=" & dbDataBase1.Recordset.Fields("�ѧO�X")
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
      Text1.Text = Text1.Text & REPORT_START_SPAN1 & "** �� �L �b - �� �� �� �� �� **" & vbCrLf
    ElseIf CheckMode = 2 Then
      Text1.Text = Text1.Text & REPORT_START_SPAN1 & "** �w �L �b - �� �� �� �� �� **" & vbCrLf
    End If
    Text1.Text = Text1.Text & REPORT_START_SPAN1 & "�s��̾�: "
    If optSort(0).value = True Then
      Text1.Text = Text1.Text & "�̽s���O" & vbCrLf
    Else
      Text1.Text = Text1.Text & "�̤���O" & vbCrLf
    End If
    Text1.Text = Text1.Text & REPORT_START_SPAN1 & "�s����: " & now_day & vbCrLf
    'If DateChk.Value = 1 Then
      Text1.Text = Text1.Text & REPORT_START_SPAN1 & "����϶�: " & DC2PC(CDate(checkDate(0))) & " �� " & DC2PC(CDate(checkDate(1))) & vbCrLf
    'End If
    Text1.Text = Text1.Text & REPORT_START_SPAN1 & "��歶��: " & CurrPage & vbCrLf
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
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & "** �� �L �b - �� �� �� �� �� **" & vbCrLf
    ElseIf CheckMode = 3 Then
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & "** �w �L �b - �� �� �� �� �� **" & vbCrLf
    End If
    Text1.Text = Text1.Text & REPORT_START_SPAN2 & "�s��̾�: "
    If optSort(0).value = True Then
      Text1.Text = Text1.Text & "�̽s���O" & vbCrLf
    Else
      Text1.Text = Text1.Text & "�̤���O" & vbCrLf
    End If
    Text1.Text = Text1.Text & REPORT_START_SPAN2 & "�s����: " & now_day & vbCrLf
    'If DateChk.Value = 1 Then
      Text1.Text = Text1.Text & REPORT_START_SPAN2 & "����϶�: " & DC2PC(CDate(checkDate(0))) & " �� " & DC2PC(CDate(checkDate(1))) & vbCrLf
    'End If
    Text1.Text = Text1.Text & REPORT_START_SPAN2 & "��歶��: " & CurrPage & vbCrLf
    If CurrList > ListPage Then
      '����
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
    Else '���
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
  Printer.FontName = "�ө���"
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
  cmd = "SELECT �����ƪ�.�ѧO�X, �����ƪ�.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W, �����ƪ�.������, �����ƪ�.���f�N��, "
  cmd = cmd & "���f��ƪ�.���f�W��, �����ƪ�.���, �����ƪ�.���q, �����ƪ�.���, �����ƪ�.�|�O, �����ƪ�.���B, �����ƪ�.��L, �����ƪ�.�X�p "
  cmd = cmd & "FROM (�|�v��ƪ� INNER JOIN (���f��ƪ� INNER JOIN �����ƪ� ON (���f��ƪ�.���f�N�� = �����ƪ�.���f�N��) "
  cmd = cmd & "AND (���f��ƪ�.���f�N�� = �����ƪ�.���f�N��)) ON (�|�v��ƪ�.�ѧO�X = �����ƪ�.�|�O) AND (�|�v��ƪ�.�ѧO�X = �����ƪ�.�|�O)) "
  cmd = cmd & "INNER JOIN �Ȥ��ƪ� ON (�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) AND (�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) "
  cmd = cmd & "ORDER BY �����ƪ�.�Ȥ�s��, �����ƪ�.������;"
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh
  cmd = "SELECT �����ƪ�.�Ȥ�s��, �����ƪ�.������, Sum(�����ƪ�.���B) AS ���B���`�p "
  cmd = cmd & "FROM (�|�v��ƪ� INNER JOIN (���f��ƪ� INNER JOIN �����ƪ� ON (���f��ƪ�.���f�N�� "
  cmd = cmd & "= �����ƪ�.���f�N��) AND (���f��ƪ�.���f�N�� = �����ƪ�.���f�N��)) ON (�|�v��ƪ�.�ѧO�X "
  cmd = cmd & "= �����ƪ�.�|�O) AND (�|�v��ƪ�.�ѧO�X = �����ƪ�.�|�O)) INNER JOIN �Ȥ��ƪ� ON "
  cmd = cmd & "(�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) AND (�����ƪ�.�Ȥ�s�� = �Ȥ��ƪ�.�Ȥ�s��) "
  cmd = cmd & "GROUP BY �����ƪ�.�Ȥ�s��, �����ƪ�.������ "
  cmd = cmd & "ORDER BY �����ƪ�.������, �����ƪ�.�Ȥ�s��;"
  dbDataBase2.CommandType = adCmdText
  dbDataBase2.RecordSource = cmd
  dbDataBase2.Refresh
  dbDataBase3.CommandType = adCmdTable
  dbDataBase3.RecordSource = "�L�b�����ƪ�"
  dbDataBase3.Refresh
  
  Set listDataBase.DataSource = dbDataBase1
  Call update_disp
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub

