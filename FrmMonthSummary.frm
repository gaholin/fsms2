VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmMonthSummary 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "��~�J�`��"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   8.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMonthSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11880
   Begin VB.TextBox PrintBuf 
      Height          =   975
      Left            =   7200
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Preview 
      BeginProperty Font 
         Name            =   "�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      HideSelection   =   0   'False
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '��̬Ҧ�
      TabIndex        =   6
      Text            =   "FrmMonthSummary.frx":0E42
      Top             =   1680
      Visible         =   0   'False
      Width           =   10695
   End
   Begin MSComDlg.CommonDialog PrintDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "�̫�@��"
      Enabled         =   0   'False
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
      Index           =   3
      Left            =   5760
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "�U�@��"
      Enabled         =   0   'False
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
      Index           =   2
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "�W�@��"
      Enabled         =   0   'False
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
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "�Ĥ@��"
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton butPrint 
      Caption         =   "�C�L"
      Enabled         =   0   'False
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
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
   Begin MSDataGridLib.DataGrid listBusSummary 
      Height          =   6015
      Left            =   840
      TabIndex        =   7
      Top             =   1560
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10610
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6855
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��~�J�`��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�w�����L"
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
End
Attribute VB_Name = "FrmMonthSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TotalCnt As Integer
Private SummaryList(999) As String
Private TotalPage As Integer
Private CurrentPage As Integer
Private MakeTime As String
Private Const PREORT_SUMMARY_HEADER1 = "==================================="
Private Const PREORT_SUMMARY_HEADER2 = "|   ��   ��   |  ��  ��  �l  �B   |"
Private Const REPORT_SUMMARY_START_SPAN = "  "
Private Const REPORT_SUMMARY_COLUMN_SPAN = "    "
Private Const REPORT_ROW_SPAN = "-----------------------------------"
Private Const NUM_LIST_PAGE = 35
Private Const NUM_LIST_PAGE2 = NUM_LIST_PAGE * 2

Private Sub set_listBusSummary()
  listBusSummary.Columns.item(0).Width = 1400
  listBusSummary.Columns.item(1).Width = 1400
  listBusSummary.Columns.item(2).Width = 1400
  listBusSummary.Columns.item(3).Width = 1400
  listBusSummary.Columns.item(4).Width = 1400
  listBusSummary.Columns.item(5).Width = 1400
End Sub

Private Sub updatePreview()
  Dim tmp As String
  Dim flag As Integer
  Preview.Text = ""
  If TotalCnt <> 0 Then
  
    flag = SetPrintPage(1)
  
  
    Preview.Text = vbCrLf & REPORT_SUMMARY_START_SPAN & FrmMonthSummary.Caption & vbCrLf & vbCrLf
    Preview.Text = Preview.Text & REPORT_SUMMARY_START_SPAN & "�s����:" & MakeTime & vbCrLf
    tmp = getPage(CurrentPage)
    Preview.Text = Preview.Text & tmp
  End If
End Sub
  
Private Function getPage(Index As Integer)
  Dim CurrentCnt As Integer
  Dim sizeJ As Integer
  Dim listOffset As Integer
  Dim i As Integer
  Dim textBuf(60) As String
  Dim page_size As Integer
  Dim page_str As String
  Dim row_size As Integer
  Dim NumListPage As Integer
  
  
  'NumListPage = 120
  NumListPage = NUM_LIST_PAGE2
  
  If Index = TotalPage Then
    CurrentCnt = ((TotalCnt - 1) Mod NumListPage)
  Else
    CurrentCnt = NumListPage - 1
  End If
  listOffset = (Index - 1) * NumListPage
  For i = 0 To CurrentCnt
    If i < NUM_LIST_PAGE Then
      textBuf(i) = REPORT_SUMMARY_START_SPAN & SummaryList(listOffset + i)
    ElseIf i < NUM_LIST_PAGE2 Then
    'Else
      textBuf(i - NUM_LIST_PAGE) = textBuf(i - NUM_LIST_PAGE) & REPORT_SUMMARY_COLUMN_SPAN & SummaryList(listOffset + i)
    Else
      textBuf(i - NUM_LIST_PAGE2) = textBuf(i - NUM_LIST_PAGE2) & REPORT_SUMMARY_COLUMN_SPAN & SummaryList(listOffset + i)
    End If
  Next i
  ' Add span line for two/three column
  If (CurrentCnt Mod 5) <> 4 Then
    If (CurrentCnt < NUM_LIST_PAGE) Then
    ElseIf (CurrentCnt < NUM_LIST_PAGE2) Then
      textBuf(CurrentCnt - (NUM_LIST_PAGE - 1)) = textBuf(CurrentCnt - (NUM_LIST_PAGE - 1)) & REPORT_SUMMARY_COLUMN_SPAN & REPORT_ROW_SPAN
    Else
      textBuf(CurrentCnt - (NUM_LIST_PAGE2 - 1)) = textBuf(CurrentCnt - (NUM_LIST_PAGE2 - 1)) & REPORT_SUMMARY_COLUMN_SPAN & REPORT_ROW_SPAN
    End If
  End If
  If (CurrentCnt < NUM_LIST_PAGE) Then
    page_size = CurrentCnt
    getPage = REPORT_SUMMARY_START_SPAN & PREORT_SUMMARY_HEADER1
    row_size = Len(getPage)
    getPage = getPage & vbCrLf & REPORT_SUMMARY_START_SPAN & PREORT_SUMMARY_HEADER2 & vbCrLf
    getPage = getPage & REPORT_SUMMARY_START_SPAN & PREORT_SUMMARY_HEADER1 & vbCrLf
    row_size = Len(REPORT_SUMMARY_START_SPAN & PREORT_SUMMARY_HEADER1)
  ElseIf (CurrentCnt < NUM_LIST_PAGE2) Then
  'Else
    page_size = NUM_LIST_PAGE - 1
    getPage = REPORT_SUMMARY_START_SPAN & PREORT_SUMMARY_HEADER1 & REPORT_SUMMARY_COLUMN_SPAN & PREORT_SUMMARY_HEADER1
    row_size = Len(getPage)
    getPage = getPage & vbCrLf & REPORT_SUMMARY_START_SPAN & PREORT_SUMMARY_HEADER2 & REPORT_SUMMARY_COLUMN_SPAN & PREORT_SUMMARY_HEADER2 & vbCrLf
    getPage = getPage & REPORT_SUMMARY_START_SPAN & PREORT_SUMMARY_HEADER1 & REPORT_SUMMARY_COLUMN_SPAN & PREORT_SUMMARY_HEADER1 & vbCrLf
  Else
    page_size = NUM_LIST_PAGE - 1
    getPage = REPORT_SUMMARY_START_SPAN & PREORT_SUMMARY_HEADER1 & REPORT_SUMMARY_COLUMN_SPAN & PREORT_SUMMARY_HEADER1
    getPage = getPage & REPORT_SUMMARY_COLUMN_SPAN & PREORT_SUMMARY_HEADER1
    row_size = Len(getPage)
    getPage = getPage & vbCrLf & REPORT_SUMMARY_START_SPAN & PREORT_SUMMARY_HEADER2 & REPORT_SUMMARY_COLUMN_SPAN & PREORT_SUMMARY_HEADER2
    getPage = getPage & REPORT_SUMMARY_COLUMN_SPAN & PREORT_SUMMARY_HEADER2 & vbCrLf
    getPage = getPage & REPORT_SUMMARY_START_SPAN & PREORT_SUMMARY_HEADER1 & REPORT_SUMMARY_COLUMN_SPAN & PREORT_SUMMARY_HEADER1
    getPage = getPage & REPORT_SUMMARY_COLUMN_SPAN & PREORT_SUMMARY_HEADER1 & vbCrLf
  End If
  
  For i = 0 To page_size
    getPage = getPage & textBuf(i) & vbCrLf
    If (i Mod 5) = 4 Then
      getPage = getPage & REPORT_SUMMARY_START_SPAN
      If (i + NUM_LIST_PAGE2 < CurrentCnt) Then
        getPage = getPage & REPORT_ROW_SPAN & REPORT_SUMMARY_COLUMN_SPAN
      End If
      If (i + NUM_LIST_PAGE) <= CurrentCnt Then
        getPage = getPage & REPORT_ROW_SPAN & REPORT_SUMMARY_COLUMN_SPAN
      End If
      getPage = getPage & REPORT_ROW_SPAN & vbCrLf
    End If
  Next
  'page_str = "#" & Index & "/" & TotalPage
  'page_str = StrAppendSpace(page_str, row_size, StrAppendRight) & vbCrLf
  'getPage = getPage & page_str

End Function

Private Sub butPrint_Click()
  Dim i As Integer
  Dim page_str As String
  PrintDialog.CancelError = True
  On Error GoTo ErrHandlerPrintSummary
  PrintDialog.Flags = cdlPDReturnDC + cdlPDPageNums + cdlPDNoSelection + cdlPDDisablePrintToFile + cdlPDAllPages + cdlPDCollate
  PrintDialog.ShowPrinter
  PrintBuf.Text = "'"
  Call SetPrintPage(1)
  Printer.FontName = "�ө���"
  Printer.FontSize = 14
  For i = 1 To TotalPage
    'PrintBuf.Text = vbCrLf & vbCrLf & vbCrLf & REPORT_SUMMARY_START_SPAN & FrmMonthSummary.Caption & vbCrLf & vbCrLf
    'PrintBuf.Text = PrintBuf.Text & REPORT_SUMMARY_START_SPAN & "�s����:" & MakeTime & vbCrLf
    'PrintBuf.Text = PrintBuf.Text & getPage(i)
    'PrintBuf.Text = Replace(PrintBuf.Text, vbCrLf, vbCrLf & "       ")
    PrintBuf.Text = vbCrLf & REPORT_SUMMARY_START_SPAN & FrmMonthSummary.Caption & vbCrLf & vbCrLf
    PrintBuf.Text = PrintBuf.Text & REPORT_SUMMARY_START_SPAN & "�s����:" & MakeTime & vbCrLf
    PrintBuf.Text = PrintBuf.Text & getPage(i)

    Printer.Print PrintBuf.Text
    If i <> TotalPage Then
      Printer.NewPage
    End If
  Next i
  Printer.EndDoc
  'txtPreview.SelStart = 0
  'txtPreview.SelLength = Len(txtPreview)
  'Call PrintRTF(PrintBuf, 400, 400, 400, 400)
  Screen.MousePointer = 0
ErrHandlerPrintSummary:
End Sub

Private Sub cmdPage_Click(Index As Integer)
  If Index = 0 Then
    CurrentPage = 1
  ElseIf Index = 1 Then
    CurrentPage = CurrentPage - 1
  ElseIf Index = 2 Then
    CurrentPage = CurrentPage + 1
  ElseIf Index = 3 Then
    CurrentPage = TotalPage
  End If
  If CurrentPage = 1 Then
    cmdPage(0).Enabled = False
    cmdPage(1).Enabled = False
  Else
    cmdPage(0).Enabled = True
    cmdPage(1).Enabled = True
  End If
  If CurrentPage = TotalPage Then
    cmdPage(2).Enabled = False
    cmdPage(3).Enabled = False
  Else
    cmdPage(2).Enabled = True
    cmdPage(3).Enabled = True
  End If
  'lblPreview = "�w���C�L (" & CurrentPage & "/" & TotalPage & ")"
  Call updatePreview
End Sub


Private Sub TabStrip1_Click()
  Dim i As Integer
  If TabStrip1.SelectedItem.Index = 1 Then
    For i = 0 To 3
      cmdPage(i).Visible = False
    Next
    listBusSummary.Visible = True
    Preview.Visible = False
  Else
    For i = 0 To 3
      cmdPage(i).Visible = True
    Next
    listBusSummary.Visible = False
    Preview.Visible = True
  End If
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
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  Me.Caption = vYY & "�~" & vMM & "��" & vDD & "�� ��~�J�`��"
  If vYY < 1990 Then
    date_str = (vYY + 1911) & "/" & vMM & "/1"
    date_stop = (vYY + 1911) & "/" & vMM & "/" & vDD  'DateAdd("d", -1, date_stop)
  Else
    date_str = vYY & "/" & vMM & "/1"
    date_stop = vYY & "/" & vMM & "/" & vDD 'DateAdd("d", -1, date_stop)
  End If
  now_str = Now
  If Year(now_str) > 1900 Then
    currentYear = Year(now_str) - 1911
  Else
    currentYear = Year(now_str)
  End If
  MakeTime = CStr(currentYear) & "/" & Month(now_str) & "/" & Day(now_str)
  MakeTime = MakeTime & "  " & Format(TimeValue(Now), "hh:mm:ss")
  date_start = date_str
  'date_stop = DateAdd("m", 1, date_str)
  dbDataBase1.ConnectionString = database_string
  cmd = "SELECT ��~�J�`��.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W, Sum(��~�J�`��.������B) "
  cmd = cmd & "AS ����l�B, Sum(��~�J�`��.������B) AS �����P��, Sum(��~�J�`��.���ڪ��B) "
  cmd = cmd & "AS ��������, [����l�B]+[�����P��]-[��������] AS �����l�B "
  cmd = cmd & "FROM �Ȥ��ƪ� INNER JOIN ��~�J�`�� ON �Ȥ��ƪ�.�Ȥ�s�� = ��~�J�`��.�Ȥ�s�� "
  cmd = cmd & "WHERE (((��~�J�`��.������) Between #" & date_start & "# And #" & date_stop & "#)) "
  cmd = cmd & "GROUP BY ��~�J�`��.�Ȥ�s��, �Ȥ��ƪ�.�Ȥ�m�W "
  cmd = cmd & "HAVING ((��~�J�`��.�Ȥ�s��) Between '" & "000" & "' And '" & "999" & "')"
  dbDataBase1.CommandType = adCmdText
  dbDataBase1.RecordSource = cmd
  dbDataBase1.Refresh
  Set listBusSummary.DataSource = dbDataBase1
  Call set_listBusSummary
  TotalCnt = 0
  If dbDataBase1.Recordset.EOF = False Then
    'dbDataBase1.Recordset.MoveFirst
    Do Until dbDataBase1.Recordset.EOF
      cid = dbDataBase1.Recordset.Fields("�Ȥ�s��")
      cName = dbDataBase1.Recordset.Fields("�Ȥ�m�W")
      cMoney = dbDataBase1.Recordset.Fields("�����l�B")
      If cMoney <> 0 Then
        cList = StrAppendSpace(cid, 5, StrAppendRight) & " "
        cList = cList & StrAppendSpace(cName, 8, StrAppendLeft)
        cList = cList & StrAppendSpace(Format(cMoney, "###,###,###"), 19, StrAppendRight)
        SummaryList(TotalCnt) = cList & "  "
        TotalCnt = TotalCnt + 1
      End If
      dbDataBase1.Recordset.MoveNext
    Loop
    dbDataBase1.Recordset.MoveFirst
  End If
  TotalPage = Fix((TotalCnt + NUM_LIST_PAGE2 - 1) / NUM_LIST_PAGE2)
  If TotalCnt = 0 Then
    CurrentPage = 0
  Else
    butPrint.Enabled = True
    cmdPage(0).Enabled = False
    cmdPage(1).Enabled = False
    If (TotalPage > 1) Then
      cmdPage(2).Enabled = True
      cmdPage(3).Enabled = True
    Else
      cmdPage(2).Enabled = False
      cmdPage(3).Enabled = False
    End If
    CurrentPage = 1
  End If
  'lblPreview = "�w���C�L (" & CurrentPage & "/" & TotalPage & ")"
  Call updatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub
