VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCalendar 
   BorderStyle     =   1  '單線固定
   Caption         =   "日期選擇"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   Icon            =   "FrmCalendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9480
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdDateEnter 
      Caption         =   "確定"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   4200
      Width           =   1815
   End
   Begin MSComCtl2.MonthView Calendar1 
      Height          =   3270
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   5768
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   21364737
      CurrentDate     =   40372
   End
End
Attribute VB_Name = "FrmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Calendar1_DateClick(ByVal DateClicked As Date)
  cmdDateEnter.SetFocus
End Sub

Private Sub cmdDateEnter_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  Calendar1.Value = CalendarValue
End Sub

Private Sub Form_Unload(Cancel As Integer)
  CalendarValue = Calendar1.Value
  vDD = Calendar1.Day
  vMM = Calendar1.Month
  vYY = Calendar1.Year
  If vYY > 1900 Then
    vYY = vYY - 1911
  End If
  FsmsMain.Visible = True
End Sub
