VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�{������"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10950
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   1680
      Picture         =   "FrmHelp.frx":0E42
      ScaleHeight     =   3195
      ScaleWidth      =   6915
      TabIndex        =   2
      Top             =   120
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�T�w"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '�������b
      TabIndex        =   0
      Top             =   3600
      Width           =   9135
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height - 200
  txtHelp = "��Ʈw�B�z�覡" & vbCrLf
  txtHelp = txtHelp & "  1. ��Ʈw�HAccess�覡�B�z, �]�i�H�Q��Access�覡�B�z" & vbCrLf
  txtHelp = txtHelp & "  2. �H�묰��찵�p��" & vbCrLf
  txtHelp = txtHelp & "  3. ����: {�e�뵲��} + {������} - {���리��} = {���뵲�l}" & vbCrLf
  txtHelp = txtHelp & "  4. �G�C����s�������, �����n���ڤW�Ӥ�l��" & vbCrLf
  txtHelp = txtHelp & "  5. �C��|������Φ���, �G�C�饲����J�����ƻP���ڸ��" & vbCrLf
  txtHelp = txtHelp & "  6. �L���b��, �Y�i�d�ߵ��l, �ΦC�L�������" & vbCrLf
  txtHelp = txtHelp & "" & vbCrLf
  txtHelp = txtHelp & "�C��u�@" & vbCrLf
  txtHelp = txtHelp & "1. ��J�C�����Φ��ڴڶ�" & vbCrLf
  txtHelp = txtHelp & "  1) ����[�����Ʒs�W�P�ק�] -> ��J������" & vbCrLf
  txtHelp = txtHelp & "  2) ����[���ڸ�Ʒs�W�P�ק�] -> ��J���ڸ��" & vbCrLf
  txtHelp = txtHelp & "2. �T�w�����ƥ��T�L�~, �N����Φ��ڹL�b" & vbCrLf
  txtHelp = txtHelp & "  1) ����[�����ƹL�b] -> [�L�b]" & vbCrLf
  txtHelp = txtHelp & "  2) ����[���ڸ�ƹL�b] -> [�L�b]" & vbCrLf
  txtHelp = txtHelp & "3. �C�L��~�J�`��" & vbCrLf
  txtHelp = txtHelp & "  1) [��~�J�`��] -> [�C�L]" & vbCrLf
  txtHelp = txtHelp & "4. �C�L��b��" & vbCrLf
  txtHelp = txtHelp & "  1) ��ܩҭn�C�L��b�檺�H��, �i�Q�Τ@�@�s�W, �]�i�H�۰ʷs�W" & vbCrLf
  txtHelp = txtHelp & "  2) ����Ψ����ҭn�C�L���H��" & vbCrLf
  txtHelp = txtHelp & "  3) [��b��C�L] -> [�C�L]" & vbCrLf
  txtHelp = txtHelp & "5. �ƥ����" & vbCrLf
  txtHelp = txtHelp & "  1) ����[��Ƴƥ�]" & vbCrLf
  txtHelp = txtHelp & "  2) �}�Ҧs����|" & vbCrLf
  txtHelp = txtHelp & "  3) ��J�x�s�W��" & vbCrLf
  txtHelp = txtHelp & "  4) [�}��]" & vbCrLf
  txtHelp = txtHelp & "" & vbCrLf
  txtHelp = txtHelp & "�C��u�@" & vbCrLf
  txtHelp = txtHelp & "1. ����e��l�� (����W�Ӥ�l��)" & vbCrLf
  txtHelp = txtHelp & "  1) �Y�n���� 99�~11���, �����N�B�z������99�~12���������@��" & vbCrLf
  txtHelp = txtHelp & "  2) ����[�e�뵲��] -> [����]" & vbCrLf
  txtHelp = txtHelp & "  " & vbCrLf
  txtHelp = txtHelp & "��L" & vbCrLf
  txtHelp = txtHelp & "1. ��l���" & vbCrLf
  txtHelp = txtHelp & "  1) �إ߰򥻸�� (�Ȥ��Ƥγ��f���)" & vbCrLf
  txtHelp = txtHelp & "  2) �إߪ�l�W�뵲�l/��l�L�b������/��l�L�b���ڸ��" & vbCrLf
  txtHelp = txtHelp & "     (�����ϥ�Access�פJ�ο�J���)" & vbCrLf
  txtHelp = txtHelp & "2. ���ƿ��~��,�i�N���e��Ƶ��^�s" & vbCrLf
  txtHelp = txtHelp & "  1) ����[��Ʀ^�s]" & vbCrLf
  txtHelp = txtHelp & "  2) ��ܶ}�Ҹ��|" & vbCrLf
  txtHelp = txtHelp & "  3) ��J�}�ҦW��" & vbCrLf
  txtHelp = txtHelp & "  4) [�}��]" & vbCrLf
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FsmsMain.Visible = True
End Sub
