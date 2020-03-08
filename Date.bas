Attribute VB_Name = "Date"
Option Explicit
Public vYY As Integer
Public vMM As Integer
Public vDD As Integer
'Public dcode As Long
Public CalendarValue As Date
Public CheckMode As Integer
Public Const StrAppendLeft = 1
Public Const StrAppendRight = 2
Public FSMS_file As String
Public fsms_name As String
Public database_string As String
Public AccountDate As String
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long

Const LOCALE_SLONGDATE = &H20
Const LOCALE_SSHORTDATE = &H1F
Const LOCALE_STIME = &H1E
Const LOCALE_ICALENDARTYPE = &H1009

Public Function appendstr(data As String, num As Integer) As String
  appendstr = String(num - Len(data), "0") & data
End Function
Public Function update_date() As String
  Dim date_string As String
  
  If vYY < 100 Then
    update_date = "0" & Trim(str(vYY))
  Else
    update_date = Trim(str(vYY))
  End If
  If vMM < 10 Then
    update_date = update_date & "/0" & Trim(str(vMM))
  Else
    update_date = update_date & "/" & Trim(str(vMM))
  End If
  If vDD < 10 Then
    update_date = update_date & "/0" & Trim(str(vDD))
  Else
    update_date = update_date & "/" & Trim(str(vDD))
  End If
  'dcode = vYY * 65536 + vMM * 256 + vDD
End Function


Public Function get_date(vDateCode As Long) As String
  Dim vYear As Integer
  Dim vMonth As Integer
  Dim vDay As Integer
  vYear = (vDateCode / 65536) And &HFFFF
  vMonth = (vDateCode / 256) And &HFF
  vDay = vDateCode And &HFF
  
  If vYear < 100 Then
    get_date = "0" & Trim(str(vYear))
  Else
    get_date = Trim(str(vYear))
  End If
  If vMonth < 10 Then
    get_date = get_date & "/0" & Trim(str(vMonth))
  Else
    get_date = get_date & "/" & Trim(str(vMonth))
  End If
  If vDay < 10 Then
    get_date = get_date & "/0" & Trim(str(vDay))
  Else
    get_date = get_date & "/" & Trim(str(vDay))
  End If
End Function


Public Sub compress_mdb()
On Error Resume Next
  Dim strFile As String
  Dim oAccess As Object
  Set oAccess = CreateObject("Access.Application")
  oAccess.CompactRepair "D:\db1_2.mdb", "D:\db2_2.mdb", True
  oAccess.Quit
  Set oAccess = Nothing
  Kill "D:\db1_2.mdb"
  Name "D:\db2_2.mdb" As "D:\db1_2.mdb"
End Sub


Public Function StrAppendSpace(str As String, str_num As Integer, mode As String) As String
  Dim size As Integer
  Dim num As Integer
  size = LenB(StrConv(str, vbFromUnicode))
  num = str_num
  If num < size Then
    num = size
  End If
  If mode = StrAppendLeft Then
    StrAppendSpace = str & String(num - size, " ")
  ElseIf mode = StrAppendRight Then
    StrAppendSpace = String(num - size, " ") & str
  Else
    StrAppendSpace = str
  End If
End Function


Public Function StrFraction(str As String, num As Integer) As String
  Dim data1 As Double
  Dim data2 As String
  Dim mul As Integer
  Dim data As String
  Dim vStr As Integer
  If num = 1 Then
    mul = 10
  ElseIf num = 2 Then
    mul = 100
  ElseIf num = 3 Then
    mul = 1000
  End If
  If str = "" Then
    StrFraction = ""
  ElseIf str = "0" Then
    StrFraction = ""
  Else
    data1 = CDbl(str)
    data2 = CStr((Abs(data1) * mul) Mod mul)
    data = String(num - Len(data2), "0") & data2
    StrFraction = Fix(data1) & "." & data
  End If
End Function

Public Function DC2PC(str As Date) As String
  Dim tmpY As Integer
  Dim tmpM As String
  Dim tmpD As String
  If Year(str) > 1900 Then
    tmpY = Year(str) - 1911
  Else
    tmpY = Year(str)
  End If
  'tmpY = Year(str)
  tmpM = Month(str)
  tmpD = Day(str)
  DC2PC = tmpY & "/" & String(2 - Len(tmpM), "0") & tmpM & "/" & String(2 - Len(tmpD), "0") & tmpD
End Function


Public Function PC2DC(str As String) As String
  Dim tmpY As Integer
  Dim tmpM As Integer
  Dim tmpD As Integer
  Dim Index As Integer
  Index = InStr(str, "/")
  PC2DC = (Mid(str, 1, Index - 1) + 1911) & Mid(str, Index, Len(str))
  'PC2DC = (Mid(str, 1, Index - 1)) & Mid(str, Index, Len(str))
End Function


Public Sub ADYear()
  Dim dwLCID As Long
  dwLCID = GetSystemDefaultLCID
  SetLocaleInfo dwLCID, LOCALE_ICALENDARTYPE, "1" '4中華民國曆，&1西元中文曆
  SetLocaleInfo dwLCID, LOCALE_SSHORTDATE, "yyyy/MM/dd" ' 短日期格式
End Sub
Public Sub RCYear()
  Dim dwLCID As Long
  dwLCID = GetSystemDefaultLCID
  SetLocaleInfo dwLCID, LOCALE_ICALENDARTYPE, "4" '4中華民國曆，&1西元中文曆
  SetLocaleInfo dwLCID, LOCALE_SSHORTDATE, "yyy/MM/dd" ' 短日期格式
End Sub

Public Sub delay1ms(X As Long)
  Dim i, j As Long
  For i = 1 To X
  For j = 0 To 50500
  Next j
  Next i
End Sub

Public Function fsmsfile() As String
  Dim tmpY As Integer
  Dim tmpM As String
  Dim tmpD As String
  Dim NowDate As Date
  NowDate = DateValue(Now)
  If Year(NowDate) > 1900 Then
    tmpY = Year(NowDate) - 1911
  Else
    tmpY = Year(NowDate)
  End If
  tmpM = Month(NowDate)
  tmpD = Day(NowDate)
  If tmpY >= 100 Then
    fsmsfile = "fsms_" & tmpY & String(2 - Len(tmpM), "0") & tmpM & String(2 - Len(tmpD), "0") & tmpD & ".mdb"
  Else
    fsmsfile = "fsms_" & "0" & tmpY & String(2 - Len(tmpM), "0") & tmpM & String(2 - Len(tmpD), "0") & tmpD & ".mdb"
  End If
End Function
