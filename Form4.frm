VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   10065
   LinkTopic       =   "Form4"
   ScaleHeight     =   8490
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   8400
      TabIndex        =   70
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton CmdPost 
      Caption         =   "Posting"
      Height          =   375
      Left            =   8400
      TabIndex        =   65
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton CmdNext1 
      Caption         =   "Next"
      Height          =   375
      Left            =   8400
      TabIndex        =   64
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton CmdFirst1 
      Caption         =   "First"
      Height          =   375
      Left            =   8400
      TabIndex        =   63
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   8400
      TabIndex        =   62
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame FrameDeduct 
      Caption         =   "Deduction"
      Height          =   4455
      Left            =   4200
      TabIndex        =   38
      Top             =   2760
      Width           =   3735
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   2040
         TabIndex        =   46
         Text            =   "Text21"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   2040
         TabIndex        =   45
         Text            =   "Text22"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   2040
         TabIndex        =   44
         Text            =   "Text23"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   2040
         TabIndex        =   43
         Text            =   "Text24"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   2040
         TabIndex        =   42
         Text            =   "Text25"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   2040
         TabIndex        =   41
         Text            =   "Text26"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   2040
         TabIndex        =   40
         Text            =   "Text27"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   2040
         TabIndex        =   39
         Text            =   "Text28"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label LabTotalDed 
         Caption         =   "LabelTotalDed"
         Height          =   255
         Left            =   2040
         TabIndex        =   59
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label LabelTotal2 
         Caption         =   "Total"
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "G  P  Fund"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   54
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Group Insurance"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "Club Subscription"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label27 
         Caption         =   "House Rent"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   51
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "Vehicle Bill"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "Electricity Bill"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label30 
         Caption         =   "Telephone Bill"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "Others"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   3720
         Width           =   1455
      End
   End
   Begin VB.Frame frameSalary 
      Caption         =   "Salary and Allowance"
      Height          =   4455
      Left            =   240
      TabIndex        =   21
      Top             =   2760
      Width           =   3615
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   29
         Text            =   "Text8"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   28
         Text            =   "Text9"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   27
         Text            =   "Text10"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Text            =   "Text11"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   25
         Text            =   "Text12"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   24
         Text            =   "Text14"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text15 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Text            =   "Text15"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Text            =   "Text16"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label LabelTotalSal 
         Caption         =   "LabelTotalSal"
         Height          =   255
         Left            =   1800
         TabIndex        =   57
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label LabelTotal 
         Caption         =   "Total"
         Height          =   255
         Left            =   480
         TabIndex        =   56
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Basic Salary"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Dearness "
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   480
         TabIndex        =   36
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "House Rent"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Medical"
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   480
         TabIndex        =   34
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Head/Charge"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Recreation"
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Festival"
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   480
         TabIndex        =   31
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Others"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   3720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   9495
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   5
         Text            =   "text7"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text29 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7680
         TabIndex        =   4
         Text            =   "Text29"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label LabGrade 
         Caption         =   "Labgrd"
         Height          =   255
         Left            =   6600
         TabIndex        =   68
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Pay  Grade"
         Height          =   255
         Left            =   4920
         TabIndex        =   67
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label LabelInc 
         Caption         =   "LabelInc"
         Height          =   255
         Left            =   6600
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label LabelScale 
         Caption         =   "LabelScale"
         Height          =   255
         Left            =   6600
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label LabelDep 
         Caption         =   "LabelDep"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label LabelDeg 
         Caption         =   "LabelDeg"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label LabelName 
         Caption         =   "LabelName"
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label LabelId 
         Caption         =   "LabelId"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Id No"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Designation"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Department"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Pay Scale"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4920
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Date of Increment"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Pay Period"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4920
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "From"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4920
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label33 
         Caption         =   "To"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   6720
         TabIndex        =   6
         Top             =   1320
         Width           =   255
      End
   End
   Begin VB.CommandButton CmdPrev1 
      Caption         =   "Previous"
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton CmdLast1 
      Caption         =   "Last"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Labelyear 
      Caption         =   "Labelyear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   69
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LabelPre 
      Caption         =   "PRE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   66
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Labelmonth 
      Caption         =   "Labelmonth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   61
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LabelAmmount 
      Caption         =   "Ammount"
      Height          =   375
      Left            =   6240
      TabIndex        =   60
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label LabeNet 
      Caption         =   "Net Pay ="
      Height          =   375
      Left            =   5040
      TabIndex        =   55
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   " POSTING  OF  SALARY    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbCnn As ADODB.Connection
Dim rsid As ADODB.Recordset
 Dim rsal As ADODB.Recordset
 Dim rstt As ADODB.Recordset
 
 Dim rs As ADODB.Recordset


Private Sub Command1_Click()
Form3.Show
End Sub

Private Sub Command2_Click()
Form1.Show
End Sub
Private Sub showpost()


If rs.BOF Then
MsgBox "There is No Id to be Posted."
Unload Form4
GoTo lst
End If

Dim rtCmd As New ADODB.Command
Set rtCmd.ActiveConnection = dbCnn
rtCmd.CommandText = "returnpost"
rtCmd.CommandType = adCmdStoredProc

rtCmd.Parameters.Append rtCmd.CreateParameter("id", adVarChar, adParamInput, 50, rs!id)

rtCmd.Parameters.Append rtCmd.CreateParameter("nm", adVarChar, adParamOutput, 50)
rtCmd.Parameters.Append rtCmd.CreateParameter("deg", adVarChar, adParamOutput, 50)
rtCmd.Parameters.Append rtCmd.CreateParameter("dep", adVarChar, adParamOutput, 50)
rtCmd.Parameters.Append rtCmd.CreateParameter("scl", adVarChar, adParamOutput, 50)
rtCmd.Parameters.Append rtCmd.CreateParameter("incdt", adVarChar, adParamOutput, 50)
rtCmd.Parameters.Append rtCmd.CreateParameter("gd", adVarChar, adParamOutput, 50)


rtCmd.Execute


'''''''

Dim dr As String, _
hr As Long, m As Long, Hd As Long _
, rec As Long, fs As Long, ot As Long, sum As Long




LabelId = rs!id

LabelName = rtCmd.Parameters("nm")
LabelDeg = rtCmd.Parameters("deg")

LabelDep = rtCmd.Parameters("dep")


LabelScale = rtCmd.Parameters("scl")

LabelInc = rtCmd.Parameters("incdt")

LabGrade = rtCmd.Parameters("gd")



Text7 = Date
Text29 = Date + 30






''''''''''

Text8 = rs!dearness

Text9 = rs!basic

dr = Text8
Text10 = rs!houserent

hr = Text10
Text11 = rs!medical

m = Text11
Text12 = rs!headcharge
Hd = Text12
Text14 = rs!recreation

rec = Text14
Text15 = rs!festival
fs = Text15
Text16 = rs!otherallowance

sum = dr + hr + m + Hd + rec + fs + ot + Text9


LabelTotalSal.Caption = sum















Text21.Text = rs!gpins
'Text21.Enabled = False


Text22.Text = rs!clubsubs
'Text22.Enabled = False

Text23.Text = rs!gpfund
'Text23.Enabled = False

Text24.Text = rs!hrent
'Text24.Enabled = False

Text25.Text = rs!vehibill
'Text25.Enabled = False

Text26.Text = rs!Elecbill
'Text26.Enabled = False

Text27.Text = rs!telbill
'Text27.Enabled = False

Text28.Text = rs!otherdeduct
'Text28.Enabled = False

LabelAmmount.Caption = rs!netpay
'Text21.Enabled = False

LabTotalDed.Caption = rs!totaldeduct


lst:

End Sub

Private Sub showid()
LabelId.Caption = rsid!id

LabelName = rsid!Name
LabelDeg = rsid!Designation
LabelDep = rsid!Department

LabelScale = rsid!Scale
LabelInc = rsid!incDate
LabGrade.Caption = rsid!grade


Text7 = Date
Text29 = Date + 30

End Sub

Private Sub showField()
'Text9 = rs!Basic
Dim dr As String, _
hr As Long, m As Long, Hd As Long _
, rec As Long, fs As Long, ot As Long, sum As Long, basic As Long


Dim gdstr As String

Dim gd As Long

gdstr = rsid!grade

If gdstr = "Grade-1" Then
gd = 400
End If


If gdstr = "Grade-2" Then
gd = 300
End If
If gdstr = "Grade-3" Then
gd = 200
End If
If gdstr = "Grade-4" Then
gd = 100
End If

Dim incdt As Date
Dim di, dp As Long
Dim mi, mp As Long
Dim yi, yp As Long

di = Day(rsid!incDate)
mi = Month(rsid!incDate)
yi = Year(rsid!incDate)
       
dp = Day(Date)
mp = Month(Date)
yp = Year(Date)
       


LabGrade.Caption = rsid!grade


Dim rtCmd As New ADODB.Command
Set rtCmd.ActiveConnection = dbCnn
rtCmd.CommandText = "returnGrade"
rtCmd.CommandType = adCmdStoredProc

rtCmd.Parameters.Append rtCmd.CreateParameter("gd", adVarChar, adParamInput, 50, LabGrade.Caption)

rtCmd.Parameters.Append rtCmd.CreateParameter("bs", adVarChar, adParamOutput, 50)
'rtCmd.Parameters.Append rtCmd.CreateParameter("mn", adVarChar, adParamInput, 50, 10.5)


''''''''''
rtCmd.Parameters.Append rtCmd.CreateParameter("dns", adVarChar, adParamOutput, 50)
rtCmd.Parameters.Append rtCmd.CreateParameter("hr", adVarChar, adParamOutput, 50)
rtCmd.Parameters.Append rtCmd.CreateParameter("md", adVarChar, adParamOutput, 50)
rtCmd.Parameters.Append rtCmd.CreateParameter("hdcg", adVarChar, adParamOutput, 50)
rtCmd.Parameters.Append rtCmd.CreateParameter("rec", adVarChar, adParamOutput, 50)
rtCmd.Parameters.Append rtCmd.CreateParameter("fes", adVarChar, adParamOutput, 50)
rtCmd.Parameters.Append rtCmd.CreateParameter("ot", adVarChar, adParamOutput, 50)

rtCmd.Execute

LabGrade.Caption = rsid!grade






Text8 = rtCmd.Parameters("dns")



Dim prepostid, idid As String

idid = rsid!id

If rs.BOF Then
prepostid = 0
Else
prepostid = rs!id
End If

If Not idid = prepostid Then    'if not even updated for asingle time
                                   'then scale otherwise basic from prepost_t
Text9 = rtCmd.Parameters("bs")
Else

If yp > yi Or (yp = yi And mp > mi) Or (yp = yi And mp = mi And dp >= di) Then
       
       basic = rs!basic + gd
Else: basic = rs!basic
End If

Text9 = basic
End If


dr = Text8
Text10 = rtCmd.Parameters("hr")

hr = Text10
Text11 = rtCmd.Parameters("md")

m = Text11
Text12 = rtCmd.Parameters("hdcg")

Hd = Text12
Text14 = rtCmd.Parameters("rec")

rec = Text14
Text15 = rtCmd.Parameters("fes")

fs = Text15
Text16 = rtCmd.Parameters("ot")

ot = Text16
sum = dr + hr + m + Hd + rec + fs + ot + Text9
LabelTotalSal.Caption = sum




Text21 = ""
Text22 = ""
Text23 = ""
Text24 = ""
Text25 = ""
Text26 = ""
Text27 = ""
Text28 = ""
LabTotalDed.Caption = ""
LabelAmmount.Caption = ""













End Sub

Private Sub Command3_Click()
Form3.Show

End Sub

Private Sub CmdFirst1_Click()

rsid.MoveFirst

If rsid.BOF Then
rsid.MoveLast
End If








If postcall = 1 Then

rs.MoveFirst
If rsid.BOF Then
rsid.MoveFirst
End If

If rs.BOF Then
rs.MoveFirst
End If


showpost

Else
showField
showid
End If
















'rsid.MoveFirst




'showField
'showid

'If postcall = 1 Then
'showpost
'End If



End Sub

Private Sub CmdLast1_Click()


rsid.MoveLast


If rsid.EOF Then
rsid.MoveLast
End If








If postcall = 1 Then

rs.MoveLast
If rsid.BOF Then
rsid.MoveLast
End If

If rs.BOF Then
rs.MoveLast
End If


showpost

Else
showField
showid
End If


















'rsid.MoveLast
'If postcall = 1 Then
'showpost

'Else

'showField
'showid
'End If





End Sub

Private Sub CmdNext1_Click()
rsid.MoveNext


If rsid.EOF Then
rsid.MoveFirst
End If








If postcall = 1 Then

rs.MoveNext
If rsid.EOF Then
rsid.MoveFirst
End If

If rs.EOF Then
rs.MoveFirst
End If

showpost

Else
showField
showid
End If


End Sub

Private Sub CmdPost_Click()



Dim stut As String

stut = MsgBox("Are you sure to Post this Information ?", vbQuestion + vbYesNo)

If stut = vbNo Then
Exit Sub
End If




Dim Cmd As New ADODB.Command
Set Cmd.ActiveConnection = dbCnn
Cmd.CommandText = "postcheck"
Cmd.CommandType = adCmdStoredProc

Cmd.Parameters.Append Cmd.CreateParameter("id", adVarChar, adParamInput, 50, LabelId.Caption)
Cmd.Parameters.Append Cmd.CreateParameter("status", adInteger, adParamOutput, 50)
Cmd.Execute

If Cmd.Parameters("status") = 1 Then
GoTo idExist
End If
















Dim objCmd As New ADODB.Command
Set objCmd.ActiveConnection = dbCnn
objCmd.CommandText = "post_bill"
objCmd.CommandType = adCmdStoredProc

objCmd.Parameters.Append objCmd.CreateParameter("RV", adVarChar, adParamReturnValue, 50)
objCmd.Parameters.Append objCmd.CreateParameter("pid", adVarChar, adParamInput, 50, LabelId.Caption)
objCmd.Parameters.Append objCmd.CreateParameter("pbs", adVarChar, adParamInput, 50, Text9)
objCmd.Parameters.Append objCmd.CreateParameter("pdns", adVarChar, adParamInput, 50, Text8)
objCmd.Parameters.Append objCmd.CreateParameter("phr", adVarChar, adParamInput, 50, Text10)

objCmd.Parameters.Append objCmd.CreateParameter("pmed", adVarChar, adParamInput, 50, Text11)
objCmd.Parameters.Append objCmd.CreateParameter("phdcg", adVarChar, adParamInput, 50, Text12)
objCmd.Parameters.Append objCmd.CreateParameter("prec", adVarChar, adParamInput, 50, Text14)
objCmd.Parameters.Append objCmd.CreateParameter("pfes", adVarChar, adParamInput, 50, Text15)
objCmd.Parameters.Append objCmd.CreateParameter("potalw", adVarChar, adParamInput, 50, Text16)


objCmd.Parameters.Append objCmd.CreateParameter("pgpins", adVarChar, adParamInput, 50, Text21)
objCmd.Parameters.Append objCmd.CreateParameter("pcbs", adVarChar, adParamInput, 50, Text22)
objCmd.Parameters.Append objCmd.CreateParameter("pgpfnd", adVarChar, adParamInput, 50, Text23)


objCmd.Parameters.Append objCmd.CreateParameter("pdedHr", adVarChar, adParamInput, 50, Text24)
objCmd.Parameters.Append objCmd.CreateParameter("pVbl", adVarChar, adParamInput, 50, Text25)
objCmd.Parameters.Append objCmd.CreateParameter("pEbl", adVarChar, adParamInput, 50, Text26)
objCmd.Parameters.Append objCmd.CreateParameter("pTbl", adVarChar, adParamInput, 50, Text27)
objCmd.Parameters.Append objCmd.CreateParameter("potDed", adVarChar, adParamInput, 50, Text28)
objCmd.Parameters.Append objCmd.CreateParameter("potDed", adVarChar, adParamInput, 50, Text7)

objCmd.Execute






If objCmd.Parameters("RV") = 0 Then
MsgBox " This Information is posted successfully.", vbInformation
Exit Sub
End If

idExist:
MsgBox "This id is already posted in this month.", vbCritical





End Sub

Private Sub CmdPrev1_Click()

rsid.MovePrevious


If rsid.BOF Then
rsid.MoveLast
End If








If postcall = 1 Then

rs.MovePrevious
If rsid.BOF Then
rsid.MoveLast
End If

If rs.BOF Then
rs.MoveLast
End If


showpost

Else
showField
showid
End If





End Sub

Private Sub CmdSearch_Click()
Form8.Show

End Sub

Private Sub CmdUpdate_Click()

If Text21.Text = "" Or Text22.Text = "" Or Text23.Text = "" Or _
Text24.Text = "" Or Text25.Text = "" Or Text26.Text = "" Or _
Text27.Text = "" Or Text28.Text = "" Then

MsgBox "You must Enter All the Fields. ", vbCritical
Exit Sub
End If

Dim cnfrmstr As String







cnfrmstr = MsgBox("Do You really want to Update ?", vbYesNo + vbQuestion)
If cnfrmstr = vbNo Then
Exit Sub
End If



LabelId.Caption = rsid!id


Dim objCmd As New ADODB.Command
Set objCmd.ActiveConnection = dbCnn
objCmd.CommandText = "update_pre_bill"
objCmd.CommandType = adCmdStoredProc

objCmd.Parameters.Append objCmd.CreateParameter("RV", adVarChar, adParamReturnValue, 50)
objCmd.Parameters.Append objCmd.CreateParameter("uid", adVarChar, adParamInput, 50, LabelId.Caption)
objCmd.Parameters.Append objCmd.CreateParameter("ubs", adVarChar, adParamInput, 50, Text9)
objCmd.Parameters.Append objCmd.CreateParameter("udns", adVarChar, adParamInput, 50, Text8)
objCmd.Parameters.Append objCmd.CreateParameter("uhr", adVarChar, adParamInput, 50, Text10)

objCmd.Parameters.Append objCmd.CreateParameter("umed", adVarChar, adParamInput, 50, Text11)
objCmd.Parameters.Append objCmd.CreateParameter("uhdcg", adVarChar, adParamInput, 50, Text12)
objCmd.Parameters.Append objCmd.CreateParameter("urec", adVarChar, adParamInput, 50, Text14)
objCmd.Parameters.Append objCmd.CreateParameter("ufes", adVarChar, adParamInput, 50, Text15)
objCmd.Parameters.Append objCmd.CreateParameter("uotalw", adVarChar, adParamInput, 50, Text16)


objCmd.Parameters.Append objCmd.CreateParameter("ugpins", adVarChar, adParamInput, 50, Text21)
objCmd.Parameters.Append objCmd.CreateParameter("ucbs", adVarChar, adParamInput, 50, Text22)
objCmd.Parameters.Append objCmd.CreateParameter("ugpfnd", adVarChar, adParamInput, 50, Text23)


objCmd.Parameters.Append objCmd.CreateParameter("udedHr", adVarChar, adParamInput, 50, Text24)
objCmd.Parameters.Append objCmd.CreateParameter("uVbl", adVarChar, adParamInput, 50, Text25)
objCmd.Parameters.Append objCmd.CreateParameter("uEbl", adVarChar, adParamInput, 50, Text26)
objCmd.Parameters.Append objCmd.CreateParameter("uTbl", adVarChar, adParamInput, 50, Text27)
objCmd.Parameters.Append objCmd.CreateParameter("uotDed", adVarChar, adParamInput, 50, Text28)

objCmd.Execute

If objCmd.Parameters("RV") = 0 Then
MsgBox " This Information is updated successfully.", vbInformation
'rs.Update
End If


End Sub

Private Sub Form_Load()

Set dbCnn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rsid = New ADODB.Recordset
Set rsal = New ADODB.Recordset
Set rstt = New ADODB.Recordset
 
dbCnn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Teacher PayRoll"

rs.Open "prepost_T ", dbCnn, adOpenKeyset, adLockOptimistic, -1
rsid.Open "Id_T ", dbCnn, adOpenKeyset, adLockOptimistic, -1
rsal.Open "Allowance_T ", dbCnn, adOpenKeyset, adLockOptimistic, -1



Labelyear = Year(Date)

Dim mon As Integer
Dim monstr As String

mon = Month(Date)



 




If mon = 1 Then
monstr = "JANUARY"

ElseIf mon = 2 Then
monstr = "FEBRUARY"

ElseIf mon = 3 Then
monstr = "MARCH"

ElseIf mon = 4 Then
monstr = "APRIL"

ElseIf mon = 5 Then
monstr = "MAY"

ElseIf mon = 6 Then
monstr = "JUNE"

ElseIf mon = 7 Then
monstr = "JULY"

ElseIf mon = 8 Then
monstr = "AUGUST"

ElseIf mon = 9 Then
monstr = "SEPTEMBER"

ElseIf mon = 10 Then
monstr = "OCTOBER"

ElseIf mon = 11 Then
monstr = "NOVEMBER"

Else: monstr = "DECEMBER"
End If



Labelmonth = monstr


If postcall = 1 Then
CmdUpdate.Visible = False
CmdPost.Visible = True

Else
CmdUpdate.Visible = True
CmdPost.Visible = False
End If


rsid.MoveFirst

Dim chkid As String


 
 
  'srsh = 1
'id2srsh = "3"

 
 
chkid = rsid!id
 
 
 



If srsh = 1 Then
While Not chkid = id2srsh
rsid.MoveNext
chkid = rsid!id
Wend

End If


If srsh = 2 Then
chkid = rs!id

While Not chkid = id2srsh
rs.MoveNext
chkid = rs!id
Wend

End If

 If postcall = 1 Then
showpost
Else




 showid
 showField
End If

'no:
End Sub
