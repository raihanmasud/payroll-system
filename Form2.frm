VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7635
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7125
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7215
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   3240
         TabIndex        =   30
         Text            =   "Text8"
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3240
         TabIndex        =   29
         Text            =   "Text4"
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3240
         TabIndex        =   28
         Text            =   "Text3"
         Top             =   2760
         Width           =   2655
      End
      Begin VB.ComboBox ComPaytype 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox dttxt 
         Height          =   285
         Left            =   3240
         TabIndex        =   25
         Top             =   4200
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3240
         TabIndex        =   24
         Text            =   "Text7"
         Top             =   3720
         Width           =   2655
      End
      Begin VB.ComboBox ComGrade 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3240
         Width           =   2655
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton CmdPrev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   840
         TabIndex        =   19
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton CmdLast 
         Caption         =   "&Find Last"
         Height          =   375
         Left            =   2520
         TabIndex        =   18
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton CmdFirst 
         Caption         =   "&Find First"
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3240
         TabIndex        =   16
         Text            =   "Text6"
         Top             =   2280
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3240
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   1800
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton Clear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   5520
         Width           =   1575
      End
      Begin VB.ComboBox ComDep 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3240
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3240
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.ComboBox ComDeg 
         Height          =   315
         ItemData        =   "Form2.frx":000C
         Left            =   3240
         List            =   "Form2.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Pay Type"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Date of Joining"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Pay  Grade"
         Height          =   375
         Left            =   840
         TabIndex        =   21
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Teacher Database"
         Height          =   495
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Id No"
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   495
         Left            =   840
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Designation"
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Department"
         Height          =   495
         Left            =   840
         TabIndex        =   6
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Pay Scale"
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Date of Increment"
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   4200
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbCnn As ADODB.Connection
Dim strError As String


Dim rsid As ADODB.Recordset
Public objerror As ADODB.Error

Private Sub Clear_Click()
Text1.Text = ""
Text2.Text = ""

dttxt.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub



Private Sub CmdAdd_Click()

On Error GoTo erroroccered

Dim objCmd As New ADODB.Command
Set objCmd.ActiveConnection = dbCnn
objCmd.CommandText = "AdNewEntry"
objCmd.CommandType = adCmdStoredProc

objCmd.Parameters.Append objCmd.CreateParameter("RV", adInteger, adParamReturnValue)

objCmd.Parameters.Append objCmd.CreateParameter("Id", adVarChar, adParamInput, 50, Text1)
objCmd.Parameters.Append objCmd.CreateParameter("Name", adVarChar, adParamInput, 50, Text2)

objCmd.Parameters.Append objCmd.CreateParameter("Designation", adVarChar, adParamInput, 50, ComDeg.Text)
objCmd.Parameters.Append objCmd.CreateParameter("Department", adVarChar, adParamInput, 50, ComDep.Text)


objCmd.Parameters.Append objCmd.CreateParameter("Basic", adVarChar, adParamInput, 50, Text3)
objCmd.Parameters.Append objCmd.CreateParameter("Grade", adVarChar, adParamInput, 50, ComGrade.Text)


objCmd.Parameters.Append objCmd.CreateParameter("JoinDate", adVarChar, adParamInput, 50, Text7)
objCmd.Parameters.Append objCmd.CreateParameter("IncDate", adVarChar, adParamInput, 50, dttxt)
objCmd.Parameters.Append objCmd.CreateParameter("paytype", adVarChar, adParamInput, 50, ComPaytype.Text)






If Text1 = "" Or Text2 = "" Or Text3 = "" _
Or dttxt = "" Or ComDeg.Text = "" Or _
ComDep.Text = "" Then

  MsgBox "You must enter all Fields."
Exit Sub
End If

 
 objCmd.Execute

If objCmd.Parameters("RV") = 0 Then

MsgBox "This Entry is Added to DataBase Successfully"
Set objCmd = Nothing

End If
Exit Sub
erroroccered:
MsgBox "This Id is already  assigned for someone."

End Sub


Private Sub CmdFirst_Click()
rsid.MoveFirst
ShowEntryField


End Sub

Private Sub CmdLast_Click()
rsid.MoveLast
ShowEntryField

End Sub

Private Sub cmdEdit_Click()


If Text1 = "" Or Text2 = "" Or Text3 = "" _
Or dttxt = "" Or Text4 = "" Or Text6 = "" Or Text4 = "" Or _
Text5 = "" Then

  MsgBox "You must enter all Fields."
Exit Sub
End If

Dim objCmd As New ADODB.Command
Set objCmd.ActiveConnection = dbCnn
objCmd.CommandText = "EditEntry"
objCmd.CommandType = adCmdStoredProc

objCmd.Parameters.Append objCmd.CreateParameter("RV", adInteger, adParamReturnValue)

objCmd.Parameters.Append objCmd.CreateParameter("Id", adVarChar, adParamInput, 50, Text1)
objCmd.Parameters.Append objCmd.CreateParameter("Name", adVarChar, adParamInput, 50, Text2)

objCmd.Parameters.Append objCmd.CreateParameter("Designation", adVarChar, adParamInput, 50, ComDeg.Text)
objCmd.Parameters.Append objCmd.CreateParameter("Department", adVarChar, adParamInput, 50, ComDep.Text)


objCmd.Parameters.Append objCmd.CreateParameter("Basic", adVarChar, adParamInput, 50, Text3.Text)

objCmd.Parameters.Append objCmd.CreateParameter("IncDate", adVarChar, adParamInput, 50, dttxt.Text)


objCmd.Parameters.Append objCmd.CreateParameter("ExstId", adVarChar, adParamOutput, 50)
objCmd.Execute

If objCmd.Parameters("ExstId") = 0 Then
MsgBox "There is no Id No as " + Text1.Text + ". Entry the Id First."
Exit Sub


ElseIf objCmd.Parameters("RV") = 0 Then
MsgBox "The Information for " + Text1.Text + " is Updated Successfully."
End If
End Sub

Private Sub CmdMain_Click()
Form1.Show

End Sub

Private Sub CmdNext_Click()
rsid.MoveNext

If rsid.EOF() Then
MsgBox " You are at the End of File."


Else: ShowEntryField
End If

End Sub

Private Sub Command1_Click()
rsid.MoveFirst
ShowEntryField

End Sub

Private Sub Command2_Click()
rsid.MoveLast
End Sub

Private Sub Command5_Click()
Form1.Show
End Sub
Private Sub ShowEntryField()
Text1 = rsid!id
Text2 = rsid!Name
Text3 = rsid!Scale

dttxt = rsid!incDate
Text5 = rsid!Designation
Text6 = rsid!Department
End Sub

Private Sub Command6_Click()
rsid.MoveNext

If rsid.EOF() Then
MsgBox " You are at the End of File."
Else: showFields
End If

End Sub


Private Sub CmdPrev_Click()
rsid.MovePrevious
If rsid.BOF() Then
MsgBox " You are at the Top of File."
Else: ShowEntryField
End If
End Sub






Private Sub ComDeg_Click()

If ComDeg.Text = "Lecturer" Then
Text3.Text = 4800
ElseIf ComDeg.Text = "Assistant Professor" Then
Text3.Text = 6400
ElseIf ComDeg.Text = "Assosiate Professor" Then
Text3.Text = 8300

ElseIf ComDeg.Text = "Professor" Then
Text3.Text = 10200
End If

End Sub

Private Sub Form_Load()
Set dbCnn = New ADODB.Connection
Set rsid = New ADODB.Recordset
dbCnn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Teacher PayRoll"
rsid.Open "Id_T ", dbCnn, adOpenKeyset, adLockOptimistic, -1

ComDeg.AddItem "Lecturer"
ComDeg.AddItem "Assistant Professor"
ComDeg.AddItem "Assosiate Professor"
ComDeg.AddItem "Professor"

ComDep.AddItem "Computer Science && Engineering "
ComDep.AddItem "Electrical && Electronics Engineering"
ComDep.AddItem "Electrical && Communication Engineering"
ComDep.AddItem "Mechanical Engineering"
ComDep.AddItem "Civil Engineering"








ComGrade.AddItem "Grade-1"
ComGrade.AddItem "Grade-2"
ComGrade.AddItem "Grade-3"
ComGrade.AddItem "Grade-4"







Dim dt As Date
dt = Date + 365

ComPaytype.AddItem "Regular"
ComPaytype.AddItem "Leave with Pay"
ComPaytype.AddItem "Leave wiyhout Pay"

dttxt.Text = Date + 365
Text7.Text = Date

Dim strnull As String


If rsid.BOF Then
strnull = "0"
Else
rsid.MoveLast
strnull = rsid!id

End If

Text2.Text = ""
Text3.Text = ""




Text1.Text = "00" + strnull + 1







If entrytype = 1 Or menusrsh = 1 Then
rsid.MoveFirst


Text5.Text = rsid!Designation
Text6.Text = rsid!Department

Text4.Text = rsid!grade
 Text8.Text = rsid!payType
 Text2 = rsid!Name
Text1 = rsid!id
Text3 = rsid!Scale
Text7 = rsid!joinDate
dttxt = rsid!incDate
End If















cmdEdit.Visible = False
entrytype = 7
End Sub

