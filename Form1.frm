VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000017&
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8640
   FillColor       =   &H00C00000&
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8190
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackColor       =   &H80000015&
      Caption         =   "TEACHER  PAY  ROLL  SYSTEM"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Menu Entry 
      Caption         =   "&Entry "
      Begin VB.Menu NewEntry 
         Caption         =   "&New Entry"
      End
      Begin VB.Menu EditEntry 
         Caption         =   "&Edit Entry"
      End
      Begin VB.Menu Delete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu View 
      Caption         =   "&View"
      Begin VB.Menu ListedId 
         Caption         =   "&Listed Id"
      End
      Begin VB.Menu Updated_Id 
         Caption         =   "&Updated Id"
      End
      Begin VB.Menu PostedId 
         Caption         =   "&Posted Id"
      End
   End
   Begin VB.Menu PayBill 
      Caption         =   "&Pay Bill"
      Begin VB.Menu PrePost 
         Caption         =   "&Pre Posting"
      End
      Begin VB.Menu Posting 
         Caption         =   "&Posting"
      End
   End
   Begin VB.Menu Search 
      Caption         =   "&Search"
      Begin VB.Menu Through 
         Caption         =   "&Through"
      End
      Begin VB.Menu Individual 
         Caption         =   "&Individual"
      End
   End
   Begin VB.Menu Report 
      Caption         =   "&Report"
      Begin VB.Menu PaySlip 
         Caption         =   "&Pay Slip"
      End
   End
   Begin VB.Menu Author 
      Caption         =   "&Author"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim dbCnn As ADODB.Connection


Dim rs As ADODB.Recordset



Private Sub Author_Click()
Form7.Show

End Sub

Private Sub Delete_Click()
form5.Show
End Sub

Private Sub EditEntry_Click()
entrytype = 1
Form2.Show

Form2.CmdAdd.Visible = False
Form2.CmdAdd.Visible = False
Form2.CmdAdd.Visible = False
Form2.ComDeg.Visible = False
Form2.ComDep.Visible = False

Form2.Text5.Visible = True
Form2.Text6.Visible = True



Form2.cmdEdit.Visible = True




End Sub

Private Sub Form_Load()


Set dbCnn = New ADODB.Connection
Set rs = New ADODB.Recordset


dbCnn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Teacher PayRoll"
rs.Open "prepost_t ", dbCnn, adOpenKeyset, adLockOptimistic, -1
postcall = 0

End Sub

Private Sub Individual_Click()
form6.Show

End Sub

Private Sub ListedId_Click()
Form3.Show
state = 0
End Sub

Private Sub NewEntry_Click()
Form2.CmdFirst.Visible = False
Form2.CmdLast.Visible = False
Form2.CmdPrev.Visible = False
Form2.CmdNext.Visible = False

Form2.Text4.Visible = False
Form2.Text8.Visible = False


Form2.Show

End Sub


Private Sub Seach_Click()

End Sub

Private Sub PaySlip_Click()
Form9.Show

End Sub

Private Sub PostedId_Click()
state = 1
Form3.Show
End Sub

Private Sub Posting_Click()
postcall = 1

Form4.CmdUpdate.Visible = False
Form4.LabelPre.Visible = False
Form4.Show



End Sub

Private Sub PrePost_Click()

Unload Form4
postcall = 0
Form4.CmdPost.Visible = False


Form4.Show

End Sub

Private Sub Through_Click()
menusrsh = 1


Form2.ComDeg.Visible = False
Form2.ComDep.Visible = False
Form2.cmdEdit.Visible = False
Form2.CmdAdd.Visible = False

Form2.Text5.Visible = True
Form2.Text6.Visible = True
Form2.Clear.Visible = False

Form2.Show

End Sub

Private Sub Updated_Id_Click()
state = 2
Form3.Show

End Sub
