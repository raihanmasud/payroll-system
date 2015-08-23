VERSION 5.00
Begin VB.Form form5 
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Id No "
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbCnn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub CmdDelete_Click()
Dim strmsg As String

strmsg = MsgBox("Are You sure to Delete this Id ?", vbYesNo + vbQuestion)


Dim objCmd As New ADODB.Command
Set objCmd.ActiveConnection = dbCnn
objCmd.CommandText = "DeleteEntry"
objCmd.CommandType = adCmdStoredProc

objCmd.Parameters.Append objCmd.CreateParameter("deleteId", adVarChar, adParamInput, 50, Text1)
objCmd.Parameters.Append objCmd.CreateParameter("ExistId", adInteger, adParamOutput)



objCmd.Execute
If objCmd.Parameters("ExistId") = 0 Then
MsgBox "Id No. " + Text1.Text + " does not Exist."

Else
MsgBox "Id No. " + Text1.Text + " has been Deleted Successfully."

Text1.Text = ""
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
Set dbCnn = New ADODB.Connection
Set rs = New ADODB.Recordset
dbCnn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Teacher PayRoll"
rs.Open "Id_T ", dbCnn, adOpenKeyset, adLockOptimistic, -1

End Sub

