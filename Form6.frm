VERSION 5.00
Begin VB.Form form6 
   Caption         =   "Form6"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5730
   LinkTopic       =   "Se"
   ScaleHeight     =   3585
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSearch 
      Caption         =   "&Search"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Id No. to Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbCnn As ADODB.Connection



Dim rs As ADODB.Recordset


Private Sub Command1_Click()

End Sub

Private Sub CmdSearch_Click()


Dim objCmd As New ADODB.Command
Set objCmd.ActiveConnection = dbCnn
objCmd.CommandText = "SearchEntry"
objCmd.CommandType = adCmdStoredProc

objCmd.Parameters.Append objCmd.CreateParameter("RV", adInteger, adParamReturnValue)

objCmd.Parameters.Append objCmd.CreateParameter("Id", adVarChar, adParamInput, 50, Text1)
objCmd.Parameters.Append objCmd.CreateParameter("Name", adVarChar, adParamOutput, 50)

objCmd.Parameters.Append objCmd.CreateParameter("Designation", adVarChar, adParamOutput, 50)
objCmd.Parameters.Append objCmd.CreateParameter("Department", adVarChar, adParamOutput, 50)



objCmd.Parameters.Append objCmd.CreateParameter("Scale", adVarChar, adParamOutput, 50)
objCmd.Parameters.Append objCmd.CreateParameter("Grade", adVarChar, adParamOutput, 50)




objCmd.Parameters.Append objCmd.CreateParameter("joinDate", adVarChar, adParamOutput, 50)
objCmd.Parameters.Append objCmd.CreateParameter("IncDate", adVarChar, adParamOutput, 50)
objCmd.Parameters.Append objCmd.CreateParameter("type", adVarChar, adParamOutput, 50)


objCmd.Parameters.Append objCmd.CreateParameter("ExstId", adInteger, adParamOutput)
objCmd.Execute

If objCmd.Parameters("ExstId") = 0 Then
MsgBox "There is no Id No. as  " + Text1.Text + ". Entry the Id First."
Exit Sub

Unload Form2





Else:
Form2.ComDeg.Visible = False
Form2.ComDep.Visible = False
Form2.cmdEdit.Visible = False
Form2.CmdAdd.Visible = False
Form2.CmdFirst.Visible = False
Form2.CmdLast.Visible = False
Form2.CmdNext.Visible = False
Form2.CmdPrev.Visible = False
Form2.Clear.Visible = False



Form2.Text5.Visible = True
Form2.Text6.Visible = True
Form2.Text4.Visible = True
Form2.Text8.Visible = True

Form2.Text1.Text = form6.Text1.Text
Form2.Text2.Text = objCmd.Parameters("Name")
Form2.Text6.Text = objCmd.Parameters("Department")
Form2.Text5.Text = objCmd.Parameters("Designation")
Form2.Text3.Text = objCmd.Parameters("Scale")

Form2.Text4.Text = objCmd.Parameters("Grade")

Form2.Text7.Text = objCmd.Parameters("joinDate")
Form2.dttxt.Text = objCmd.Parameters("IncDate")
Form2.Text8.Text = objCmd.Parameters("Type")

'Form2.text4.Text = objCmd.Parameters("Pay Grade")
'




Form2.Show

End If








End Sub

Private Sub Form_Load()
Text1.Text = ""

Set dbCnn = New ADODB.Connection
Set rs = New ADODB.Recordset
dbCnn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Teacher PayRoll"
rs.Open "Id_T ", dbCnn, adOpenKeyset, adLockOptimistic, -1






End Sub
