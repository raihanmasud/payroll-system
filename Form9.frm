VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Comrep 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton CmdReport 
      Caption         =   "Create Pay Slip"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsch As ADODB.Recordset

Private Sub CmdReport_Click()
Dim de As New DataEnvironment1
de.Command1 Comrep.Text

DataReport1.Show





End Sub





Private Sub Form_Load()



Call Connection
Set rsch = New ADODB.Recordset

rsch.Open "post_t", dbCnn, adOpenKeyset, adLockOptimistic, -1





Do Until rsch.EOF = True
Comrep.AddItem rsch!id
rsch.MoveNext
Loop

End Sub
