VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6120
   LinkTopic       =   "Form8"
   ScaleHeight     =   4080
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox ComSrhbill 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton CmdSearchId 
      Caption         =   "Search"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label LabelSearch 
      Caption         =   "Select  an  Id  No.   "
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsch As ADODB.Recordset





Private Sub CmdSearchId_Click()
Unload Form4





id2srsh = ComSrhbill.Text


If postcall = 1 Then
srsh = 2
Form4.LabelPre.Visible = False
Else
srsh = 1


End If




Form4.Show




End Sub

Private Sub Form_Load()

Call Connection
Set rsch = New ADODB.Recordset

If postcall = 1 Then
rsch.Open "prepost_t", dbCnn, adOpenKeyset, adLockOptimistic, -1
Else
rsch.Open " id_T", dbCnn, adOpenKeyset, adLockOptimistic, -1
End If

'rsch.Update

Do Until rsch.EOF = True
ComSrhbill.AddItem rsch!id
rsch.MoveNext
Loop
  
End Sub
