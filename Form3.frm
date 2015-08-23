VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Pay Bill"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7215
   LinkTopic       =   "Form3"
   ScaleHeight     =   5445
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Comlist 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Posted ID"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Updated ID"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.OptionButton Option0 
      Caption         =   "Listed ID"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Labstate 
      Caption         =   "Labstate"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Labinf 
      Caption         =   "INFORMATION OF "
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Id  No"
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsu As ADODB.Recordset
Dim rsid As ADODB.Recordset
Dim rspst As ADODB.Recordset


Option Explicit



Private Sub Command1_Click()
Form4.Show
End Sub

Private Sub Command2_Click()
Form1.Show
End Sub
Private Sub AddItem()
Comlist.Clear

'rsid.MoveFirst

If state = 0 Then

Do Until rsid.EOF = True
Comlist.AddItem rsid!id
rsid.MoveNext
Loop


ElseIf state = 1 Then

Do Until rsu.EOF = True
Comlist.AddItem rsu!id
rsu.MoveNext
Loop


Else

Do Until rspst.EOF = True
Comlist.AddItem rspst!id
rspst.MoveNext
Loop
End If

state = 7

End Sub
Private Sub Form_Load()

Call Connection

 Set rsu = New ADODB.Recordset
 Set rsid = New ADODB.Recordset
 Set rspst = New ADODB.Recordset



rsu.Open "prepost_T ", dbCnn, adOpenKeyset, adLockOptimistic, -1
rsid.Open "Id_T ", dbCnn, adOpenKeyset, adLockOptimistic, -1
rspst.Open "post_T ", dbCnn, adOpenKeyset, adLockOptimistic, -1

state = 7

End Sub

Private Sub Option0_Click()
'Unload Form3
state = 0
AddItem

Form_Load

End Sub

Private Sub Option1_Click()
'Unload Form3

state = 1 'posted
AddItem
Form_Load

End Sub

Private Sub Option2_Click()
'Unload Form3

state = 2
AddItem
Form_Load
End Sub
