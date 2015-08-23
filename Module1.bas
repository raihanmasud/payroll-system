Attribute VB_Name = "Module1"
Public rs As ADODB.Recordset
 Public postcall As Integer
 Public state As Integer
 Public entrytype As Integer
 Public id2srsh As String
 Public edit As Integer
 
 
 Public srsh As Integer
 Public menusrsh As Integer
 Public postsrsh As Integer
 
Public dbCnn As ADODB.Connection


Public Sub Connection()
On Error Resume Next
Set dbCnn = New ADODB.Connection
dbCnn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Teacher PayRoll"
End Sub


