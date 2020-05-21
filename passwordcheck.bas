Attribute VB_Name = "Module1"
Function connect(usrname As String)
Dim conn As ADODB.Connection
Dim mrc As ADODB.Recordset
Set conn = New ADODB.Connection
Set mrc = New ADODB.Recordset
Dim constring As String
constring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=43.mdb;Persist Security Info=False"
conn.Open constring
conn.CursorLocation = adUseClient
mrc.Open "select * from user where username='" & usrname & "'", conn, adOpenKeyset, adLockOptimistic
End Function
