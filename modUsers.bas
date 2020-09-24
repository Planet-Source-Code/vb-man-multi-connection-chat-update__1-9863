Attribute VB_Name = "modUsers"
Option Explicit

Public Sub LoadListBox(lb As ListBox)

Dim rsTemp As Recordset
Dim strSelect As String

lb.Clear
strSelect = "SELECT * FROM Users ORDER BY Username"
Set rsTemp = dBase.OpenRecordset(strSelect)

Do Until rsTemp.EOF
    lb.AddItem rsTemp!UserName
    rsTemp.MoveNext
Loop
rsTemp.Close

End Sub

Function GetPassword(strUsername As String) As String
Dim rsTemp As Recordset
Dim strSelect As String

strSelect = "SELECT * FROM Users WHERE Username=" & Chr(34) & strUsername & Chr(34)
Set rsTemp = dBase.OpenRecordset(strSelect)

Do Until rsTemp.EOF
    GetPassword = rsTemp!Password
    rsTemp.MoveNext
Loop

rsTemp.Close

End Function

Public Sub AddUser(strUser As String, strPassword As String)
Dim rsTemp As Recordset
Dim strSelect As String

strSelect = "SELECT * FROM Users"
Set rsTemp = dBase.OpenRecordset(strSelect)

With rsTemp
    .AddNew
    rsTemp!UserName = strUser
    rsTemp!Password = strPassword
    .Update
    .Bookmark = .LastModified
End With

rsTemp.Close
End Sub

Public Sub UpdateUser(strUser As String, strPassword As String)
Dim rsTemp As Recordset
Dim strSelect As String

strSelect = "SELECT * FROM Users WHERE Username=" & Chr(34) & strUser & Chr(34)
Set rsTemp = dBase.OpenRecordset(strSelect)

With rsTemp
    .Edit
    rsTemp!Password = strPassword
    .Update
    .Bookmark = .LastModified
End With

rsTemp.Close
End Sub

Public Sub DeleteUser(strUser As String)
Dim rsTemp As Recordset
Dim strSelect As String

strSelect = "SELECT * FROM Users WHERE Username=" & Chr(34) & strUser & Chr(34)
Set rsTemp = dBase.OpenRecordset(strSelect)

With rsTemp
    .Delete
End With

rsTemp.Close
End Sub
