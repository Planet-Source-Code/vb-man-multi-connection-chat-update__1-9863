Attribute VB_Name = "modServer"
Option Explicit
Public Type User
    Connection As String
    RequestID As Long
    Name As String
End Type

Public iPort As Integer
Public uUser(50) As User
Public iClients As Integer
Public strOriginalUser As String
Public dBase As Database

Private Sub Main()

Dim i As Integer
Dim strTemp As String

strTemp = modIni.GetINIValue("Server", "Port", App.Path & "\chat.ini")
If strTemp <> "" Then
    iPort = Int(strTemp)
Else
    iPort = 0
End If

If iPort = 0 Then
    iPort = 80
    i = modIni.SetINIValue("Server", "Port", "80", App.Path & "\chat.ini")
End If

Set dBase = OpenDatabase(App.Path & "\auth.mdb")

frmServer.Show
frmServer.Caption = "Chat Server"

End Sub


Function FindOpenSocket(frmMe As Form)
'this function finds and returns the first open socket
'available for the incoming request
'this fixes the problem of sockets staying open after a user has
'disconnected.  With this code in place, the next user will pick
'up the previous users spot, rather than get a brand new one

Dim i As Integer

For i = 1 To frmMe.wsServer.ubound
    If uUser(i).Connection = "" Or uUser(i).Connection = "Disconnected" Then
        FindOpenSocket = i
        i = frmMe.wsServer.ubound + 1
    End If
Next i

If FindOpenSocket = 0 Then
    FindOpenSocket = i
End If
End Function

Function DupeName(frmMe As Form, strName As String) As Boolean
'look for a duplicate name
Dim i As Integer

For i = 1 To frmMe.wsServer.ubound
    If uUser(i).Connection = "Connected" Then
        If UCase(uUser(i).Name) = strName Then
            DupeName = True
            i = frmMe.wsServer.ubound + 1
        End If
    End If
Next i

End Function

Function Authorize(strUser As String, strPassword As String) As Boolean

Dim RsTemp As Recordset
Dim strSelect As String

strSelect = "SELECT * FROM Users"
Set RsTemp = dBase.OpenRecordset(strSelect)

Do Until RsTemp.EOF
    If UCase(strUser) = UCase(RsTemp!UserName) Then
        If strPassword = RsTemp!Password Then
            Authorize = True
            RsTemp.MoveLast
        End If
    End If
    RsTemp.MoveNext
Loop
RsTemp.Close
End Function
