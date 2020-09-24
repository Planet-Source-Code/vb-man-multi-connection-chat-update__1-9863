VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmServer 
   Caption         =   "Chat Server"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtStream 
      Height          =   2190
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   3863
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmServer.frx":0000
   End
   Begin VB.ListBox lstUsers 
      Height          =   2205
      Left            =   6150
      TabIndex        =   2
      Top             =   75
      Width           =   1365
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   2955
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   1500
      Top             =   1650
   End
   Begin MSWinsockLib.Winsock wsServer 
      Index           =   0
      Left            =   840
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtBuffer 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   2475
      Width           =   4635
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOnline 
      Caption         =   "&Online"
      Begin VB.Menu mnuListen 
         Caption         =   "&Listen"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuCommands 
      Caption         =   "&Commands"
      Begin VB.Menu mnuKick 
         Caption         =   "&Kick User"
      End
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "&Users"
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const minWidth As Integer = 100
Private Const minHeight As Integer = 100
Private bAuthorizationFailed(50) As Boolean

Private Sub Form_Load()
Me.StatusBar1.Panels.Item(1).AutoSize = sbrSpring
Me.StatusBar1.Panels.Add 2
Me.StatusBar1.Panels.Item(2).Alignment = sbrCenter
Me.StatusBar1.Panels.Item(2) = "Port: " & iPort
mnuListen_Click
End Sub

Private Sub Form_Resize()

If Me.Height > 1500 Then
    
    'txtbuffer
    Me.txtBuffer.Left = 25
    Me.txtBuffer.Top = Me.Height - 1000 - Me.StatusBar1.Height
    Me.txtBuffer.Width = Me.Width - 100
    Me.txtBuffer.Height = 315
    
    'lstusers
    Me.lstUsers.Left = Me.Width - 1500
    Me.lstUsers.Top = 50
    Me.lstUsers.Width = 1365
    Me.lstUsers.Height = Me.txtBuffer.Top - 200
    
    'resize the txtStream
    Me.txtStream.Left = 25
    Me.txtStream.Top = 50
    Me.txtStream.Height = Me.lstUsers.Height
    Me.txtStream.Width = Me.lstUsers.Left - 200

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
dBase.Close
End Sub

Private Sub mnuDisconnect_Click()
Me.wsServer(0).Close

Me.StatusBar1.Panels.Item(1) = "Not Listening"

Me.mnuListen.Enabled = True
Me.mnuDisconnect.Enabled = False
Me.mnuOptions.Enabled = True
End Sub

Private Sub mnuKick_Click()
KickUser
End Sub

Private Sub mnuListen_Click()
On Error GoTo EH
Me.mnuOptions.Enabled = False

Me.wsServer(0).LocalPort = iPort
Me.wsServer(0).Listen

Me.StatusBar1.Panels.Item(1) = "Listening on port " & iPort

Me.mnuListen.Enabled = False
Me.mnuDisconnect.Enabled = True
Exit Sub

EH:
Dim i As Integer
i = MsgBox("Error # " & Str(Err.Number) & " was generated by " _
         & Err.Source & Chr(13) & Err.Description & Chr(13) _
         & "Please select a new port.", 64, "Error")
        
Me.mnuOptions.Enabled = True

End Sub

Private Sub mnuOptions_Click()
frmServerOptions.Show vbModal
End Sub

Private Sub mnuUsers_Click()
frmUsers.Show vbModal
End Sub

Private Sub Timer1_Timer()
MsgBox ("It worked")
Unload Me
End Sub



Private Sub wsServer_Close(Index As Integer)
Dim iListIndex As Integer

Me.wsServer(Index).Close
Unload wsServer(Index)

InsertText "Disconnected from " & uUser(Index).RequestID

uUser(Index).Connection = "Disconnected"

'only send disconnection notice if the user has already successfully logged in
'don't want the notice going out if the user fails to login successfully
If bAuthorizationFailed(Index) = False Then
    
    'send disconnection notice to other users
    SendNoticeOfDisconnectedUser uUser(Index).Name

End If

'remove name from list
iListIndex = FindListItemIndex(uUser(Index).Name)
If iListIndex <> -1 Then
    Me.lstUsers.RemoveItem (iListIndex)
End If

uUser(Index).RequestID = 0
uUser(Index).Name = ""

iClients = iClients - 1
UpdateTitleBar
End Sub

Private Sub wsServer_Connect(Index As Integer)
Me.txtBuffer = "Connected"
End Sub

Private Sub wsserver_ConnectionRequest(Index As Integer, ByVal RequestID As Long)
Dim iNextSocket As Integer

iNextSocket = modServer.FindOpenSocket(Me)

'for multiple connections
Load wsServer(iNextSocket)
wsServer(iNextSocket).Accept RequestID
InsertText "Connected to " & RequestID

uUser(iNextSocket).Connection = "Connected"
uUser(iNextSocket).RequestID = RequestID

End Sub

Private Sub wsServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Dim strFormattedData As String
Dim strUsername As String
Dim sCommand As String
Dim strFrom As String
Dim iUserNumber As Integer
Dim bAuth As Boolean

Me.wsServer(Index).GetData strData

If Left(strData, 2) = "__" Or Left(strData, 1) = "/" Then
    
    'for Name command
    If Left(strData, 6) = "__NAME" Then
    
        strUsername = Right(strData, Len(strData) - 6)
        
        'check for duplicate name
        If DupeName(Me, UCase(strUsername)) Then
            'if duplicate name then send a command to the client forcing
            'them to change the name
            'disconnect from the client with a duplicate name
            SendDataNow "COMMAND_NEWNAME" & strUsername, Index
            DoEvents
            wsServer_Close (Index)
        Else
            'if username is unique then add name to the array
            uUser(Index).Name = strUsername
            'add to total clients
            iClients = iClients + 1
            'update the title bar with new number of clients
            UpdateTitleBar
            'add username to list
            Me.lstUsers.AddItem strUsername
            
            'if security is enabled, then ask for password
            If modIni.GetINIValue("Server", "Security", App.Path & "\chat.ini") = "True" Then
                SendDataNow "COMMAND_PASSWORD", Index
            Else
                'tell all clients that someone else has joined the chat
                PassDataToAllClients "COMMAND_SERVER:" & strUsername & " has joined the chat"
                ShowUsers Index
            End If
        End If

    End If
    
    'for Task List
    If Left(strData, 10) = "__TASKLIST" Then
        SendTaskListCommand (strData)
    End If
    
    'receiving the task list
    If Left(strData, 6) = "__INFO" Then
        ParseTaskList (strData)
    End If
    
    'request for a list of all users
    If UCase(Mid(strData, 2, 3)) = "WHO" Then
        ShowUsers Index
    End If
    
    'send disconnect notice
    If UCase(Mid(strData, 2, 10)) = "DISCONNECT" Then
        'do something
    End If
    
    'private message
    If UCase(Left(strData, 9) = "__PRIVATE") Then
        'get the username, it will be all in caps
        strUsername = Mid(strData, 10, InStr(10, strData, "&&") - 10)
        'find the usernames connection number
        iUserNumber = FindUserNumber(strUsername)
        'get the properly formatted name
        strUsername = uUser(iUserNumber).Name
        strFrom = uUser(Index).Name
        strFormattedData = strFrom & "(Private Msg): " & Right(strData, Len(strData) - Len(strUsername) - 11)
        SendDataNow strFormattedData, iUserNumber
        InsertText "Private Message: ", False, 5485348, True
        InsertText strFrom, False, 8388863, True
        InsertText " to ", False, 5485348, False
        InsertText strUsername, True, 8388863, True
        InsertText "--->" & Right(strData, Len(strData) - Len(strUsername) - 11), True, 16711808, False
    End If
    
    'accept password
    If UCase(Mid(strData, 3, 8)) = "PASSWORD" Then
        bAuthorizationFailed(Index) = False
        bAuth = modServer.Authorize(uUser(Index).Name, Mid(strData, 11, Len(strData) - 10))
        If Not bAuth Then
            'something
            SendDataNow "COMMAND_SERVER:Password incorrect", Index
            bAuthorizationFailed(Index) = True
            wsServer_Close Index
            DoEvents
        Else
            PassDataToAllClients "COMMAND_SERVER:" & uUser(Index).Name & " has joined the chat"
            ShowUsers Index
        End If
    End If
Else
    'set the name of the user who sent the request
    strUsername = uUser(Index).Name
    strFormattedData = strUsername & ": " & Left(strData, Len(strData) - 2) 'get rid of vbcrlf
    PassDataToAllClients strFormattedData
    InsertText strFormattedData, True, , False
End If

End Sub

Private Sub PassDataToAllClients(strData As String)
'pass the text received onto all clients.
Dim i As Integer
Dim iConnections As Integer
Dim j As Long

iConnections = wsServer.ubound

'start at 1 because 0 is never connected, it just creates another
'winsock control and connects to the new one
For i = 1 To iConnections
    If uUser(i).Connection = "Connected" Then
        SendDataNow strData, i
        DoEvents
    End If
Next i

End Sub

Private Sub SendDataNow(strData As String, iConnection As Integer)
    If uUser(iConnection).Connection = "Connected" Then
        Me.wsServer(iConnection).SendData strData
        DoEvents
    End If
End Sub

Private Sub UpdateTitleBar()
If iClients > 1 Then
    Me.Caption = iClients & " users connected"
Else
    If iClients = 1 Then
        Me.Caption = iClients & " user connected"
    Else
        Me.Caption = "No users connected"
    End If
End If
End Sub

Private Sub SendTaskListCommand(strData As String)

'send the command to only the client with the matching name in the request

Dim strUsername As String
Dim i As Integer
Dim iUserNumber As Integer

strUsername = UCase(Right(strData, Len(strData) - 10))
strUsername = Left(strUsername, InStr(1, strUsername, vbCrLf) - 1)
strOriginalUser = UCase(Right(strData, Len(strData) - InStr(1, strData, "USER:") - 4))

iUserNumber = FindUserNumber(strUsername)

If iUserNumber <> 0 Then    'if user is found
    Me.wsServer(iUserNumber).SendData "COMMAND_TASKLIST" & strUsername
End If

End Sub

Private Sub ParseTaskList(strData As String)
'take apart string of tasks and display in the server window
Dim i As Integer
Dim strTmp As String
Dim strTask As String
Dim iUserNumber As Integer
Dim strText As String

strTmp = Right(strData, Len(strData) - 6)

'find the winsock connection of the user who sent the request
'used below to send the information back to only the requester
iUserNumber = FindUserNumber(strOriginalUser)

i = 1
strText = "COMMAND_TASKS***"
Do Until i = 0
    
    i = InStr(1, strTmp, "***")
    
    If i <> 0 Then
        strTask = Left(strTmp, i - 1)
        Me.txtStream = Me.txtStream & vbCrLf & strTask
        strText = strText & strTask & "***"
        strTmp = Right(strTmp, Len(strTmp) - i - 2)
    End If

Loop

If iUserNumber <> 0 Then
    SendDataNow strText, iUserNumber
End If

End Sub

Private Sub ShowUsers(iIndex As Integer)
'send a list of all connected users to strUsername
Dim strList As String
Dim i As Integer
Dim iUserNumber As Integer

strList = "COMMAND_SERVER:"

'build the list of all current users
For i = 1 To Me.wsServer.ubound
    If uUser(i).Connection = "Connected" Then
        strList = strList & uUser(i).Name & vbCrLf
    End If
Next i

'remove the last vbcrlf
strList = Left(strList, Len(strList) - 2)
strList = strList & "FILLLISTBOX"
SendDataNow strList, iIndex

End Sub

Function FindUserNumber(strUsername) As Integer
'finds the usernumber given a username

Dim i As Integer

For i = 1 To wsServer.ubound
    If uUser(i).Connection = "Connected" Then
        If UCase(uUser(i).Name) = UCase(strUsername) Then
            FindUserNumber = i
            i = wsServer.ubound + 1
        End If
    End If
Next i
End Function

Private Sub SendNoticeOfDisconnectedUser(strUsername As String)

PassDataToAllClients "COMMAND_SERVER:" & strUsername & " has left the chat"

End Sub

Private Function FindListItemIndex(strItem As String) As Integer
'find an item's index number in the list box
Dim i As Integer
Dim bFound As Boolean

For i = 0 To Me.lstUsers.ListCount - 1
    If Me.lstUsers.List(i) = strItem Then
        FindListItemIndex = i
        bFound = True
        i = Me.lstUsers.ListCount
    End If
Next i

If Not bFound Then
    FindListItemIndex = -1
End If

End Function

Private Sub InsertText(strText As String, Optional bCarriageReturn As Boolean = True, Optional lColour As Long, Optional bBold As Boolean)
'add text to object with colour

'default to black if 0
If lColour = 0 Then
    lColour = 16512
End If

Me.txtStream.SelStart = Len(Me.txtStream)
Me.txtStream.SelColor = lColour

If bBold Then
    Me.txtStream.SelBold = True
Else
    Me.txtStream.SelBold = False
End If

If bCarriageReturn Then
    strText = strText & vbCrLf
End If

Me.txtStream.SelText = strText

End Sub

Private Sub KickUser()
'kicks the selected user from the Listbox
Dim strUsertoKick As String
Dim iUserNumber As Integer

strUsertoKick = Me.lstUsers.List(Me.lstUsers.ListIndex)
iUserNumber = FindUserNumber(strUsertoKick)

If iUserNumber <> 0 Then
    Me.wsServer(iUserNumber).SendData "COMMAND_KICK"
End If
DoEvents

End Sub
