VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmClient 
   Caption         =   "Chat Client"
   ClientHeight    =   4695
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstUsers 
      Height          =   2790
      Left            =   7350
      TabIndex        =   3
      Top             =   0
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Left            =   2250
      Top             =   1050
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1350
      Top             =   1050
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2670
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   60
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   4710
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmClient.frx":0000
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsClient 
      Left            =   840
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtBuffer 
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3000
      Width           =   4515
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
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
      End
   End
   Begin VB.Menu mnuCommands 
      Caption         =   "&Commands"
      Begin VB.Menu mnuWhoIsOnline 
         Caption         =   "&Who Is Online"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearCommandBuffer 
         Caption         =   "&Clear Command Buffer"
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'declare const and variables for the form
Private Const minWidth As Integer = 100
Private Const minHeight As Integer = 100
Private sngTimer As Single
Private Const iTimeout As Integer = 7

Private Sub Form_Load()

'clear the status bar
Me.StatusBar1.Panels.Clear

'set the status bar values
Me.StatusBar1.Panels.Add 1, "", "Not Connected", 0
Me.StatusBar1.Panels.Item(1).Bevel = sbrInset
Me.StatusBar1.Panels.Item(1).AutoSize = sbrSpring

Me.StatusBar1.Panels.Add 2
Me.StatusBar1.Panels.Item(2).AutoSize = sbrContents

Me.StatusBar1.Panels.Add 3, "", , sbrTime
Me.StatusBar1.Panels.Item(3).AutoSize = sbrContents

UpdateTitleBar ("Chat - " & strUsername)
EnableDisable True

End Sub

Private Sub Form_Resize()

If Me.Height > 1300 Then
    'txtbuffer
    Me.txtBuffer.Left = 0
    Me.txtBuffer.Top = Me.Height - 1000 - 255
    Me.txtBuffer.Width = Me.Width - 100
    Me.txtBuffer.Height = 315
    
    'resizes user list
    Me.lstUsers.Left = Me.Width - 1500
    Me.lstUsers.Top = 50
    Me.lstUsers.Height = Me.txtBuffer.Top - 110
    Me.lstUsers.Width = 1365
    
    'resize the txtStream
    Me.RichTextBox1.Left = 0
    Me.RichTextBox1.Top = 50
    Me.RichTextBox1.Height = Me.lstUsers.Height
    Me.RichTextBox1.Width = Me.Width - Me.lstUsers.Width - 250
    

    
Else
    'txtbuffer
    Me.txtBuffer.Left = 0
    Me.txtBuffer.Top = minHeight
    Me.txtBuffer.Width = minWidth
    Me.txtBuffer.Height = 315
    
    'resize the txtStream
    Me.RichTextBox1.Left = 0
    Me.RichTextBox1.Top = 50
    Me.RichTextBox1.Height = minHeight
    Me.RichTextBox1.Width = minWidth
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
wsClient.Close
wsClient_Close
End
End Sub

Private Sub lstUsers_DblClick()
frmPrivate.lblUser = Me.lstUsers.List(Me.lstUsers.ListIndex)
frmPrivate.Show , Me
End Sub

Private Sub mnuClearCommandBuffer_Click()
'clear the commands
modClient.ClearArray
End Sub

Private Sub mnuConnect_Click()
Dim Response As Integer

'if the username is not blank
If strServer <> "" Then
    'if the port is not = 0
    If iPort <> 0 Then
        'set the server
        Me.wsClient.RemoteHost = strServer
        'set the port
        Me.wsClient.RemotePort = iPort
        'connect to the remote server
        Me.wsClient.Connect
        'make sure all connection events happen
        DoEvents
        'wait for connection
        WaitForConnect
    Else
        'display error
        Response = MsgBox("You must enter a port number.", 64, "Error")
    End If
Else
    'display error
    Response = MsgBox("You must enter a server.", 64, "Error")
End If
End Sub

Private Sub mnuDisconnect_Click()
'close winsock connection
wsClient.Close
EnableDisable True
Me.StatusBar1.Panels.Item(1) = "Not connected"
End Sub

Private Sub mnuExit_Click()
'unload form
Unload Me
End Sub

Private Sub mnuOptions_Click()
'show the options form modally
frmClientOptions.Show vbModal
End Sub

Private Sub mnuWhoIsOnline_Click()
'send the command to the server
If bConnected Then
    Me.wsClient.SendData "/WHO"
End If
End Sub


Private Sub Timer1_Timer()
If bConnected Then
    Me.wsClient.SendData "/WHO" & vbCrLf & "USER:" & strUsername
End If
Timer1.Interval = 0
End Sub

Private Sub txtBuffer_Change()
Dim strCommand As String
Dim iCurPos As Integer

'is the field blank?
If Me.txtBuffer <> "" Then
    'if not blank, then check if the Enter key was pressed
    'and make sure that the length is longer than just the enter key value
    'find the position of the Enter keypress
    iCurPos = InStr(1, Me.txtBuffer, Chr(10))
    'if enter keypress exists and the length of the text is longer than the
    'Enter keypress
    If iCurPos <> 0 And Len(txtBuffer) <> 2 Then
    
        'check if the line contains an enter keystroke
        If iCurPos <> 0 Then
            'shorten the command to the left of the enter keypress
            'this is only noticed when the Enter key is pressed in the middle
            'of a line rather than at the end
            strCommand = Left(Me.txtBuffer, Me.txtBuffer.SelStart)
            'add the item into the commands database without the Enter
            modClient.InsertItemToArray Left(strCommand, Len(strCommand) - 2)
            'reset the array index to 0
            iArrayIndex = -1
            'send the info
            SendBuffer strCommand
        End If
    Else
        'if only enter was entered, clear the field
        If Right(Me.txtBuffer, 1) = Chr(10) Then
            Me.txtBuffer = ""
        End If
    End If
Else
    'clear the field
    Me.txtBuffer = ""
End If
End Sub

Private Sub SendBuffer(strDataToSend As String)

Dim strTemp As String

'this is to issue a command to the server
'this is not passed on to other clients, but processed
'by the server
If Left(Me.txtBuffer, 2) = "__" Or Left(Me.txtBuffer, 1) = "/" Then
    'add the username of the client so the server knows who to send
    'it back to
    strTemp = Me.txtBuffer & "USER:" & strUsername
Else
    'regular text being sent to all clients
    'add the username and a distinguishing code (&&) to the server
    'strTemp = strUsername & "&&" & strDataToSend
    strTemp = strDataToSend
End If

'send the data
Me.wsClient.SendData strTemp
'clear the buffer
Me.txtBuffer = ""

End Sub

Private Sub txtBuffer_KeyDown(KeyCode As Integer, Shift As Integer)
'check if it's the up arrow or down arrow
If KeyCode = 38 Or KeyCode = 40 Then
    'if up arrow
    If KeyCode = 38 Then
        'get the previous item in the command array
        Me.txtBuffer = modClient.GetPrevItem
        'position the cursor at the end of the line
        Me.txtBuffer.SelStart = Len(Me.txtBuffer)
    End If
    'if down arrow
    If KeyCode = 40 Then
        'get the next item in the command array
        Me.txtBuffer = modClient.GetNextItem
        'position the cursor at the end of the line
        Me.txtBuffer.SelStart = Len(Me.txtBuffer)
    End If
End If
End Sub

Private Sub wsClient_Close()

DisplayServerInfo "You have been disconnected from the server"
'reset the statusbar
Me.StatusBar1.Panels.Item(1) = "Not Connected"
'close the winsock
Me.wsClient.Close

EnableDisable True
End Sub

Private Sub wsClient_Connect()
'set the status info to something useful
Me.StatusBar1.Panels.Item(1) = "Connected to " & strServer
'send the server your username to confirm that no one else
'is using it
Me.wsClient.SendData "__NAME" & strUsername
'set the timer interval so that it will poll the server for a list
'of logged on users.
'Timer1.Interval = 5000

End Sub

Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)

Dim strData As String
Dim strCommand As String
Dim strUser As String
Dim strInfo As String
Dim strFullData As String

'get the data
Me.wsClient.GetData strData

'check if the data is a command
If Left(strData, 7) <> "COMMAND" Then
    
    'parse out the username from the data
    strUser = Left(strData, InStr(1, strData, ":"))
    'parse out the rest of the text passed
    strInfo = Right(strData, Len(strData) - InStr(1, strData, ":"))
    
    'put username and data together
    strFullData = strUser & strInfo
    
    'check if the username that sent this text is the same as the client
    'if it is then display text with different colours than if it is a different
    'user that sent the text
    If UCase(Left(strUser, Len(strUser) - 1)) = UCase(strUsername) Then
        AddText strFullData, True
    Else
        AddText strFullData, False
    End If
Else
    'parse out the command
    strCommand = Right(strData, Len(strData) - 8)
    If Len(strCommand) > Len(strCommand) - Len(strUsername) Then
        strCommand = Left(strCommand, Len(strCommand) - Len(strUsername))
    End If
    
    'parse out the command again
    If Left(strData, 13) = "COMMAND_TASKS" Then
        strCommand = "TASKS"
    End If
    
    'SERVER COMMAND
    If Left(strData, 15) = "COMMAND_SERVER:" Then
        strCommand = "SERVER"
    End If
    
    'kicked
    If Left(strData, 12) = "COMMAND_KICK" Then
        strCommand = "KICK"
    End If
    
    'password requested
    If Left(strData, 16) = "COMMAND_PASSWORD" Then
        strCommand = "PASSWORD"
    End If
    
    'find the appropriate command and act accordingly
    Select Case strCommand
    Case "NEWNAME"
        Me.Timer1.Interval = 0
        ForceNewUser
    Case "TASKLIST"
        GetAllTasks
    Case "TASKS"
        DisplayTasks (strData)
    Case "SERVER"
        'check if this is a command from the server or just information
        If Right(strData, 11) = "FILLLISTBOX" Then
            FillListBox (strData)
        Else
            DisplayServerInfo (strData)
            'get username
            strUser = Mid(strData, 16, InStr(1, strData, " ") - 16)
            'remove person from list
            If InStr(1, strData, " has left the chat") <> 0 Then
                'add to list if not your self
                RemoveListItem Mid(strData, 16, InStr(1, strData, " ") - 16)
            End If
            'add person to list
            If InStr(1, strData, " has joined the chat") <> 0 Then
                If UCase(strUser) <> UCase(strUsername) Then
                    AddListItem Mid(strData, 16, InStr(1, strData, " has joined the chat") - 16)
                End If
            End If
            'password incorrect
            If InStr(1, strData, "Password incorrect") <> 0 Then
                Me.StatusBar1.Panels.Item(1) = "Password incorrect"
            End If
        End If
    Case "KICK"
        Me.wsClient.Close
        mnuDisconnect_Click
        AddText "You were kicked from the server.", False
    Case "PASSWORD"
        frmPassword.Show vbModal
    End Select
End If

End Sub

Private Sub ForceNewUser()

Dim Response As Integer

'show user that their username has been taken
Response = MsgBox(strUsername & " has already been selected as a username" & Chr(13) & "Please select a new one", vbOKOnly, "Invalid Username")
'clear the buffer
frmClient.txtBuffer = ""
'clear the main textbox
frmClient.RichTextBox1.Text = ""
'show the options screen
frmClientOptions.Show vbModal

End Sub

Private Sub GetAllTasks()

'get list of all open tasks
Dim i As Integer
Dim FullTaskList As String

'get the tasklist
FillTaskList Me.hwnd
'build a list of all active tasks, prefixed with __INFO for the server's sake
FullTaskList = "__INFO"
For i = 0 To NumTasks - 1
    FullTaskList = FullTaskList & Left(TaskList(i).TaskName, Len(TaskList(i).TaskName) - 1) & "***"
Next i

'send the task list to the server
Me.wsClient.SendData FullTaskList
DoEvents
End Sub

Public Sub UpdateTitleBar(strTitle)
'put new title on folder
Me.Caption = strTitle
End Sub

Private Sub AddText(strText As String, bSameUser As Boolean)
'bsameuser determines whether or not the user receiving the data also sent it
'if so, then it gets a different colour
'add text to text box
Dim lColour1 As Long
Dim lColour2 As Long
Dim strTemp As String
Dim strLog As String

If bSameUser = True Then
    lColour1 = lNameColour
    lColour2 = lTextColour
Else
    lColour1 = lVisitorName
    lColour2 = lVisitorText
End If

'add timestamp if requested
TimeStamp

If InStr(1, strText, ":") <> 0 Then
    'start the select at end of textbox
    Me.RichTextBox1.SelStart = Len(Me.RichTextBox1.Text)
    'set the colour for the username
    Me.RichTextBox1.SelColor = lColour1
    'set the username to bold
    Me.RichTextBox1.SelBold = True
    'add username to textbox
    Me.RichTextBox1.SelText = Left(strText, InStr(1, strText, ":") + 1)
    'prepare strLog to contain whole string for log
    strLog = Left(strText, InStr(1, strText, ":") + 1)
End If
'start the select at end of textbox
Me.RichTextBox1.SelStart = Len(Me.RichTextBox1.Text)
'set back to non-bold
Me.RichTextBox1.SelBold = False
'set the colour for the text
Me.RichTextBox1.SelColor = lColour2
'add the rest of the text
strTemp = Mid(strText, InStr(1, strText, ":") + 1) & vbCrLf
Me.RichTextBox1.SelText = strTemp

strLog = GetTime & strLog & strTemp
modLogfile.AddtoLogFile strLog
End Sub

Private Sub DisplayTasks(strTasks As String)
'add the tasks to the client window

Dim lColour1 As Long
Dim lColour2 As Long
Dim strTemp As String
Dim i As Integer
Dim strTask As String

'set colours to the same as the visitors text
lColour1 = lVisitorName
lColour2 = lVisitorText

'parse out useless data
strTemp = Right(strTasks, Len(strTasks) - 13)

'do until strTemp does not contain ***
i = 1
Do Until i = 0
    'find position in strTemp of ***
    i = InStr(1, strTemp, "***")
    'if there is a ***
    If i <> 0 Then
        'set strTask to the leftmost data until the ***
        strTask = Left(strTemp, i - 1)
        'shorten strTemp to not include the strTask we just extracted
        strTemp = Right(strTemp, Len(strTemp) - i - 2)
                
        'add timestamp if requested
        TimeStamp
        
        'start the select at end of textbox
        Me.RichTextBox1.SelStart = Len(Me.RichTextBox1.Text)
        'set back to non-bold
        Me.RichTextBox1.SelBold = False
        'set the colour for the text
        Me.RichTextBox1.SelColor = lColour2
        'add to logfile
        modLogfile.AddtoLogFile GetTime & strTask
        'add the rest of the text
        'strTemp = Mid(strText, InStr(1, strText, ":") + 1) & vbCrLf
        Me.RichTextBox1.SelText = strTask & vbCrLf
        'start the select at end of textbox
        Me.RichTextBox1.SelStart = Len(Me.RichTextBox1.Text)
    End If
Loop

End Sub

Private Sub DisplayServerInfo(strText As String)

Dim lColour As Long
Dim strTemp As String
Dim strUsers As String


lColour = 8484992
strText = Right(strText, Len(strText) - 15)

'add timestamp if requested
TimeStamp
'add to log file
modLogfile.AddtoLogFile GetTime & strText

'start the select at end of textbox
Me.RichTextBox1.SelStart = Len(Me.RichTextBox1.Text)
'set back to non-bold
Me.RichTextBox1.SelBold = False
'set the colour for the text
Me.RichTextBox1.SelColor = lColour
'add the rest of the text
Me.RichTextBox1.SelText = strText & vbCrLf
'start the select at end of textbox
Me.RichTextBox1.SelStart = Len(Me.RichTextBox1.Text)

End Sub

Private Sub EnableDisable(bEnabled As Boolean)
'enable and disable buttons at connect/disconnect times

'this code will set everything up in the initial mode
'then if bEnabled = false, it will go to the connected mode
Me.mnuDisconnect.Enabled = Not bEnabled
Me.mnuConnect.Enabled = bEnabled
Me.mnuWhoIsOnline.Enabled = Not bEnabled

Me.txtBuffer.Enabled = Not bEnabled

frmClientOptions.txtPort.Enabled = bEnabled
frmClientOptions.txtServer.Enabled = bEnabled
frmClientOptions.txtUsername.Enabled = bEnabled

Me.mnuClearCommandBuffer.Enabled = Not bEnabled


End Sub
Private Sub DisableAll()
'enable and disable buttons at connect/disconnect times

'this code will set everything disabled

Me.mnuDisconnect.Enabled = False
Me.mnuConnect.Enabled = False
Me.mnuWhoIsOnline.Enabled = False

Me.txtBuffer.Enabled = False

frmClientOptions.txtPort.Enabled = False
frmClientOptions.txtServer.Enabled = False
frmClientOptions.txtUsername.Enabled = False

Me.mnuClearCommandBuffer.Enabled = False

End Sub
Private Sub WaitForConnect()

'start counting the time before timing out the connection attempt
setTimer
'disable all menu items
DisableAll
'while we haven't timed out and the winsock control is not connected
While getTimer And Me.wsClient.State <> 7
    'update the status bar according to what's happening
    ConnectionStatus
    DoEvents
Wend

If Not getTimer Then
    'connect failed
    Me.StatusBar1.Panels.Item(1) = "Connect failed"
    Me.wsClient.Close
    Me.txtBuffer = ""
    EnableDisable True
Else
    'connect successful
    EnableDisable False
End If

Me.Timer1.Interval = 0
'clear statusbar item
Me.StatusBar1.Panels.Item(2) = ""
    
End Sub
Private Sub setTimer()
    'set sngtimer to time
    sngTimer = Timer
End Sub

Private Function getTimer() As Boolean
'returns whether or not we've timed out
    'if sngTimer + the timeout amount is <= timer, then we have timed out
    If sngTimer + iTimeout <= Timer Then
        getTimer = False
    Else
        getTimer = True
    End If
End Function

Private Sub ConnectionStatus()
'put the winsock state into the status bar
Dim strStatus As String
Dim i As Integer

i = Me.wsClient.State

Select Case i
    Case 0
        strStatus = "Closed"
    Case 1
        strStatus = "Open"
    Case 2
        strStatus = "Listening..."
    Case 3
        strStatus = "Connection Pending"
    Case 4
        strStatus = "Resolving Host"
    Case 5
        strStatus = "Host Resolved"
    Case 6
        strStatus = "Connecting..."
    Case 7
        strStatus = "Connected"
    Case 8
        strStatus = "Peer is closing connection"
    Case 9
        strStatus = "Socket Error"
End Select

Me.StatusBar1.Panels.Item(1) = strStatus
End Sub

Private Sub RemoveListItem(strItem As String)
'remove the passed text from the listbox
Dim i As Integer

For i = 0 To Me.lstUsers.ListCount - 1
    If Me.lstUsers.List(i) = strItem Then
        Me.lstUsers.RemoveItem (i)
        i = Me.lstUsers.ListCount
    End If
Next i
    
End Sub

Private Sub AddListItem(strItem As String)
'add strItem to list
Me.lstUsers.AddItem strItem
End Sub

Private Sub FillListBox(strData As String)
'fill the list box with names
Dim i As Integer
Dim strNames(50) As String
Dim strTemp As String
Dim Counter As Integer

'trim out the FILLLISTBOX
strTemp = Left(strData, Len(strData) - 11)
'trim out COMMAND_SERVER:
strTemp = Right(strTemp, Len(strTemp) - 15)
'add a vbcrlf
strTemp = strTemp & vbCrLf

Me.lstUsers.Clear

i = 1
'trim out and add all names to the array
Do Until i = 0
    i = InStr(1, strTemp, vbCrLf)
    If i <> 0 Then
        strNames(Counter) = Left(strTemp, InStr(1, strTemp, vbCrLf) - 1)
        strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, vbCrLf) - 1)
        Counter = Counter + 1
    End If
Loop

'add names to listbox, might sort them eventually  =)
For i = 0 To Counter - 1
    AddListItem strNames(i)
Next i
End Sub

Private Sub TimeStamp()

If bTimestamp Then

    'if timestamp is selected then add the time
    'start the select at end of textbox
    Me.RichTextBox1.SelStart = Len(Me.RichTextBox1.Text)
    'set the colour for the time
    Me.RichTextBox1.SelColor = lTime
    'set the time to bold
    Me.RichTextBox1.SelBold = True
    'set the time to italic
    Me.RichTextBox1.SelItalic = True
    'add time to textbox
    Me.RichTextBox1.SelText = Time & "  "
    'set bold and italic to false
    Me.RichTextBox1.SelItalic = False
    Me.RichTextBox1.SelBold = False
    
End If
End Sub

Private Function bConnected() As Boolean
'returns true if the winsock control is connected to something
If Me.wsClient.State = 7 Then
    bConnected = True
End If
End Function

Private Function GetTime() As String
If bTimestamp Then
    GetTime = Time
End If
End Function
