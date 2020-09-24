VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmClientOptions 
   Caption         =   "Options"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3240
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   5715
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "&Basic Info"
      TabPicture(0)   =   "frmClientOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtUsername"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtPort"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtServer"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Colours"
      TabPicture(1)   =   "frmClientOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "chkTimestamp"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "lblTime"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&Logging"
      TabPicture(2)   =   "frmClientOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkLog"
      Tab(2).Control(1)=   "txtLog"
      Tab(2).ControlCount=   2
      Begin VB.TextBox txtLog 
         Height          =   285
         Left            =   -74550
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   975
         Width           =   3615
      End
      Begin VB.CheckBox chkLog 
         Caption         =   "Log all events"
         Height          =   240
         Left            =   -74775
         TabIndex        =   5
         Top             =   600
         Width           =   1365
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   750
         TabIndex        =   1
         Top             =   600
         Width           =   3540
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   780
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   1005
         Width           =   675
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   765
         TabIndex        =   3
         Top             =   1440
         Width           =   3540
      End
      Begin VB.Frame Frame1 
         Caption         =   "Colours"
         Height          =   990
         Left            =   -74775
         TabIndex        =   9
         Top             =   825
         Width           =   3615
         Begin VB.Label lblNick 
            Caption         =   "Nick"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   150
            TabIndex        =   13
            Top             =   300
            Width           =   390
         End
         Begin VB.Label lblVisitorsNick 
            Caption         =   "Visitor's Nick"
            Height          =   165
            Left            =   150
            TabIndex        =   12
            Top             =   600
            Width           =   990
         End
         Begin VB.Label lblText 
            Caption         =   "Local Text"
            Height          =   165
            Left            =   1575
            TabIndex        =   11
            Top             =   300
            Width           =   840
         End
         Begin VB.Label lblVisitorsText 
            Caption         =   "Visitor's Text"
            Height          =   165
            Left            =   1575
            TabIndex        =   10
            Top             =   600
            Width           =   990
         End
      End
      Begin VB.CheckBox chkTimestamp 
         Caption         =   "Time Stamp all events"
         Height          =   195
         Left            =   -74775
         TabIndex        =   4
         Top             =   2070
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Server"
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   585
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Port"
         Height          =   255
         Left            =   225
         TabIndex        =   17
         Top             =   1005
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "Nick"
         Height          =   240
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label Label4 
         Caption         =   "Click to choose the colours of the following options"
         Height          =   240
         Left            =   -74700
         TabIndex        =   15
         Top             =   525
         Width           =   3915
      End
      Begin VB.Label lblTime 
         Caption         =   "Time Colour"
         Height          =   240
         Left            =   -74475
         TabIndex        =   14
         Top             =   2400
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   3495
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2340
      TabIndex        =   7
      Top             =   3495
      Width           =   1035
   End
End
Attribute VB_Name = "frmClientOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkLog_Click()
'if the check box is clicked, set the enabled value of the
'textbox accordingly
If Me.chkLog.Value = 1 Then
    Me.txtLog.Enabled = True
Else
    Me.txtLog.Enabled = False
End If
    
End Sub

Private Sub chkTimestamp_Click()
'if the check box is clicked, set the enabled value of the
'label accordingly
If Me.chkTimestamp.Value = 1 Then
    Me.lblTime.Enabled = True
Else
    Me.lblTime.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
'unload the form
Unload Me
End Sub

Private Sub cmdOK_Click()
'check if filename is valid
'save the options, then reload them so the globals contain the new values
If ValidName Then
    SaveOptions
    LoadOptions
    'unload the form
    Unload Me
End If
End Sub

Private Sub Form_Load()
'load the options
LoadOptions

End Sub

Private Sub lblNick_Click()
'let the user choose a colour and assign it to the corresponding global
frmClient.CommonDialog1.ShowColor
lNameColour = frmClient.CommonDialog1.Color
Me.lblNick.ForeColor = lNameColour
End Sub

Private Sub lblText_Click()
'let the user choose a colour and assign it to the corresponding global
frmClient.CommonDialog1.ShowColor
lTextColour = frmClient.CommonDialog1.Color
Me.lblText.ForeColor = lTextColour
End Sub

Private Sub lblTime_Click()
'let the user choose a colour and assign it to the corresponding global
frmClient.CommonDialog1.ShowColor
lTime = frmClient.CommonDialog1.Color
Me.lblTime.ForeColor = lTime
End Sub

Private Sub lblVisitorsNick_Click()
'let the user choose a colour and assign it to the corresponding global
frmClient.CommonDialog1.ShowColor
lVisitorName = frmClient.CommonDialog1.Color
Me.lblVisitorsNick.ForeColor = lVisitorName
End Sub

Private Sub lblVisitorsText_Click()
'let the user choose a colour and assign it to the corresponding global
frmClient.CommonDialog1.ShowColor
lVisitorText = frmClient.CommonDialog1.Color
Me.lblVisitorsText.ForeColor = lVisitorText
End Sub

Private Sub txtLog_GotFocus()
'select all text when control gets focus
Me.txtLog.SelStart = 0
Me.txtLog.SelLength = Len(Me.txtLog)
End Sub

Private Sub txtPort_GotFocus()
'select all text when control gets focus
Me.txtPort.SelStart = 0
Me.txtPort.SelLength = Len(Me.txtPort)
End Sub

Private Sub txtServer_GotFocus()
'select all text when control gets focus
Me.txtServer.SelStart = 0
Me.txtServer.SelLength = Len(Me.txtServer)
End Sub

Private Sub txtUsername_GotFocus()
'select all text when control gets focus
Me.txtUsername.SelStart = 0
Me.txtUsername.SelLength = Len(Me.txtUsername)
End Sub

Private Sub SaveOptions()
'write all options to an inifile
Dim i As Integer

If Me.chkTimestamp.Value = 1 Then
    i = modIni.SetINIValue("Options", "Timestamp", "True", App.Path & "\chat.ini")
Else
    i = modIni.SetINIValue("Options", "Timestamp", "False", App.Path & "\chat.ini")
End If
i = modIni.SetINIValue("Options", "NickColour", lNameColour, App.Path & "\chat.ini")
i = modIni.SetINIValue("Options", "TextColour", lTextColour, App.Path & "\chat.ini")
i = modIni.SetINIValue("Options", "VisitorColour", lVisitorName, App.Path & "\chat.ini")
i = modIni.SetINIValue("Options", "VisitorText", lVisitorText, App.Path & "\chat.ini")
i = modIni.SetINIValue("Options", "Server", Me.txtServer, App.Path & "\chat.ini")
i = modIni.SetINIValue("Options", "Port", Me.txtPort, App.Path & "\chat.ini")
i = modIni.SetINIValue("Options", "Username", Me.txtUsername, App.Path & "\chat.ini")
i = modIni.SetINIValue("Options", "TimeColour", lTime, App.Path & "\chat.ini")
If Me.chkLog.Value = 1 Then
    i = modIni.SetINIValue("Options", "Log", "True", App.Path & "\chat.ini")
Else
    i = modIni.SetINIValue("Options", "Log", "False", App.Path & "\chat.ini")
End If
i = modIni.SetINIValue("Options", "LogFile", Me.txtLog, App.Path & "\chat.ini")
End Sub

Public Sub LoadOptions()
'load all options from the ini file and then set the globals to the values
Dim i As Integer
Dim strTemp As String

'time stamp
strTemp = modIni.GetINIValue("Options", "Timestamp", App.Path & "\chat.ini")
If strTemp = "True" Then
    Me.chkTimestamp.Value = 1
    Me.lblTime.Enabled = True
    bTimestamp = True
Else
    Me.chkTimestamp.Value = 0
    Me.lblTime.Enabled = False
    bTimestamp = False
End If

'time stamp colour
strTemp = modIni.GetINIValue("Options", "TimeColour", App.Path & "\chat.ini")
If strTemp <> "" Then
    lTime = strTemp
Else
    lTime = 0
End If

'server
Me.txtServer = modIni.GetINIValue("Options", "Server", App.Path & "\chat.ini")
strServer = Me.txtServer

'port number
Me.txtPort = modIni.GetINIValue("Options", "Port", App.Path & "\chat.ini")
iPort = Me.txtPort

'username
Me.txtUsername = modIni.GetINIValue("Options", "Username", App.Path & "\chat.ini")
strUsername = Me.txtUsername

'nick colour
strTemp = modIni.GetINIValue("Options", "NickColour", App.Path & "\chat.ini")
If strTemp <> "" Then
    lNameColour = strTemp
Else
    lNameColour = 0
End If

'text colour
strTemp = modIni.GetINIValue("Options", "TextColour", App.Path & "\chat.ini")
If strTemp <> "" Then
    lTextColour = strTemp
Else
    lTextColour = 0
End If

'visitor nick colour
strTemp = modIni.GetINIValue("Options", "VisitorColour", App.Path & "\chat.ini")
If strTemp <> "" Then
    lVisitorName = strTemp
Else
    lVisitorName = 0
End If

'visitor text colour
strTemp = modIni.GetINIValue("Options", "VisitorText", App.Path & "\chat.ini")
If strTemp <> "" Then
    lVisitorText = strTemp
Else
    lVisitorText = 0
End If

'logging stuff
strTemp = modIni.GetINIValue("Options", "Log", App.Path & "\chat.ini")
If strTemp = "True" Then
    Me.chkLog.Value = 1
    Me.txtLog.Enabled = True
    bLog = True
Else
    Me.chkLog.Value = 0
    bLog = False
    Me.txtLog.Enabled = False
End If
strTemp = modIni.GetINIValue("Options", "LogFile", App.Path & "\chat.ini")
Me.txtLog = strTemp
strLogFile = strTemp

'assign colours to text on form
Me.lblNick.ForeColor = lNameColour
Me.lblText.ForeColor = lTextColour
Me.lblVisitorsNick.ForeColor = lVisitorName
Me.lblVisitorsText.ForeColor = lVisitorText
Me.lblTime.ForeColor = lTime

'set the status bar text
frmClient.StatusBar1.Panels.Item(2) = " " & strServer & ":" & iPort & " "

'set the title
frmClient.UpdateTitleBar ("Chat - " & strUsername)
End Sub

Private Function ValidName() As Boolean
'check if username is valid
Dim Response As Integer

ValidName = True
If InStr(1, Me.txtUsername, ":") <> 0 Then
    Response = MsgBox("Username cannot contain a colon (:)", 64, "Error")
    ValidName = False
End If

End Function
