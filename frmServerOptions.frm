VERSION 5.00
Begin VB.Form frmServerOptions 
   Caption         =   "Options"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSecurity 
      Caption         =   "Enable security"
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   675
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1260
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2340
      TabIndex        =   2
      Top             =   1260
      Width           =   1035
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   525
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   240
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   435
   End
End
Attribute VB_Name = "frmServerOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

If StoreOptions Then
    frmServer.StatusBar1.Panels.Item(2) = "Port: " & iPort
    Unload Me
End If

End Sub

Private Sub Form_Load()

ReadOptions
End Sub

Private Sub txtPort_GotFocus()
Me.txtPort.SelStart = 0
Me.txtPort.SelLength = Len(Me.txtPort)
End Sub

Private Function StoreOptions() As Boolean

Dim i As Integer

StoreOptions = True

If IsNumeric(Me.txtPort) Then
    i = modIni.SetINIValue("Server", "Port", Me.txtPort, App.Path & "\chat.ini")
    iPort = Me.txtPort
Else
    StoreOptions = False
    i = MsgBox("You must enter a number between 1 and 65335", vbOKOnly, "Error")
End If

If StoreOptions Then
    If Me.chkSecurity.Value = 0 Then
        i = modIni.SetINIValue("Server", "Security", "False", App.Path & "\chat.ini")
    Else
        i = modIni.SetINIValue("Server", "Security", "True", App.Path & "\chat.ini")
    End If
End If

End Function

Private Sub ReadOptions()
Dim i As Integer
Dim strTemp As String
 
strTemp = modIni.GetINIValue("Server", "Port", App.Path & "\chat.ini")
Me.txtPort = Int(strTemp)

DoEvents
strTemp = modIni.GetINIValue("Server", "Security", App.Path & "\chat.ini")
If strTemp = "True" Then
    Me.chkSecurity.Value = 1
End If
End Sub
