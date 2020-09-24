VERSION 5.00
Begin VB.Form frmPrivate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Private Message"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkOpen 
      Caption         =   "Keep Window Open After Sending Message"
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   1425
      Width           =   3540
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   1050
      Width           =   915
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   75
      TabIndex        =   1
      Top             =   600
      Width           =   4440
   End
   Begin VB.Label lblUser 
      Height          =   240
      Left            =   1950
      TabIndex        =   3
      Top             =   150
      Width           =   915
   End
   Begin VB.Label lblTo 
      Caption         =   "Send Private Message to"
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   1890
   End
End
Attribute VB_Name = "frmPrivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSend_Click()
frmClient.wsClient.SendData "__PRIVATE" & UCase(Me.lblUser) & "&&" & Me.txtSend
If Me.chkOpen.Value = 0 Then
    Unload Me
Else
    Me.txtSend = ""
End If

End Sub

Private Sub Form_Load()
Me.lblUser.Caption = frmClient.lstUsers.List(frmClient.lstUsers.ListIndex)
End Sub
