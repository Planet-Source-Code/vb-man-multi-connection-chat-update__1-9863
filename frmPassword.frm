VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   675
      Width           =   990
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1125
      TabIndex        =   1
      Top             =   225
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   240
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   765
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()

frmClient.wsClient.SendData "__PASSWORD" & Me.txtPassword
Unload Me
End Sub
