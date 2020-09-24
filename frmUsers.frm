VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Users"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   3000
      TabIndex        =   12
      Top             =   2325
      Width           =   990
   End
   Begin VB.Frame fraNewPerson 
      Height          =   2115
      Left            =   75
      TabIndex        =   6
      Top             =   75
      Visible         =   0   'False
      Width           =   2790
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   2775
         TabIndex        =   11
         Top             =   1125
         Width           =   990
      End
      Begin VB.TextBox txtNewPassword 
         Height          =   285
         Left            =   1275
         TabIndex        =   8
         Top             =   750
         Width           =   2490
      End
      Begin VB.TextBox txtNewUsername 
         Height          =   285
         Left            =   1275
         TabIndex        =   7
         Top             =   375
         Width           =   2490
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
         Height          =   165
         Left            =   75
         TabIndex        =   10
         Top             =   750
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   240
         Left            =   75
         TabIndex        =   9
         Top             =   450
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   2700
      Width           =   990
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Top             =   1125
      Width           =   990
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   315
      Left            =   1950
      TabIndex        =   3
      Top             =   1125
      Width           =   990
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1950
      TabIndex        =   2
      Top             =   600
      Width           =   2040
   End
   Begin VB.ListBox lstUsers 
      Height          =   2790
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   240
      Left            =   1950
      TabIndex        =   1
      Top             =   225
      Width           =   2565
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click()

Me.cmdNew.Visible = True
Me.cmdUpdate.Visible = True
Me.cmdClose.Visible = True
Me.cmdDelete.Visible = True

modUsers.AddUser Me.txtNewUsername, Me.txtNewPassword
modUsers.LoadListBox Me.lstUsers

Me.fraNewPerson.Visible = False

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
modUsers.DeleteUser Me.lstUsers.List(Me.lstUsers.ListIndex)
Me.txtPassword = ""
modUsers.LoadListBox Me.lstUsers
End Sub

Private Sub cmdNew_Click()

Me.txtNewPassword = ""
Me.txtNewUsername = ""

Me.fraNewPerson.Height = 3015
Me.fraNewPerson.Width = 4065

Me.cmdNew.Visible = False
Me.cmdUpdate.Visible = False
Me.cmdClose.Visible = False
Me.cmdDelete.Visible = False

Me.fraNewPerson.Visible = True
Me.txtNewUsername.SetFocus

End Sub

Private Sub cmdUpdate_Click()
modUsers.UpdateUser Me.lstUsers.List(Me.lstUsers.ListIndex), Me.txtPassword
End Sub

Private Sub Form_Load()

modUsers.LoadListBox Me.lstUsers

End Sub

Private Sub lstUsers_Click()

Me.txtPassword = modUsers.GetPassword(Me.lstUsers.List(Me.lstUsers.ListIndex))

End Sub

Private Sub txtNewPassword_GotFocus()
Me.txtNewPassword.SelStart = 0
Me.txtNewPassword.SelLength = Len(Me.txtNewPassword)
End Sub

Private Sub txtNewUsername_GotFocus()
Me.txtNewUsername.SelStart = 0
Me.txtNewUsername.SelLength = Len(Me.txtNewUsername)
End Sub

Private Sub txtPassword_GotFocus()
Me.txtPassword.SelStart = 0
Me.txtPassword.SelLength = Len(Me.txtPassword)
End Sub
