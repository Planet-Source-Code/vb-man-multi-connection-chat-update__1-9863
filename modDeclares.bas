Attribute VB_Name = "modDeclares"
Option Explicit
'
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINEFROMCHAR = &HC9

Public Sub InsertText(sText As String, frmForm As Form)
    With frmForm.txtPrompt
        If .SelStart <> 0 Then .SelStart = Len(.Text)
        On Error Resume Next
    
        .SelText = vbCrLf & "--> " & sText & vbCrLf
    End With
End Sub

