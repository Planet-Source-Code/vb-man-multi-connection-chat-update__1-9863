Attribute VB_Name = "modClient"
Option Explicit
Public strServer As String
Public iPort As Integer
Public strUsername
Public lNameColour As Long
Public lTextColour As Long
Public lVisitorName As Long
Public lVisitorText As Long
Public strCommands(50) As String
Public strTemp(50) As String
Public iArrayIndex As Integer
Public bTimestamp As Boolean
Public lTime As Long
Public bLog As Boolean
Public strLogFile As String

Private Sub Main()
frmClientOptions.LoadOptions
frmClient.Show
End Sub


Public Sub InsertItemToArray(strItem As String)
'insert strItem into spot 0 of the array and move
'all the other items down one

Dim i As Integer
Dim j As Integer
Dim strTemp(50) As String

'set first item to strItem
strTemp(0) = strItem

'move items up one spot
For i = 1 To 50
    strTemp(i) = strCommands(i - 1)
Next i

'move data back into strCommands array
For i = 0 To 50
    strCommands(i) = strTemp(i)
Next i

End Sub

Function GetPrevItem() As String
'get the previous item entered
If iArrayIndex < FindHighestUsedItem - 1 Then
    iArrayIndex = iArrayIndex + 1
    GetPrevItem = strCommands(iArrayIndex)
Else
    GetPrevItem = "*** Beginning of Commands ***"
    iArrayIndex = FindHighestUsedItem
End If
End Function

Function FindHighestUsedItem() As Integer

'returns the highest used item in the commands array
Dim i As Integer

For i = 0 To 50
    If strCommands(i) = "" Then
        FindHighestUsedItem = i
        i = 51
    End If
Next i

If FindHighestUsedItem = 0 Then
    FindHighestUsedItem = 50
End If
End Function

Function GetNextItem() As String
'get the next item in array
If iArrayIndex > 0 Then
    iArrayIndex = iArrayIndex - 1
    GetNextItem = strCommands(iArrayIndex)
Else
    iArrayIndex = -1
    GetNextItem = ""
End If
End Function

Public Sub ClearArray()
Dim i As Integer

For i = 0 To 50
    strCommands(i) = ""
    strTemp(i) = ""
Next i

iArrayIndex = 0
End Sub
