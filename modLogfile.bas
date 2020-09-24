Attribute VB_Name = "modLogfile"
Option Explicit

Public Sub AddtoLogFile(strText As String)
Dim strFile As String

If bLog Then
    If Right(strText, 2) = vbCrLf Then
        strText = Left(strText, Len(strText) - 2)
    End If
    
    Open strLogFile For Append As #1
        Print #1, strText
    Close #1

End If
End Sub

