Attribute VB_Name = "modTaskStuff"
Option Explicit

' API Constants
Const WS_MINIMIZE = &H20000000 ' Style bit 'is minimized'
Const HWND_TOP = 0 ' Move to top of z-order
Const SWP_NOSIZE = &H1 ' Do not re-size window
Const SWP_NOMOVE = &H2 ' Do not reposition window
Const SWP_SHOWWINDOW = &H40 ' Make window visible/active
Const GW_HWNDFIRST = 0 ' Get first Window handle
Const GW_HWNDNEXT = 2 ' Get next window handle
Const GWL_STYLE = (-16) ' Get Window's style bits
Const SW_RESTORE = 9 ' Restore window

' The following constants will be combined to define properties
' of a 'normal' task top-level window. Any window with ' these set will be
' included in the list:
Const WS_VISIBLE = &H10000000 ' Window is not hidden
Const WS_BORDER = &H800000 ' Window has a border
' Other bits that are normally set include:
Const WS_CLIPSIBLINGS = &H4000000 ' can clip windows
Const WS_THICKFRAME = &H40000 ' Window has thick border
Const WS_GROUP = &H20000 ' Window is top of group
Const WS_TABSTOP = &H10000 ' Window has tabstop

' API Functions Definition
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

' Public Task Item Structure
Public Type TASK_STRUCT
    TaskName As String
    TaskID As Long
End Type

'Structure filled by FillTaskList Sub call
Public TaskList(1000) As TASK_STRUCT
Public NumTasks As Long

' Returns if a Process is a Visible Window
Public Function IsTask(hwndTask As Long) As Boolean
    Dim WndStyle As Long
    Const IsTaskStyle = WS_VISIBLE Or WS_BORDER

    WndStyle = GetWindowLong(hwndTask, GWL_STYLE)
    If (WndStyle And IsTaskStyle) = IsTaskStyle Then IsTask = True
End Function

' Fills the Task structure with captions and hWnd of all active programs
Public Sub FillTaskList(hwnd As Long)
    Dim hwndTask As Long
    Dim intLen As Long
    Dim strTitle As String
    Dim cnt As Integer

    cnt = 0
    ' process all top-level windows in master window list
    hwndTask = GetWindow(hwnd, GW_HWNDFIRST) ' get first window
    Do While hwndTask ' repeat for all windows
        If hwndTask <> hwnd And IsTask(hwndTask) Then
            intLen = GetWindowTextLength(hwndTask) + 1 ' Get length
            strTitle = Space(intLen) ' Get caption
            intLen = GetWindowText(hwndTask, strTitle, intLen)
            If intLen > 0 Then ' If we have anything, add it
                TaskList(cnt).TaskName = strTitle
                TaskList(cnt).TaskID = hwndTask
                cnt = cnt + 1
            End If
        End If
        hwndTask = GetWindow(hwndTask, GW_HWNDNEXT)
    Loop
    NumTasks = cnt
End Sub

' Give focus to another Task
Public Sub SwitchTo(hwnd As Long)
    Dim ret As Long
    Dim WStyle As Long ' Window Style bits

   ' Get style bits for window
    WStyle = GetWindowLong(hwnd, GWL_STYLE)
    ' If minimized do a restore
    If WStyle And WS_MINIMIZE Then
        ret = ShowWindow(hwnd, SW_RESTORE)
    End If
    ' Move window to top of z-order/activate; no move/resize
    ret = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
End Sub


