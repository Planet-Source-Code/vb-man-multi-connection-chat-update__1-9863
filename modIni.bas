Attribute VB_Name = "modIni"
Option Explicit


Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Global Const MDG_1K_BUFFER = 1024
    
'**************************************
' Name: Get & Set .ini Values
' Description:GetInIValue:The function r
'     eturns a string representing the text which
'The entry 'strEntry' has been Set To In 'strSection' of
'strINIFileName' '.INI' file.
'SetINIValue: ' This Function sets the entry 'strEntry' To the value of
'strSetting' in the 'strSection' section
'     of 'strINIFileName' file. The function r
'     eturns an integer which is 'True' if the
'     entry was written to the '.INI' file or
'     'False' if the entry could not be writte
'     n.
' By: Killcrazy
'
'
' Inputs:SetINIValue:
'strSection = The section from which the entry value is
'To be retrieved i.e. '[SECTION]'.
'strEntry = The name of an entry In the '.INI' file.
'strSetting = The text To be saved in the '.INI' file's
'strEntry' entry.
'strINIFileName = The name of the '.INI' file.
'GetINIValue:
'strSection = The section from which the entry value is
'To be retrieved i.e. '[SECTION]'.
'strEntry = The name of an entry In the '.INI' file.
'strINIFileName = The name of the '.INI' file.
'
' Returns:GetINIValue returns the requir
'     ed value
'SetINIValue returns an Integer indicating True or False
'
'Assumes:None
'
'Side Effects:none
'
'Warranty:
'code provided by Planet Source Code(tm)
'     (http://www.Planet-Source-Code.com) 'as
'     is', without warranties as to performanc
'     e, fitness, merchantability,and any othe
'     r warranty (whether expressed or implied
'     ).
'Terms of Agreement:
'By using this source code, you agree to
'     the following terms...
' 1) You may use this source code in per
'     sonal projects and may compile it into a
'     n .exe/.dll/.ocx and distribute it in bi
'     nary format freely and with no charge.
' 2) You MAY NOT redistribute this sourc
'     e code (for example to a web site) witho
'     ut written permission from the original
'     author.Failure to do so is a violation o
'     f copyright laws.
' 3) You may link to this code from anot
'     her website, provided it is not wrapped
'     in a frame.
' 4) The author of this code may have re
'     tained certain additional copyright righ
'     ts.If so, this is indicated in the autho
'     r's description.
'**************************************



Public Function GetINIValue(ByVal strSection As String, ByVal strEntry As String, ByVal strINIFileName As String) As String
    Dim strReturnString As String
    Dim lngPointer As Long
    ' Set error handler for case when the IN
    '     I value cannot be read.
    On Error GoTo Error_Handler
    ' Make the return string of fixed length
    '
    strReturnString = String(MDG_1K_BUFFER, " ")
    ' Get the string from the INI file
    lngPointer = GetPrivateProfileString(strSection, strEntry, "", strReturnString, MDG_1K_BUFFER, strINIFileName)
    ' If we found an INI item, return it


    If lngPointer > 0 Then
        GetINIValue = Left$(strReturnString, lngPointer)
    End If
    ' Exit Early to avoid error handler.
    Exit Function
    ' Cannot read from INI file.
Error_Handler:
    ' Raise an error.
    Err.Raise Err.Number, "GetINIValue", "Cannot read from INI File.", Err.Description
    ' Reset normal error checking.
    On Error GoTo 0
    ' Resume via fail exit point.
    Resume Exit_GetINIValue
    ' Error handler exit point.
Exit_GetINIValue:
    ' Could not read from INI - return recor
    '     ded 'strEntry'.
    GetINIValue = strEntry
End Function
'---------------------------------


Public Function SetINIValue(ByVal strSection As String, ByVal strEntry As String, ByVal strSetting As String, ByVal strINIFileName As String) As Integer
    Dim lngReturn As Long
    ' Set error handler for case when the INI value cannot be written.
    On Error GoTo Error_Handler
    ' Write restart point setting to the INI' file.
    lngReturn = WritePrivateProfileString(strSection, strEntry, strSetting, strINIFileName)
    ' Return True or False depending upon API call validity.
    SetINIValue = IIf(lngReturn = 0, False, True)
    ' Exit Early to avoid error handler.
    Exit Function
    ' Cannot write to INI file.
Error_Handler:
    ' Raise an error.
    Err.Raise Err.Number, "SetINIValue", "Cannot write To INI file.", Err.Description
    ' Reset normal error checking.
    On Error GoTo 0
    ' Resume via fail exit point.
    Resume Exit_SetINIValue
    ' Error handler exit point.
Exit_SetINIValue:
    ' Could not write to INI - return False.
    '
    SetINIValue = False
End Function

Public Sub GetAccountNames(frmOpts As Form, frmMainForm As Form)
Dim strInput As String
Dim sExists As String
Dim bAccounts As Boolean
Dim strFilename As String

strFilename = App.Path & "\mail.ini"
'sExists = FileExist(strFilename)

'clear listbox
frmOpts.lstAccounts.Clear
'clear combo box
frmMainForm.cboAccounts.Clear

On Error GoTo EH
Open strFilename For Input As #1

Do Until EOF(1)
    
    Input #1, strInput
    If InStr(1, strInput, "[") <> 0 And InStr(1, strInput, "]") <> 0 Then
    
        frmOpts.lstAccounts.AddItem Mid(strInput, 2, Len(strInput) - 2)
        frmMainForm.cboAccounts.AddItem Mid(strInput, 2, Len(strInput) - 2)
        
    End If
        
Loop
Close #1

frmMainForm.cboAccounts.Text = frmMainForm.cboAccounts.List(0)

EH:
End Sub
