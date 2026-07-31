Attribute VB_Name = "Main"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Main
'
' Description:
'   This is a PowerPoint 9.0 (aka 2000) and 10.0 (aka XP) add-in.
'   It is intended for use in dual monitor PowerPoint presentations.
'
' Author:
'   Paul Bender <pbender@alumni.ucsd.edu>
'
' Copyright:
'   Copyright (c) 2000,2001,2002 Paul Bender
'
'   All rights reserved.
'
'   Permission is hereby granted, free of charge, to any person obtaining a
'   copy of this software and associated documentation files (the
'   "Software"), to deal in the Software without restriction, including
'   without limitation the rights to use, copy, modify, merge, publish,
'   distribute, and/or sell copies of the Software, and to permit persons
'   to whom the Software is furnished to do so, provided that the above
'   copyright notice(s) and this permission notice appear in all copies of
'   the Software and that both the above copyright notice(s) and this
'   permission notice appear in supporting documentation.
'
'   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
'   OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
'   MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT
'   OF THIRD PARTY RIGHTS. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
'   HOLDERS INCLUDED IN THIS NOTICE BE LIABLE FOR ANY CLAIM, OR ANY SPECIAL
'   INDIRECT OR CONSEQUENTIAL DAMAGES, OR ANY DAMAGES WHATSOEVER RESULTING
'   FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT,
'   NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION
'   WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
'
'   Except as contained in this notice, the name of a copyright holder
'   shall not be used in advertising or otherwise to promote the sale, use
'   or other dealings in this Software without prior written authorization
'   of the copyright holder.
'
' Change History:
'   1.01.0004:
'     Added a check for Microsoft Office 10.0.
'   1.01.0000:
'     Added a text banner above the slide show.
'   1.00.0000:
'     Initial revision.
'===============================================================================


'===============================================================================
' Options.
'===============================================================================
Option Explicit
Option Compare Text
Option Base 0


'===============================================================================
' Public Constants.
'===============================================================================


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'   Automatically loads the project.
'-------------------------------------------------------------------------------
Public Sub Auto_Open()
    '
    ' Validate operating system.
    '
    If (ValidOperatingSystem = False) Then
        Exit Sub
    End If
    
    '
    ' Validate application.
    '
    If (ValidApplication = False) Then
        Exit Sub
    End If
    
    Project_Load
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Automatically unloads the project.
'-------------------------------------------------------------------------------
Public Sub Auto_Close()
    Project_Unload
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'   This function makes checks to make sure that the add-in is being run on a
'   valid operating system.  If the add-in has not been tested on the operating
'   system, then the user is prompted.
'-------------------------------------------------------------------------------
Private Function ValidOperatingSystem() As Boolean
    Dim ApplicationOperatingSystem As String
    
    '
    ' Assume that add-in has been loaded on an invalid operating system.
    '
    ValidOperatingSystem = False
    
    '
    ' Get operating system.
    '
    ApplicationOperatingSystem = Application.OperatingSystem
    
    '
    ' Check operating system.  Only consider the operating system valid
    ' if the add-in has been tested on the operating system or the user
    ' indicates the operating system is valid.
    '
    Select Case ApplicationOperatingSystem
        Case "Windows (32-bit) 5.00"
            ValidOperatingSystem = True
        Case "Windows (32-bit) 5.01"
            ValidOperatingSystem = True
        Case Else
            Dim msgPrompt As String
            Dim msgTitle As String
            Dim msgResponse As VBA.VbMsgBoxResult
            msgPrompt = _
                "The 'Worship Service Assistant' add-in " & _
                "has not been tested on " & _
                "'" & ApplicationOperatingSystem & "'" & ".  " & _
                Chr(13) & Chr(10) & _
                "Would you like to run it anyway?"
            msgTitle = _
                "'Worship Service Assistant' warning"
            msgResponse = _
                MsgBox(msgPrompt, vbYesNo, msgTitle)
            If (msgResponse = vbYes) Then
                ValidOperatingSystem = True
            Else
                ValidOperatingSystem = False
            End If
    End Select
End Function

'-------------------------------------------------------------------------------
' Description:
'   This function checks to make sure that the add-in is being run on a
'   valid application.  If the add-in has not been tested on the application,
'   then the user is prompted.
'-------------------------------------------------------------------------------
Private Function ValidApplication() As Boolean
    Dim ApplicationName As String
    Dim ApplicationVersion As String
    
    '
    ' Assume that add-in has been loaded on an invalid application or
    ' application version.
    '
    ValidApplication = False
    
    '
    ' Get operating system, application and application version.
    '
    ApplicationName = Application.Name
    ApplicationVersion = Application.Version
    
    '
    ' Check operating system, application and application version.
    ' Only allow the add-in to start if they are all valid.
    '
    Select Case ApplicationName & " " & ApplicationVersion
        Case "Microsoft PowerPoint 9.0"
            ValidApplication = True
        Case "Microsoft PowerPoint 10.0"
            ValidApplication = True
        Case Else
            Dim msgPrompt As String
            Dim msgTitle As String
            Dim msgResponse As VBA.VbMsgBoxResult
            msgPrompt = _
                "The 'Worship Service Assistant' add-in " & _
                "has not been tested on " & _
                "'" & ApplicationName & " " & _
                ApplicationVersion & "'" & ".  " & _
                Chr(13) & Chr(10) & _
                "Would you like to run it anyway?"
            msgTitle = _
                "'Worship Service Assistant' warning"
            msgResponse = _
                MsgBox(msgPrompt, vbYesNo, msgTitle)
            If (msgResponse = vbYes) Then
                ValidApplication = True
            Else
                ValidApplication = False
            End If
    End Select
End Function
