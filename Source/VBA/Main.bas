Attribute VB_Name = "Main"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Main
'
' Description:
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
'   2.00.0000:
'     (1) Reset change history.
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

Public Const wsaApplicationName As String = "WorshipServiceAssistant"
Public Const wsaApplicationNamePretty As String = "Worship Service Assistant"
Public Const wsaApplicationVersion As String = "2.00.0000"
Public Const wsaApplicationAuthor As String = "Paul Bender"
Public Const wsaApplicationCopyright As String = "Copyright (c) 2000,2001,2002 Paul Bender"
Public Const wsaApplicationHomepage As String = "http://home.san.rr.com/benderfamily/software/wsa/"
Public Const wsaApplicationEmail As String = "mailto:pbender@alumni.ucsd.edu"


'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Public Variables.
'===============================================================================

Public Application As WorshipServiceAssistant.Application


'===============================================================================
' Private Variables.
'===============================================================================


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
'-------------------------------------------------------------------------------
Public Sub Auto_Open _
( _
)
    Auto_Close
    
    If (Not mblnOperatingSystemValid) Then Exit Sub
    If (Not mblnApplicationValid) Then Exit Sub
    
    Set WorshipServiceAssistant.Application = New WorshipServiceAssistant.Application
    WorshipServiceAssistant.Application.Initialize PowerPoint.Application
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'-------------------------------------------------------------------------------
Public Sub Auto_Close _
( _
)
    If (Not WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Application.Terminate
    End If
    Set WorshipServiceAssistant.Application = Nothing
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Returns:
'-------------------------------------------------------------------------------
Private Function mblnOperatingSystemValid _
( _
) As Boolean
    Dim strApplicationOperatingSystem As String
    Dim strPrompt As String
    Dim strTitle As String
    Dim intResponse As VBA.VbMsgBoxResult
    
    ' Assume that add-in has been loaded on an invalid operating system.
    mblnOperatingSystemValid = False
    
    ' Get operating system.
    strApplicationOperatingSystem = PowerPoint.Application.OperatingSystem
    
    ' Check operating system.  Only consider the operating system valid
    ' if the add-in has been tested on the operating system or the user
    ' indicates the operating system is valid.
    Select Case strApplicationOperatingSystem
        Case "Windows (32-bit) 5.01"
            mblnOperatingSystemValid = True
        Case Else
            strPrompt = _
                "The 'Worship Service Assistant' add-in " & _
                "has not been tested on " & _
                "'" & strApplicationOperatingSystem & "'" & ".  " & _
                VBA.vbNewLine & _
                "Would you like to run it anyway?"
            strTitle = _
                "'Worship Service Assistant' warning"
            intResponse = _
                VBA.MsgBox(strPrompt, VBA.vbYesNo, strTitle)
            If (intResponse = VBA.vbYes) Then
                mblnOperatingSystemValid = True
            Else
                mblnOperatingSystemValid = False
            End If
    End Select
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Returns:
'-------------------------------------------------------------------------------
Private Function mblnApplicationValid _
( _
) As Boolean
    Dim strApplicationName As String
    Dim strApplicationVersion As String
    Dim strPrompt As String
    Dim strTitle As String
    Dim intResponse As VBA.VbMsgBoxResult
    
    ' Assume that add-in has been loaded on an invalid application or
    ' application version.
    mblnApplicationValid = False
    
    ' Get operating system, application and application version.
    strApplicationName = PowerPoint.Application.Name
    strApplicationVersion = PowerPoint.Application.Version
    
    ' Check operating system, application and application version.
    ' Only allow the add-in to start if they are all valid.
    Select Case strApplicationName & " " & strApplicationVersion
        Case "Microsoft PowerPoint 10.0"
            mblnApplicationValid = True
        Case Else
            strPrompt = _
                "The 'Worship Service Assistant' add-in " & _
                "has not been tested on " & _
                "'" & strApplicationName & " " & _
                strApplicationVersion & "'" & ".  " & _
                VBA.vbNewLine & _
                "Would you like to run it anyway?"
            strTitle = _
                "'Worship Service Assistant' warning"
            intResponse = _
                VBA.MsgBox(strPrompt, VBA.vbYesNo, strTitle)
            If (intResponse = VBA.vbYes) Then
                mblnApplicationValid = True
            Else
                mblnApplicationValid = False
            End If
    End Select
End Function
