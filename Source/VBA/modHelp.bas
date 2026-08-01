Attribute VB_Name = "modHelp"
'===============================================================================
' Name:
'   WorshipServiceAssistant.modHelp
'
' Description:
'   This module contains constants associated with the help system.
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
'   1.04.0003:
'     (1) Created the gTopicShow routine to put a better wrapper around the
'         showing of HTML help.
'     (2) Added the gGotoHomepage and gAuthorEmail routines
'   1.03.0002:
'     (1) Made changes to the source code so that it follows Microsoft's
'         Visual Basic coding conventions.
'   1.01.0007:
'     (1) Added KnownIssues help constants.
'   1.01.0001:
'     (1) Updated help constants.
'   1.01.0000:
'     (1) Removed unused help constants,
'         because I was tired of keeping them up to date.
'   1.00.0001:
'     (1) Removed context sensitive help constants, because VBA 6.0 does
'         not support HTML help context sensitive help.
'     (2) Added HTML help topic path constants.
'     (3) Added HTML Help control API, because the HTML Help control API
'         allows more control than the PowerPoint help interface.
'   1.00.0000:
'     Initial revision.
'===============================================================================


'===============================================================================
' Options.
'===============================================================================
Option Private Module
Option Explicit
Option Compare Text
Option Base 0


'===============================================================================
' Public Constants.
'===============================================================================

Public Enum HelpTopic
    WSA = 0
    WSAHowTo = 1
    WSACommandBar = 2
    WSACommandBarNavigator = 3
    WSAKnownIssues = 4
    WSACopyrightPermission = 5
    WSAHistory = 6
End Enum


'===============================================================================
' Private Constants.
'===============================================================================

' HTML Help (HH_) constants.
Private Const HH_DISPLAY_TOPIC As Long = &H0
Private Const HH_HELP_CONTEXT  As Long = &HF

' Show Window (SW_) contants.
Private Const SW_SHOWNORMAL    As Long = 1


'===============================================================================
' Public Function Declarations.
'===============================================================================


'===============================================================================
' Private  Function Declarations.
'===============================================================================

Private Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
    ( _
        ByVal hwndCaller As Long, _
        ByVal pszFile As String, _
        ByVal uCommand As Long, _
        ByVal dwData As Any _
    ) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    ( _
        ByVal hwnd As Long, _
        ByVal lpVerb As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Integer _
    ) As Long


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gTopicShow _
( _
    ByRef lngTopic As HelpTopic _
)
    Dim strHelpFile As String
    Dim astrTopicPath(6) As String
    
    astrTopicPath(HelpTopic.WSA) = "HTML/WSA.htm"
    astrTopicPath(HelpTopic.WSAHowTo) = "HTML/WSA/HowTo.htm"
    astrTopicPath(HelpTopic.WSACommandBar) = "HTML/WSA/CommandBar.htm"
    astrTopicPath(HelpTopic.WSACommandBarNavigator) = "HTML/WSA/CommandBar/Navigator.htm"
    astrTopicPath(HelpTopic.WSAKnownIssues) = "HTML/WSA/KnownIssues.htm"
    astrTopicPath(HelpTopic.WSACopyrightPermission) = "HTML/WSA/CopyrightPermission.htm"
    astrTopicPath(HelpTopic.WSAHistory) = "HTML/WSA/History.htm"
    
    strHelpFile = mstrFileNameGet(True)
    
    If (strHelpFile = "") Then
        Exit Sub
     End If
    
    Call modHelp.HtmlHelp( _
        0&, _
        strHelpFile, _
        modHelp.HH_DISPLAY_TOPIC, _
        astrTopicPath(lngTopic))
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gHomepageVisit _
( _
)
    ShellExecute _
        0&, _
        VBA.vbNullString, _
        modProject.GstrHomepage, _
        VBA.vbNullString, _
        VBA.vbNullString, _
        SW_SHOWNORMAL
End Sub


'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gAuthorEmail _
( _
)
    ShellExecute _
        0&, _
        VBA.vbNullString, _
        modProject.GstrEmail & "?subject=" & modProject.GstrName & "%20" & modProject.GstrVersion & ":%20", _
        VBA.vbNullString, _
        VBA.vbNullString, _
        SW_SHOWNORMAL
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mstrFileNameGet _
( _
    ByRef blnNotFoundMessageShow As Boolean _
) As String
    Dim strHelpFileName As String
    Dim blnHelpFileFound As Boolean
    Dim objFileSystem As Object
    Dim lngIndex As Long
    
    Set objFileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    
    blnHelpFileFound = False
    strHelpFileName = ""
    
    If (blnHelpFileFound = False) Then
        For lngIndex = 1 To Application.AddIns.Count
            If (Application.AddIns(lngIndex).Name = modProject.GstrName) Then
                blnHelpFileFound = True
                strHelpFileName = Application.AddIns(lngIndex).Path
                strHelpFileName = objFileSystem.BuildPath(strHelpFileName, modProject.GstrName & ".chm")
                Set objFileSystem = VBA.CreateObject("Scripting.FileSystemObject")
                If (objFileSystem.FileExists(strHelpFileName) = False) Then
                    blnHelpFileFound = False
                    strHelpFileName = ""
                 End If
            End If
        Next
    End If
    
    If (blnHelpFileFound = False) Then
        For lngIndex = 1 To Application.Presentations.Count
            If (VBA.LCase(Application.Presentations(lngIndex).Name) = VBA.LCase(modProject.GstrName & ".ppt")) Then
                blnHelpFileFound = True
                strHelpFileName = Application.Presentations(lngIndex).Path
                strHelpFileName = VBA.Left(strHelpFileName, InStrRev(strHelpFileName, "/"))
                strHelpFileName = objFileSystem.BuildPath(strHelpFileName, modProject.GstrName & ".chm")
                If (objFileSystem.FileExists(strHelpFileName) = False) Then
                    blnHelpFileFound = False
                    strHelpFileName = ""
                 End If
            End If
        Next
    End If
    
    If (blnHelpFileFound = False) Then
        If (blnNotFoundMessageShow = True) Then
            VBA.MsgBox _
                buttons:= _
                    VBA.vbExclamation, _
                Title:= _
                    modProject.GstrNamePretty, _
                Prompt:= _
                    "The Worship Service Assistant help file could not be " & _
                    "found. Be sure that both the add-in and help files are " & _
                    "in the same folder and that both files have the name " & _
                    "'" & modProject.GstrName & "' (excluding the file extension)."
        End If
     End If
    
    mstrFileNameGet = strHelpFileName
End Function
