VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.Help
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

Public Enum WsaHelpTopic
    wsaHelpTopicApplication = 0
    wsaHelpTopicHowTo = 1
    wsaHelpTopicToolbar = 2
    wsaHelpTopicToolbarNavigator = 3
    wsaHelpTopicKnownIssues = 4
    wsaHelpTopicCopyrightPermission = 5
    wsaHelpTopicHistory = 6
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
Private mblnInvariant    As Boolean
Private mudtApplication  As WorshipServiceAssistant.Application
Private mudtParent       As WorshipServiceAssistant.Application
Private mastrTopicPath() As String


'===============================================================================
' Public Properties and Methods.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Initialize _
( _
    ByRef udtApplication As WorshipServiceAssistant.Application, _
    ByRef udtParent As WorshipServiceAssistant.Application _
)
    mblnInvariant = False
    
    Set mudtApplication = udtApplication
    Set mudtParent = udtParent
    
    mTopicPathInitialize
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Terminate _
( _
)
    Set mudtParent = Nothing
    Set mudtApplication = Nothing
    
    mblnInvariant = False
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Invariant _
( _
) As Boolean
    Invariant = mblnInvariant
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Let Invariant _
( _
    ByRef blnInvariant As Boolean _
)
    If (Not blnInvariant = mblnInvariant) Then
        If (blnInvariant = True) Then
            mblnInvariant = blnInvariant
        End If
        If (blnInvariant = False) Then
            mblnInvariant = blnInvariant
        End If
    End If
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Application _
( _
) As WorshipServiceAssistant.Application
    Set Application = mudtApplication
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Parent _
( _
) As WorshipServiceAssistant.Application
    Set Parent = mudtParent
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub ShowTopic _
( _
    ByRef lngTopic As WorshipServiceAssistant.WsaHelpTopic _
)
    Dim strHelpFile As String
    
    strHelpFile = mstrFileNameGet(True)
    
    If (strHelpFile = "") Then
        Exit Sub
     End If
    
    Call HtmlHelp( _
        0&, _
        strHelpFile, _
        HH_DISPLAY_TOPIC, _
        mastrTopicPath(lngTopic))
End Sub

Public Sub ShowAbout _
( _
)
    Dim strPrompt As String
    
    strPrompt = _
        WorshipServiceAssistant.wsaApplicationNamePretty & " (version " & WorshipServiceAssistant.wsaApplicationVersion & ")" & _
        VBA.vbNewLine & VBA.vbNewLine & _
        WorshipServiceAssistant.wsaApplicationCopyright & _
        VBA.vbNewLine & VBA.vbNewLine & _
        "Worship Service Assistant is a Microsoft PowerPoint add-in " & _
        "that is designed to make PowerPoint a more useful tool in a " & _
        "worship service. " & _
        VBA.vbNewLine & VBA.vbNewLine & _
        WorshipServiceAssistant.wsaApplicationNamePretty & " comes with ABSOLUTELY NO WARRANTY; " & _
        "for details view the 'Help on Copyright and Permission' in the 'Help' menu. " & _
        "This is free software, " & _
        "and you are welcome to redistribute it under certain conditions; " & _
        "for details view the 'Help on Copyright and Permission' in the 'Help' menu. "

    VBA.MsgBox _
        Buttons:= _
            VBA.vbOKOnly + VBA.vbMsgBoxHelpButton, _
        Title:= _
            "About " & WorshipServiceAssistant.wsaApplicationNamePretty, _
        Prompt:= _
            strPrompt
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub VisitHomepage _
( _
)
    ShellExecute _
        0&, _
        VBA.vbNullString, _
        WorshipServiceAssistant.wsaApplicationHomepage, _
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
Public Sub EmailAuthor _
( _
)
    ShellExecute _
        0&, _
        VBA.vbNullString, _
        WorshipServiceAssistant.wsaApplicationEmail & "?subject=" & WorshipServiceAssistant.wsaApplicationName & "%20" & WorshipServiceAssistant.wsaApplicationVersion & ":%20", _
        VBA.vbNullString, _
        VBA.vbNullString, _
        SW_SHOWNORMAL
End Sub


'===============================================================================
' Private Properties and Methods.
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
    
    strHelpFileName = ""
    
    blnHelpFileFound = False
    
    If (Not blnHelpFileFound) Then
        With mudtApplication.Doppelganger.AddIns
            For lngIndex = 1 To .Count Step 1
                If (.Item(lngIndex).Name = WorshipServiceAssistant.wsaApplicationName) Then
                    blnHelpFileFound = True
                    strHelpFileName = .Item(lngIndex).Path
                    strHelpFileName = objFileSystem.BuildPath(strHelpFileName, WorshipServiceAssistant.wsaApplicationName & ".chm")
                    Set objFileSystem = VBA.CreateObject("Scripting.FileSystemObject")
                    If (Not objFileSystem.FileExists(strHelpFileName)) Then
                        blnHelpFileFound = False
                        strHelpFileName = ""
                     End If
                End If
            Next
        End With
    End If
    
    If (Not blnHelpFileFound) Then
        With mudtApplication.Doppelganger.AddIns
            For lngIndex = 1 To .Count Step 1
                If (VBA.LCase(.Item(lngIndex).Name) = VBA.LCase(WorshipServiceAssistant.wsaApplicationName & ".ppt")) Then
                    blnHelpFileFound = True
                    strHelpFileName = .Item(lngIndex).Path
                    strHelpFileName = VBA.Left(strHelpFileName, InStrRev(strHelpFileName, "/"))
                    strHelpFileName = objFileSystem.BuildPath(strHelpFileName, WorshipServiceAssistant.wsaApplicationName & ".chm")
                    If (Not objFileSystem.FileExists(strHelpFileName)) Then
                        blnHelpFileFound = False
                        strHelpFileName = ""
                     End If
                End If
            Next
        End With
    End If
    
    If (Not blnHelpFileFound) Then
        If (blnNotFoundMessageShow) Then
            VBA.MsgBox _
                Buttons:= _
                    VBA.vbExclamation, _
                Title:= _
                    WorshipServiceAssistant.wsaApplicationNamePretty, _
                Prompt:= _
                    "The Worship Service Assistant help file could not be " & _
                    "found. Be sure that both the add-in and help files are " & _
                    "in the same folder and that both files have the name " & _
                    "'" & WorshipServiceAssistant.wsaApplicationName & "' (excluding the file extension)."
        End If
     End If
    
    mstrFileNameGet = strHelpFileName
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mTopicPathInitialize _
( _
)
    ReDim mastrTopicPath(6) As String
    mastrTopicPath(wsaHelpTopicApplication) = "HTML/WSA.htm"
    mastrTopicPath(wsaHelpTopicHowTo) = "HTML/WSA/HowTo.htm"
    mastrTopicPath(wsaHelpTopicToolbar) = "HTML/WSA/Toolbar.htm"
    mastrTopicPath(wsaHelpTopicToolbarNavigator) = "HTML/WSA/Toolbar/Navigator.htm"
    mastrTopicPath(wsaHelpTopicKnownIssues) = "HTML/WSA/KnownIssues.htm"
    mastrTopicPath(wsaHelpTopicCopyrightPermission) = "HTML/WSA/CopyrightPermission.htm"
    mastrTopicPath(wsaHelpTopicHistory) = "HTML/WSA/History.htm"
End Sub


'===============================================================================
' Private Event Handlers.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Class_Initialize _
( _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Class_Terminate _
( _
)
End Sub
