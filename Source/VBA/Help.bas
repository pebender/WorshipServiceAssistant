Attribute VB_Name = "Help"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Help
'
' Description:
'   This module contains constants associated with the help system.
'
' Author:
'   Paul Bender <pebender@san.rr.com>
'
' Copyright:
'   Copyright (C) 2000, 2001 Paul Bender
'
'   This program is free software; you can redistribute it and/or
'   modify it under the terms of the GNU General Public License
'   as published by the Free Software Foundation; version 2 of the License.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program; if not, write to the Free Software
'   Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'
' Change History:
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

'-------------------------------------------------------------------------------
' HTML Help control API constants.
'-------------------------------------------------------------------------------
Public Const HH_DISPLAY_TOPIC As Long = &H0
Public Const HH_HELP_CONTEXT  As Long = &HF

'-------------------------------------------------------------------------------
' Topic identifiers.
'-------------------------------------------------------------------------------
Public Const IDH_Topic_WSA                        As Long = 10000
Public Const IDH_Topic_WSAHowTo                   As Long = 11000
Public Const IDH_Topic_WSACommandBar              As Long = 12000
Public Const IDH_Topic_WSACommandBarNavigator     As Long = 12100
Public Const IDH_Topic_WSACopyright               As Long = 10100
Public Const IDH_Topic_WSALicense                 As Long = 10200
Public Const IDH_Topic_WSAHistory                 As Long = 10300

Public Const IDH_TopicPath_WSA                    As String = "HTML/WSA.htm"
Public Const IDH_TopicPath_WSAHowTo               As String = "HTML/WSA/HowTo.htm"
Public Const IDH_TopicPath_WSACommandBar          As String = "HTML/WSA/CommandBar.htm"
Public Const IDH_TopicPath_WSACommandBarNavigator As String = "HTML/WSA/CommandBar/Navigator.htm"
Public Const IDH_TopicPath_WSACopyright           As String = "HTML/WSA/Copyright.htm"
Public Const IDH_TopicPath_WSALicense             As String = "HTML/WSA/License.htm"
Public Const IDH_TopicPath_WSAHistory             As String = "HTML/WSA/History.htm"


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
'   HTML Help control API function declarations.
'-------------------------------------------------------------------------------
Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
    (ByVal hwndCaller As Long, _
     ByVal pszFile As String, _
     ByVal uCommand As Long, _
     ByVal dwData As Any) As Long

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Function Help_GetHelpFileName(ByVal ShowNotFoundMessage As Boolean) As String
    Dim HelpFileName As String
    Dim HelpFileFound As Boolean
    Dim FileSystem As Object
    Dim Index As Long
    
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    
    HelpFileFound = False
    HelpFileName = ""
    
    If (HelpFileFound = False) Then
        For Index = 1 To Application.AddIns.Count
            If (Application.AddIns(Index).Name = ProjectName) Then
                HelpFileFound = True
                HelpFileName = Application.AddIns(Index).Path
                HelpFileName = FileSystem.BuildPath(HelpFileName, ProjectName & ".chm")
                Set FileSystem = CreateObject("Scripting.FileSystemObject")
                If (FileSystem.FileExists(HelpFileName) = False) Then
                    HelpFileFound = False
                    HelpFileName = ""
                 End If
            End If
        Next
    End If
    
    If (HelpFileFound = False) Then
        For Index = 1 To Application.Presentations.Count
            If (LCase(Application.Presentations(Index).Name) = LCase(ProjectName & ".ppt")) Then
                HelpFileFound = True
                HelpFileName = Application.Presentations(Index).Path
                HelpFileName = Left(HelpFileName, InStrRev(HelpFileName, "/"))
                HelpFileName = FileSystem.BuildPath(HelpFileName, ProjectName & ".chm")
                If (FileSystem.FileExists(HelpFileName) = False) Then
                    HelpFileFound = False
                    HelpFileName = ""
                 End If
            End If
        Next
    End If
    
    If (HelpFileFound = False) Then
        If (ShowNotFoundMessage = True) Then
            MsgBox _
                buttons:= _
                    vbExclamation, _
                Title:= _
                    ProjectNamePretty, _
                Prompt:= _
                    "The Worship Service Assistant help file could not be " & _
                    "found. Be sure that both the add-in and help files are " & _
                    "in the same folder and that both files have the name " & _
                    "'" & ProjectName & "' (excluding the file extension)."
        End If
     End If
    
    Help_GetHelpFileName = HelpFileName
End Function


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
