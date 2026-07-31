Attribute VB_Name = "Main"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Main
'
' Description:
'   This is a PowerPoint 9.0 (aka 2000) add-in.  It is intended for use in
'   dual monitor PowerPoint presentations.
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
        Case "Windows (32-bit) 4.10"
            ValidOperatingSystem = True
        Case "Windows (32-bit) 5.00"
            ValidOperatingSystem = True
        Case "Windows (32-bit) 5.01"
            ValidOperatingSystem = True
        Case Else
            Dim msgPrompt As String
            Dim msgTitle As String
            Dim msgResponse As VbMsgBoxResult
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
        Case Else
            Dim msgPrompt As String
            Dim msgTitle As String
            Dim msgResponse As VbMsgBoxResult
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
