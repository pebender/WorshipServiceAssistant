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
' Topic identifiers.
'-------------------------------------------------------------------------------
Public Const IDH_Topic_WSA = 10000
Public Const IDH_Topic_WSAHowTO = 11000
Public Const IDH_Topic_WSAHowToEquipment = 11010
Public Const IDH_Topic_WSAHowToSongCreate = 11020
Public Const IDH_Topic_WSAHowToSongCategorize = 11030
Public Const IDH_Topic_WSAHowToSongSort = 11040
Public Const IDH_Topic_WSAHowToSongIndex = 11050
Public Const IDH_Topic_WSAHowToSlideDisplay = 11060
Public Const IDH_Topic_WSAHowToManualPresentationDisplay = 11070
Public Const IDH_Topic_WSAHowToTimedPresentationDisplay = 11080
Public Const IDH_Topic_WSACommandBar = 12000
Public Const IDH_Topic_WSACommandBarNavigator = 12100
Public Const IDH_Topic_WSACommandBarNavigatorSlideShow = 12110
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowLoad = 12111
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowHide = 12112
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowRun = 12113
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowPause = 12114
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowNextEffect = 12115
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowPrevEffect = 12116
Public Const IDH_Topic_WSACommandBarNavigatorGeneral = 12120
Public Const IDH_Topic_WSACommandBarNavigatorGeneralExit = 12121
Public Const IDH_Topic_WSACommandBarNavigatorGeneralHelp = 12122
Public Const IDH_Topic_WSACommandBarNavigatorPresentationSelection = 12130
Public Const IDH_Topic_WSACommandBarNavigatorPresentationSelectionTitle = 12131
Public Const IDH_Topic_WSACommandBarNavigatorPresentationSelectionPrev = 12132
Public Const IDH_Topic_WSACommandBarNavigatorPresentationSelectionNext = 12133
Public Const IDH_Topic_WSACommandBarNavigatorSlideSelection = 12140
Public Const IDH_Topic_WSACommandBarNavigatorSlideSelectionNumber = 12141
Public Const IDH_Topic_WSACommandBarNavigatorSlideSelectionTitle = 12142
Public Const IDH_Topic_WSACommandBarNavigatorSlideSelectionClear = 12143
Public Const IDH_Topic_WSACommandBarNavigatorSlideSelectionList = 12144
Public Const IDH_Topic_WSACommandBarLoad = 12200
Public Const IDH_Topic_WSACommandBarHide = 12300
Public Const IDH_Topic_WSACommandBarSongEdit = 12400
Public Const IDH_Topic_WSACommandBarSongEditCategory = 12410
Public Const IDH_Topic_WSACommandBarSongEditSort = 12420
Public Const IDH_Topic_WSACommandBarSongEditIndex = 12430
Public Const IDH_Topic_WSACommandBarDebug = 12500
Public Const IDH_Topic_WSACommandBarDebugSSWDisplay = 12510
Public Const IDH_Topic_WSACommandBarDebugSSWSize = 12520
Public Const IDH_Topic_WSACommandBarHelp = 12600
Public Const IDH_Topic_WSACommandBarHelpHelp = 12610
Public Const IDH_Topic_WSACommandBarHelpHelpHowTo = 12620
Public Const IDH_Topic_WSACommandBarHelpHelpCommandBar = 12630
Public Const IDH_Topic_WSACommandBarHelpHelpCopyright = 12640
Public Const IDH_Topic_WSACommandBarHelpHelpLicense = 12650
Public Const IDH_Topic_WSACommandBarHelpVisitHomepage = 12660
Public Const IDH_Topic_WSACommandBarHelpEmailAuthor = 12670
Public Const IDH_Topic_WSACommandBarHelpDebug = 12680
Public Const IDH_Topic_WSACommandBarHelpAbout = 12690
Public Const IDH_Topic_WSACopyright = 10100
Public Const IDH_Topic_WSALicense = 10200

'-------------------------------------------------------------------------------
' Context Sensitive Help identifiers.
'-------------------------------------------------------------------------------
Public Const IDH_CSHelp_WSA = 20000
Public Const IDH_CSHelp_WSAHowTO = 21000
Public Const IDH_CSHelp_WSAHowToEquipment = 21010
Public Const IDH_CSHelp_WSAHowToSongCreate = 21020
Public Const IDH_CSHelp_WSAHowToSongCategorize = 21030
Public Const IDH_CSHelp_WSAHowToSongSort = 21040
Public Const IDH_CSHelp_WSAHowToSongIndex = 21050
Public Const IDH_CSHelp_WSAHowToSlideDisplay = 21060
Public Const IDH_CSHelp_WSAHowToManualPresentationDisplay = 21070
Public Const IDH_CSHelp_WSAHowToTimedPresentationDisplay = 21080
Public Const IDH_CSHelp_WSACommandBar = 22000
Public Const IDH_CSHelp_WSACommandBarNavigator = 22100
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideShow = 22110
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideShowLoad = 22111
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideShowHide = 22112
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideShowRun = 22113
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideShowPause = 22114
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideShowNextEffect = 22115
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideShowPrevEffect = 22116
Public Const IDH_CSHelp_WSACommandBarNavigatorGeneral = 22120
Public Const IDH_CSHelp_WSACommandBarNavigatorGeneralExit = 22121
Public Const IDH_CSHelp_WSACommandBarNavigatorGeneralHelp = 22122
Public Const IDH_CSHelp_WSACommandBarNavigatorPresentationSelection = 22130
Public Const IDH_CSHelp_WSACommandBarNavigatorPresentationSelectionTitle = 22131
Public Const IDH_CSHelp_WSACommandBarNavigatorPresentationSelectionPrev = 22132
Public Const IDH_CSHelp_WSACommandBarNavigatorPresentationSelectionNext = 22133
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideSelection = 22140
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideSelectionNumber = 22141
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideSelectionTitle = 22142
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideSelectionClear = 22143
Public Const IDH_CSHelp_WSACommandBarNavigatorSlideSelectionList = 22144
Public Const IDH_CSHelp_WSACommandBarLoad = 22200
Public Const IDH_CSHelp_WSACommandBarHide = 22300
Public Const IDH_CSHelp_WSACommandBarSongEdit = 22400
Public Const IDH_CSHelp_WSACommandBarSongEditCategory = 22410
Public Const IDH_CSHelp_WSACommandBarSongEditSort = 22420
Public Const IDH_CSHelp_WSACommandBarSongEditIndex = 22430
Public Const IDH_CSHelp_WSACommandBarDebug = 22500
Public Const IDH_CSHelp_WSACommandBarDebugSSWDisplay = 22510
Public Const IDH_CSHelp_WSACommandBarDebugSSWSize = 22520
Public Const IDH_CSHelp_WSACommandBarHelp = 12600
Public Const IDH_CSHelp_WSACommandBarHelpHelp = 12610
Public Const IDH_CSHelp_WSACommandBarHelpHelpHowTo = 12620
Public Const IDH_CSHelp_WSACommandBarHelpHelpCommandBar = 12630
Public Const IDH_CSHelp_WSACommandBarHelpHelpCopyright = 12640
Public Const IDH_CSHelp_WSACommandBarHelpHelpLicense = 12650
Public Const IDH_CSHelp_WSACommandBarHelpVisitHomepage = 22660
Public Const IDH_CSHelp_WSACommandBarHelpEmailAuthor = 22670
Public Const IDH_CSHelp_WSACommandBarHelpDebug = 22680
Public Const IDH_CSHelp_WSACommandBarHelpAbout = 12690


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
                HelpFileName = Left(HelpFileName, InStrRev(HelpFileName, "\"))
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
