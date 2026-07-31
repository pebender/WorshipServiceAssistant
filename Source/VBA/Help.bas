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
Public Const IDH_Topic_WSA                                                  As Long = 10000
Public Const IDH_Topic_WSAHowTo                                             As Long = 11000
Public Const IDH_Topic_WSAHowToEquipment                                    As Long = 11010
Public Const IDH_Topic_WSAHowToSongCreate                                   As Long = 11020
Public Const IDH_Topic_WSAHowToSongCategorize                               As Long = 11030
Public Const IDH_Topic_WSAHowToSongSort                                     As Long = 11040
Public Const IDH_Topic_WSAHowToSongIndex                                    As Long = 11050
Public Const IDH_Topic_WSAHowToSlideDisplay                                 As Long = 11060
Public Const IDH_Topic_WSAHowToManualPresentationDisplay                    As Long = 11070
Public Const IDH_Topic_WSAHowToTimedPresentationDisplay                     As Long = 11080
Public Const IDH_Topic_WSACommandBar                                        As Long = 12000
Public Const IDH_Topic_WSACommandBarNavigator                               As Long = 12100
Public Const IDH_Topic_WSACommandBarNavigatorSlideShow                      As Long = 12110
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowLoad                  As Long = 12111
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowHide                  As Long = 12112
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowRun                   As Long = 12113
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowPause                 As Long = 12114
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowNextEffect            As Long = 12115
Public Const IDH_Topic_WSACommandBarNavigatorSlideShowPrevEffect            As Long = 12116
Public Const IDH_Topic_WSACommandBarNavigatorGeneral                        As Long = 12120
Public Const IDH_Topic_WSACommandBarNavigatorGeneralExit                    As Long = 12121
Public Const IDH_Topic_WSACommandBarNavigatorGeneralHelp                    As Long = 12122
Public Const IDH_Topic_WSACommandBarNavigatorPresentationSelection          As Long = 12130
Public Const IDH_Topic_WSACommandBarNavigatorPresentationSelectionTitle     As Long = 12131
Public Const IDH_Topic_WSACommandBarNavigatorPresentationSelectionPrev      As Long = 12132
Public Const IDH_Topic_WSACommandBarNavigatorPresentationSelectionNext      As Long = 12133
Public Const IDH_Topic_WSACommandBarNavigatorSlideSelection                 As Long = 12140
Public Const IDH_Topic_WSACommandBarNavigatorSlideSelectionNumber           As Long = 12141
Public Const IDH_Topic_WSACommandBarNavigatorSlideSelectionTitle            As Long = 12142
Public Const IDH_Topic_WSACommandBarNavigatorSlideSelectionClear            As Long = 12143
Public Const IDH_Topic_WSACommandBarNavigatorSlideSelectionList             As Long = 12144
Public Const IDH_Topic_WSACommandBarLoad                                    As Long = 12200
Public Const IDH_Topic_WSACommandBarHide                                    As Long = 12300
Public Const IDH_Topic_WSACommandBarSongEdit                                As Long = 12400
Public Const IDH_Topic_WSACommandBarSongEditCategory                        As Long = 12410
Public Const IDH_Topic_WSACommandBarSongEditSort                            As Long = 12420
Public Const IDH_Topic_WSACommandBarSongEditIndex                           As Long = 12430
Public Const IDH_Topic_WSACommandBarCategory                                As Long = 12500
Public Const IDH_Topic_WSACommandBarDebug                                   As Long = 12600
Public Const IDH_Topic_WSACommandBarDebugSSWDisplay                         As Long = 12610
Public Const IDH_Topic_WSACommandBarDebugSSWSize                            As Long = 12620
Public Const IDH_Topic_WSACommandBarHelp                                    As Long = 12700
Public Const IDH_Topic_WSACommandBarHelpHelp                                As Long = 12710
Public Const IDH_Topic_WSACommandBarHelpHelpHowTo                           As Long = 12720
Public Const IDH_Topic_WSACommandBarHelpHelpCommandBar                      As Long = 12730
Public Const IDH_Topic_WSACommandBarHelpHelpCopyright                       As Long = 12740
Public Const IDH_Topic_WSACommandBarHelpHelpLicense                         As Long = 12750
Public Const IDH_Topic_WSACommandBarHelpVisitHomepage                       As Long = 12760
Public Const IDH_Topic_WSACommandBarHelpEmailAuthor                         As Long = 12770
Public Const IDH_Topic_WSACommandBarHelpDebug                               As Long = 12780
Public Const IDH_Topic_WSACommandBarHelpAbout                               As Long = 12790
Public Const IDH_Topic_WSACopyright                                         As Long = 10100
Public Const IDH_Topic_WSALicense                                           As Long = 10200
Public Const IDH_Topic_WSAHistory                                           As Long = 10300

Public Const IDH_TopicPath_WSA                                              As String = "HTML/WSA.htm"
Public Const IDH_TopicPath_WSAHowTo                                         As String = "HTML/WSA/HowTo.htm"
Public Const IDH_TopicPath_WSAHowToEquipment                                As String = "HTML/WSA/HowTo/EquipmentSetup.htm"
Public Const IDH_TopicPath_WSAHowToSongCreate                               As String = "HTML/WSA/HowTo/SongCreate.htm"
Public Const IDH_TopicPath_WSAHowToSongCategorize                           As String = "HTML/WSA/HowTo/SongCategorize.htm"
Public Const IDH_TopicPath_WSAHowToSongSort                                 As String = "HTML/WSA/HowTo/SongSort.htm"
Public Const IDH_TopicPath_WSAHowToSongIndex                                As String = "HTML/WSA/HowTo/SongIndex.htm"
Public Const IDH_TopicPath_WSAHowToSlideDisplay                             As String = "HTML/WSA/HowTo/SlideDisplay.htm"
Public Const IDH_TopicPath_WSAHowToManualPresentationDisplay                As String = "HTML/WSA/HowTo/ManualPresentationDisplay.htm"
Public Const IDH_TopicPath_WSAHowToTimedPresentationDisplay                 As String = "HTML/WSA/HowTo/TimedPresentationDisplay.htm"
Public Const IDH_TopicPath_WSACommandBar                                    As String = "HTML/WSA/CommandBar.htm"
Public Const IDH_TopicPath_WSACommandBarNavigator                           As String = "HTML/WSA/CommandBar/Navigator.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideShow                  As String = "HTML/WSA/CommandBar/Navigator/SlideShow.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideShowLoad              As String = "HTML/WSA/CommandBar/Navigator/SlideShow/Load.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideShowHide              As String = "HTML/WSA/CommandBar/Navigator/SlideShow/Hide.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideShowRun               As String = "HTML/WSA/CommandBar/Navigator/SlideShow/Run.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideShowPause             As String = "HTML/WSA/CommandBar/Navigator/SlideShow/Pause.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideShowNextEffect        As String = "HTML/WSA/CommandBar/Navigator/SlideShow/NextEffect.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideShowPrevEffect        As String = "HTML/WSA/CommandBar/Navigator/SlideShow/PrevEffect.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorGeneral                    As String = "HTML/WSA/CommandBar/Navigator/General.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorGeneralExit                As String = "HTML/WSA/CommandBar/Navigator/General/Exit.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorGeneralHelp                As String = "HTML/WSA/CommandBar/Navigator/General/Help.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorPresentationSelection      As String = "HTML/WSA/CommandBar/Navigator/PresentationSelection.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorPresentationSelectionTitle As String = "HTML/WSA/CommandBar/Navigator/PresentationSelection/Title.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorPresentationSelectionPrev  As String = "HTML/WSA/CommandBar/Navigator/PresentationSelection/Prev.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorPresentationSelectionNext  As String = "HTML/WSA/CommandBar/Navigator/PresentationSelection/Next.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideSelection             As String = "HTML/WSA/CommandBar/Navigator/SlideSelection.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideSelectionNumber       As String = "HTML/WSA/CommandBar/Navigator/SlideSelection/Number.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideSelectionTitle        As String = "HTML/WSA/CommandBar/Navigator/SlideSelection/Title.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideSelectionClear        As String = "HTML/WSA/CommandBar/Navigator/SlideSelection/Clear.htm"
Public Const IDH_TopicPath_WSACommandBarNavigatorSlideSelectionList         As String = "HTML/WSA/CommandBar/Navigator/SlideSelection/List.htm"
Public Const IDH_TopicPath_WSACommandBarLoad                                As String = "HTML/WSA/CommandBar/Load.htm"
Public Const IDH_TopicPath_WSACommandBarHide                                As String = "HTML/WSA/CommandBar/Hide.htm"
Public Const IDH_TopicPath_WSACommandBarSongEdit                            As String = "HTML/WSA/CommandBar/SongEdit.htm"
Public Const IDH_TopicPath_WSACommandBarSongEditCategory                    As String = "HTML/WSA/CommandBar/SongEdit/Category.htm"
Public Const IDH_TopicPath_WSACommandBarSongEditSort                        As String = "HTML/WSA/CommandBar/SongEdit/Sort.htm"
Public Const IDH_TopicPath_WSACommandBarSongEditIndex                       As String = "HTML/WSA/CommandBar/SongEdit/Index.htm"
Public Const IDH_TopicPath_WSACommandBarCategory                            As String = "HTML/WSA/CommandBar/Category.htm"
Public Const IDH_TopicPath_WSACommandBarDebug                               As String = "HTML/WSA/CommandBar/Debug.htm"
Public Const IDH_TopicPath_WSACommandBarDebugSSWDisplay                     As String = "HTML/WSA/CommandBar/Debug/SSWDisplay.htm"
Public Const IDH_TopicPath_WSACommandBarDebugSSWSize                        As String = "HTML/WSA/CommandBar/Debug/SSWSize.htm"
Public Const IDH_TopicPath_WSACommandBarHelp                                As String = "HTML/WSA/CommandBar/Help.htm"
Public Const IDH_TopicPath_WSACommandBarHelpHelp                            As String = "HTML/WSA/CommandBar/Help/Help.htm"
Public Const IDH_TopicPath_WSACommandBarHelpHelpHowTo                       As String = "HTML/WSA/CommandBar/Help/HelpHowTo.htm"
Public Const IDH_TopicPath_WSACommandBarHelpHelpCommandBar                  As String = "HTML/WSA/CommandBar/Help/HelpCommandBar.htm"
Public Const IDH_TopicPath_WSACommandBarHelpHelpCopyright                   As String = "HTML/WSA/CommandBar/Help/HelpCopyright.htm"
Public Const IDH_TopicPath_WSACommandBarHelpHelpLicense                     As String = "HTML/WSA/CommandBar/Help/HelpLicense.htm"
Public Const IDH_TopicPath_WSACommandBarHelpVisitHomepage                   As String = "HTML/WSA/CommandBar/Help/VisitHomepage.htm"
Public Const IDH_TopicPath_WSACommandBarHelpEmailAuthor                     As String = "HTML/WSA/CommandBar/Help/EmailAuthor.htm"
Public Const IDH_TopicPath_WSACommandBarHelpDebug                           As String = "HTML/WSA/CommandBar/Help/Debug.htm"
Public Const IDH_TopicPath_WSACommandBarHelpAbout                           As String = "HTML/WSA/CommandBar/Help/About.htm"
Public Const IDH_TopicPath_WSACopyright                                     As String = "HTML/WSA/Copyright.htm"
Public Const IDH_TopicPath_WSALicense                                       As String = "HTML/WSA/License.htm"
Public Const IDH_TopicPath_WSAHistory                                       As String = "HTML/WSA/History.htm"


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
