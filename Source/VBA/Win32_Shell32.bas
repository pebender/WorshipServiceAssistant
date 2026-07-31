Attribute VB_Name = "Win32_Shell32"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Win32_Shell32
'
' Description:
'   Interfaces to the Windows 32-bit (Win32) Shell (Shell32) API.
'   This module only contains the needed parts of the API, not the complete
'   API. For an overview of this API, visit
'   <http://msdn.microsoft.com/library/default.asp?URL=/library/psdk/buildapp/win32api_4fos.htm>
'
' Author:
'   Paul Bender <pebender@san.rr.com>
'
' Copyright:
'   Copyright (C) 2001 Paul Bender
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
' Constants.
'===============================================================================
'
' General constants.
'
Public Const MAX_PATH                      As Long = 260
'
' Special IDentifier Location (SIDL) constants.
'
Public Const CSIDL_DESKTOP                 As Long = &H0
Public Const CSIDL_INTERNET                As Long = &H1
Public Const CSIDL_PROGRAMS                As Long = &H2
Public Const CSIDL_CONTROLS                As Long = &H3
Public Const CSIDL_PRINTERS                As Long = &H4
Public Const CSIDL_PERSONAL                As Long = &H5
Public Const CSIDL_FAVORITES               As Long = &H6
Public Const CSIDL_STARTUP                 As Long = &H7
Public Const CSIDL_RECENT                  As Long = &H8
Public Const CSIDL_SENDTO                  As Long = &H9
Public Const CSIDL_BITBUCKET               As Long = &HA
Public Const CSIDL_STARTMENU               As Long = &HB
Public Const CSIDL_DESKTOPDIRECTORY        As Long = &H10
Public Const CSIDL_DRIVES                  As Long = &H11
Public Const CSIDL_NETWORK                 As Long = &H12
Public Const CSIDL_NETHOOD                 As Long = &H13
Public Const CSIDL_FONTS                   As Long = &H14
Public Const CSIDL_TEMPLATES               As Long = &H15
Public Const CSIDL_COMMON_STARTMENU        As Long = &H16
Public Const CSIDL_COMMON_PROGRAMS         As Long = &H17
Public Const CSIDL_COMMON_STARTUP          As Long = &H18
Public Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19
Public Const CSIDL_APPDATA                 As Long = &H1A
Public Const CSIDL_PRINTHOOD               As Long = &H1B
Public Const CSIDL_ALTSTARTUP              As Long = &H1D
Public Const CSIDL_COMMON_ALTSTARTUP       As Long = &H1E
Public Const CSIDL_COMMON_FAVORITES        As Long = &H1F
Public Const CSIDL_INTERNET_CACHE          As Long = &H20
Public Const CSIDL_COOKIES                 As Long = &H21
Public Const CSIDL_HISTORY                 As Long = &H22
Public Const CSIDL_COMMON_APPDATA          As Long = &H23
Public Const CSIDL_WINDOWS                 As Long = &H24
Public Const CSIDL_SYSTEM                  As Long = &H25
Public Const CSIDL_PROGRAM_FILES           As Long = &H26
Public Const CSIDL_MYPICTURES              As Long = &H27
Public Const CSIDL_PROFILE                 As Long = &H28
Public Const CSIDL_SYSTEMX86               As Long = &H29
Public Const CSIDL_PROGRAM_FILESX86        As Long = &H2A
Public Const CSIDL_PROGRAM_FILES_COMMON    As Long = &H2B
Public Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2C
Public Const CSIDL_COMMON_TEMPLATES        As Long = &H2D
Public Const CSIDL_COMMON_DOCUMENTS        As Long = &H2E
Public Const CSIDL_COMMON_ADMINTOOLS       As Long = &H2F
Public Const CSIDL_ADMINTOOLS              As Long = &H30
Public Const CSIDL_CONNECTIONS             As Long = &H31
'
' Show Window (SW_) contants.
'
Public Const SW_HIDE                       As Long = 0
Public Const SW_NORMAL                     As Long = 1
Public Const SW_SHOWNORMAL                 As Long = 1
Public Const SW_SHOWMINIMIZED              As Long = 2
Public Const SW_SHOWMAXIMIZED              As Long = 3
Public Const SW_MAXIMIZE                   As Long = 3
Public Const SW_SHOWNOACTIVATE             As Long = 4
Public Const SW_SHOW                       As Long = 5
Public Const SW_MINIMIZE                   As Long = 6
Public Const SW_SHOWMINNOACTIVE            As Long = 7
Public Const SW_SHOWNA                     As Long = 8
Public Const SW_RESTORE                    As Long = 9
Public Const SW_SHOWDEFAULT                As Long = 10

'===============================================================================
' Function declarations.
'===============================================================================
Public Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" _
    (ByVal hwndOwner As Long, _
     ByVal lpszPath As String, _
     ByVal nFolder As Long, _
     ByVal fCreate As Boolean) As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
     ByVal lpVerb As String, _
     ByVal lpFile As String, _
     ByVal lpParameters As String, _
     ByVal lpDirectory As String, _
     ByVal nShowCmd As Integer) As Long
