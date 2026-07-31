Attribute VB_Name = "Win32_Shell32"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Win32_Shell32
'
' Description:
'   Interfaces to the Windows 32-bit (Win32) Shell (Shell32) API.
'   This module only contains the needed parts of the API, not the complete
'   API.
'
' Author:
'   Paul Bender <pbender@alumni.ucsd.edu>
'
' Copyright:
'   Copyright (c) 2000,2001 Paul Bender
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
