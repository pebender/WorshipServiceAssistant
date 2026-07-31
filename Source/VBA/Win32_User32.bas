Attribute VB_Name = "Win32_User32"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Win32_User32
'
' Description:
'   Interfaces to the Windows 32-bit (Win32) User Interface (User32) API.
'   This module only contains the needed parts of the API, not the complete
'   API. For an overview of this API, visit
'   <http://msdn.microsoft.com/library/default.asp?URL=/library/psdk/buildapp/win32api_5rtx.htm>
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
' GetWindowLong (GWL_) contants.
'
Public Const GWL_WNDPROC                   As Long = (-4)
Public Const GWL_HINSTANCE                 As Long = (-6)
Public Const GWL_HWNDPARENT                As Long = (-8)
Public Const GWL_STYLE                     As Long = (-16)
Public Const GWL_EXSTYLE                   As Long = (-20)
Public Const GWL_USERDATA                  As Long = (-21)
Public Const GWL_ID                        As Long = (-12)
'
' Lock Set Foreground Window (LSFW) constants
'
Public Const LSFW_LOCK                     As Long = 1
Public Const LSFW_UNLOCK                   As Long = 2
'
' ReDraW (RDW_) constants.
'
Public Const RDW_INVALIDATE                As Long = &H1
Public Const RDW_INTERNALPAINT             As Long = &H2
Public Const RDW_ERASE                     As Long = &H4
Public Const RDW_VALIDATE                  As Long = &H8
Public Const RDW_NOINTERNALPAINT           As Long = &H10
Public Const RDW_NOERASE                   As Long = &H20
Public Const RDW_NOCHILDREN                As Long = &H40
Public Const RDW_ALLCHILDREN               As Long = &H80
Public Const RDW_UPDATENOW                 As Long = &H100
Public Const RDW_ERASENOW                  As Long = &H200
Public Const RDW_FRAME                     As Long = &H400
Public Const RDW_NOFRAME                   As Long = &H800
'
' System Metric (SM) constants
'
Public Const SM_CXSCREEN                As Integer = 0
Public Const SM_CYSCREEN                As Integer = 1
Public Const SM_CXVSCROLL               As Integer = 2
Public Const SM_CYHSCROLL               As Integer = 3
Public Const SM_CYCAPTION               As Integer = 4
Public Const SM_CXBORDER                As Integer = 5
Public Const SM_CYBORDER                As Integer = 6
Public Const SM_CXDLGFRAME              As Integer = 7
Public Const SM_CYDLGFRAME              As Integer = 8
Public Const SM_CYVTHUMB                As Integer = 9
Public Const SM_CXHTHUMB                As Integer = 10
Public Const SM_CXICON                  As Integer = 11
Public Const SM_CYICON                  As Integer = 12
Public Const SM_CXCURSOR                As Integer = 13
Public Const SM_CYCURSOR                As Integer = 14
Public Const SM_CYMENU                  As Integer = 15
Public Const SM_CXFULLSCREEN            As Integer = 16
Public Const SM_CYFULLSCREEN            As Integer = 17
Public Const SM_CYKANJIWINDOW           As Integer = 18
Public Const SM_MOUSEPRESENT            As Integer = 19
Public Const SM_CYVSCROLL               As Integer = 20
Public Const SM_CXHSCROLL               As Integer = 21
Public Const SM_DEBUG                   As Integer = 22
Public Const SM_SWAPBUTTON              As Integer = 23
Public Const SM_RESERVED1               As Integer = 24
Public Const SM_RESERVED2               As Integer = 25
Public Const SM_RESERVED3               As Integer = 26
Public Const SM_RESERVED4               As Integer = 27
Public Const SM_CXMIN                   As Integer = 28
Public Const SM_CYMIN                   As Integer = 29
Public Const SM_CXSIZE                  As Integer = 30
Public Const SM_CYSIZE                  As Integer = 31
Public Const SM_CXFRAME                 As Integer = 32
Public Const SM_CYFRAME                 As Integer = 33
Public Const SM_CXMINTRACK              As Integer = 34
Public Const SM_CYMINTRACK              As Integer = 35
Public Const SM_CXDOUBLECLK             As Integer = 36
Public Const SM_CYDOUBLECLK             As Integer = 37
Public Const SM_CXICONSPACING           As Integer = 38
Public Const SM_CYICONSPACING           As Integer = 39
Public Const SM_MENUDROPALIGNMENT       As Integer = 40
Public Const SM_PENWINDOWS              As Integer = 41
Public Const SM_DBCSENABLED             As Integer = 42
Public Const SM_CMOUSEBUTTONS           As Integer = 43
Public Const SM_CXFIXEDFRAME            As Integer = SM_CXDLGFRAME
Public Const SM_CYFIXEDFRAME            As Integer = SM_CYDLGFRAME
Public Const SM_CXSIZEFRAME             As Integer = SM_CXFRAME
Public Const SM_CYSIZEFRAME             As Integer = SM_CYFRAME
Public Const SM_SECURE                  As Integer = 44
Public Const SM_CXEDGE                  As Integer = 45
Public Const SM_CYEDGE                  As Integer = 46
Public Const SM_CXMINSPACING            As Integer = 47
Public Const SM_CYMINSPACING            As Integer = 48
Public Const SM_CXSMICON                As Integer = 49
Public Const SM_CYSMICON                As Integer = 50
Public Const SM_CYSMCAPTION             As Integer = 51
Public Const SM_CXSMSIZE                As Integer = 52
Public Const SM_CYSMSIZE                As Integer = 53
Public Const SM_CXMENUSIZE              As Integer = 54
Public Const SM_CYMENUSIZE              As Integer = 55
Public Const SM_ARRANGE                 As Integer = 56
Public Const SM_CXMINIMIZED             As Integer = 57
Public Const SM_CYMINIMIZED             As Integer = 58
Public Const SM_CXMAXTRACK              As Integer = 59
Public Const SM_CYMAXTRACK              As Integer = 60
Public Const SM_CXMAXIMIZED             As Integer = 61
Public Const SM_CYMAXIMIZED             As Integer = 62
Public Const SM_NETWORK                 As Integer = 63
Public Const SM_CLEANBOOT               As Integer = 67
Public Const SM_CXDRAG                  As Integer = 68
Public Const SM_CYDRAG                  As Integer = 69
Public Const SM_SHOWSOUNDS              As Integer = 70
Public Const SM_CXMENUCHECK             As Integer = 71
Public Const SM_CYMENUCHECK             As Integer = 72
Public Const SM_SLOWMACHINE             As Integer = 73
Public Const SM_MIDEASTENABLED          As Integer = 74
Public Const SM_MOUSEWHEELPRESENT       As Integer = 75
Public Const SM_XVIRTUALSCREEN          As Integer = 76
Public Const SM_YVIRTUALSCREEN          As Integer = 77
Public Const SM_CXVIRTUALSCREEN         As Integer = 78
Public Const SM_CYVIRTUALSCREEN         As Integer = 79
Public Const SM_CMONITORS               As Integer = 80
Public Const SM_SAMEDISPLAYFORMAT       As Integer = 81
Public Const SM_IMMENABLED              As Integer = 82
Public Const SM_CMETRICS                As Integer = 83
Public Const SM_REMOTESESSION           As Integer = &H1000
'
' Show Window (SW_) contants.
'
Public Const SW_HIDE                       As Long = 0
Public Const SW_SHOWNORMAL                 As Long = 1
Public Const SW_NORMAL                     As Long = 1
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
Public Const SW_FORCEMINIMIZE              As Long = 11
Public Const SW_MAX                        As Long = 11
'
' Window State (WS_) constants.
'
Public Const WS_OVERLAPPED                 As Long = &H0
Public Const WS_POPUP                      As Long = &H80000000
Public Const WS_CHILD                      As Long = &H40000000
Public Const WS_MINIMIZE                   As Long = &H20000000
Public Const WS_VISIBLE                    As Long = &H10000000
Public Const WS_DISABLED                   As Long = &H8000000
Public Const WS_CLIPSIBLINGS               As Long = &H4000000
Public Const WS_CLIPCHILDREN               As Long = &H2000000
Public Const WS_MAXIMIZE                   As Long = &H1000000
Public Const WS_BORDER                     As Long = &H800000
Public Const WS_DLGFRAME                   As Long = &H400000
Public Const WS_CAPTION                    As Long = WS_BORDER Or WS_DLGFRAME
Public Const WS_VSCROLL                    As Long = &H200000
Public Const WS_HSCROLL                    As Long = &H100000
Public Const WS_SYSMENU                    As Long = &H80000
Public Const WS_THICKFRAME                 As Long = &H40000
Public Const WS_GROUP                      As Long = &H20000
Public Const WS_TABSTOP                    As Long = &H10000
Public Const WS_MINIMIZEBOX                As Long = &H20000
Public Const WS_MAXIMIZEBOX                As Long = &H10000
Public Const WS_TILED                      As Long = WS_OVERLAPPED
Public Const WS_ICONIC                     As Long = WS_MINIMIZE
Public Const WS_SIZEBOX                    As Long = WS_THICKFRAME
Public Const WS_OVERLAPPEDWINDOW           As Long = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
Public Const WS_TILEDWINDOW                As Long = WS_OVERLAPPEDWINDOW
Public Const WS_POPUPWINDOW                As Long = WS_POPUP Or WS_BORDER Or WS_SYSMENU
Public Const WS_CHILDWINDOW                As Long = WS_CHILD
'
' Extended Window State (WS_EX_) constants.
'
Public Const WS_EX_DLGMODALFRAME           As Long = &H1
Public Const WS_EX_NOPARENTNOTIFY          As Long = &H4
Public Const WS_EX_TOPMOST                 As Long = &H8
Public Const WS_EX_ACCEPTFILES             As Long = &H10
Public Const WS_EX_TRANSPARENT             As Long = &H20
Public Const WS_EX_MDICHILD                As Long = &H40
Public Const WS_EX_TOOLWINDOW              As Long = &H80
Public Const WS_EX_WINDOWEDGE              As Long = &H100
Public Const WS_EX_CLIENTEDGE              As Long = &H200
Public Const WS_EX_CONTEXTHELP             As Long = &H400
Public Const WS_EX_RIGHT                   As Long = &H1000
Public Const WS_EX_LEFT                    As Long = &H0
Public Const WS_EX_RTLREADING              As Long = &H2000
Public Const WS_EX_LTRREADING              As Long = &H0
Public Const WS_EX_LEFTSCROLLBAR           As Long = &H4000
Public Const WS_EX_RIGHTSCROLLBAR          As Long = &H0
Public Const WS_EX_CONTROLPARENT           As Long = &H10000
Public Const WS_EX_STATICEDGE              As Long = &H20000
Public Const WS_EX_APPWINDOW               As Long = &H40000
Public Const WS_EX_NOACTIVATE              As Long = &H8000000
Public Const WS_EX_OVERLAPPEDWINDOW        As Long = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
Public Const WS_EX_PALETTEWINDOW           As Long = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST

'===============================================================================
' Function declarations.
'===============================================================================
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) As Long
Public Declare Function GetSystemMetrics Lib "user32.dll" _
    (ByVal nIndex As Integer) As Integer
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long) As Long
Public Declare Function LockSetForegroundWindow Lib "user32.dll" _
    (ByVal uLockCode As Long) As Boolean
Public Declare Function LockWindowUpdate Lib "user32.dll" _
    (ByVal hwnd As Long) As Boolean
Public Declare Function RedrawWindow Lib "user32.dll" _
    (ByVal hwnd As Long, _
     ByVal lprcUpdate As Long, _
     ByVal hrgnUpdate As Long, _
     ByVal flags As Long) As Boolean
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
Public Declare Function ShowWindow Lib "user32.dll" _
    (ByVal hwnd As Long, _
     ByVal nCmdShow As Long) As Long
