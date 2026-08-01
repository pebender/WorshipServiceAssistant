Attribute VB_Name = "modWin32User32"
'===============================================================================
' Name:
'   WorshipServiceAssistant.modWin32User32
'
' Description:
'   Interfaces to the Windows 32-bit (Win32) User Interface (User32) API.
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
'   1.04.0002:
'     (1) Removed unused APIs.
'     (2) Added APIs for getting the monitor name.
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

' System Metric (SM) constants
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

' Display device constants
Public Const DISPLAY_DEVICE_ATTACHED_TO_DESKTOP = &H1
Public Const DISPLAY_DEVICE_MULTI_DRIVER = &H2
Public Const DISPLAY_DEVICE_PRIMARY_DEVICE = &H4
Public Const DISPLAY_DEVICE_MIRRORING_DRIVER = &H8
Public Const DISPLAY_DEVICE_VGA_COMPATIBLE = &H10
Public Const DISPLAY_DEVICE_REMOVABLE = &H20
Public Const DISPLAY_DEVICE_MODESPRUNED = &H8000000
Public Const DISPLAY_DEVICE_REMOTE = &H4000000
Public Const DISPLAY_DEVICE_DISCONNECT = &H2000000

Public Const DISPLAY_DEVICE_ACTIVE = &H1
Public Const DISPLAY_DEVICE_ATTACHED = &H2

'===============================================================================
' Type declarations.
'===============================================================================

' Display device types
Public Type DISPLAY_DEVICE
    cb As Long
    DeviceName As String * 32
    DeviceString As String * 128
    StateFlags As Long
    DeviceID As String * 128
    DeviceKey As String * 128
End Type


'===============================================================================
' Function declarations.
'===============================================================================
Public Declare Function GetSystemMetrics Lib "user32.dll" _
    ( _
        ByVal nIndex As Integer _
    ) As Integer
Public Declare Function EnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesA" _
    ( _
        ByVal lpDevice As String, _
        ByVal iDevNum As Long, _
        ByRef lpDisplayDevice As DISPLAY_DEVICE, _
        ByRef dwFlags As Long _
    ) As Long
