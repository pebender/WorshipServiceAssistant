Attribute VB_Name = "modWin32AdvAPI32"
'===============================================================================
' Name:
'   WorshipServiceAssistant.modWin32AdvAPI32
'
' Description:
'   Interfaces to the Windows 32-bit (Win32) Advanced API (AdvAPI32) API.
'   This module only contains the needed parts of the API, not the complete
'   API.
'
' Author:
'   Paul Bender <pbender@alumni.ucsd.edu>
'
' Copyright:
'   Copyright (c) 2002 Paul Bender
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
'   1.02.0000:
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
Public Const ERROR_SUCCESS                As Long = 0&

Public Const HKEY_CLASSES_ROOT            As Long = &H80000000
Public Const HKEY_CURRENT_USER            As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE           As Long = &H80000002
Public Const HKEY_USERS                   As Long = &H80000003
Public Const HKEY_PERFORMANCE_DATA        As Long = &H80000004

Public Const SYNCHRONIZE                  As Long = &H100000
Public Const STANDARD_RIGHTS_ALL          As Long = &H1F0000
Public Const KEY_QUERY_VALUE              As Long = &H1
Public Const KEY_SET_VALUE                As Long = &H2
Public Const KEY_CREATE_LINK              As Long = &H20
Public Const KEY_CREATE_SUB_KEY           As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS       As Long = &H8
Public Const KEY_EVENT                    As Long = &H1
Public Const KEY_NOTIFY                   As Long = &H10
Public Const READ_CONTROL                 As Long = &H20000
Public Const STANDARD_RIGHTS_READ         As Long = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE        As Long = (READ_CONTROL)
Public Const KEY_ALL_ACCESS               As Long = ((STANDARD_RIGHTS_ALL Or _
                                                      KEY_QUERY_VALUE Or _
                                                      KEY_SET_VALUE Or _
                                                      KEY_CREATE_SUB_KEY Or _
                                                      KEY_ENUMERATE_SUB_KEYS Or _
                                                      KEY_NOTIFY Or _
                                                      KEY_CREATE_LINK) And _
                                                     (Not SYNCHRONIZE))
Public Const KEY_READ                     As Long = ((STANDARD_RIGHTS_READ Or _
                                                      KEY_QUERY_VALUE Or _
                                                      KEY_ENUMERATE_SUB_KEYS Or _
                                                      KEY_NOTIFY) And _
                                                     (Not SYNCHRONIZE))
Public Const KEY_EXECUTE                  As Long = (KEY_READ)
Public Const KEY_WRITE                    As Long = ((STANDARD_RIGHTS_WRITE Or _
                                                      KEY_SET_VALUE Or _
                                                      KEY_CREATE_SUB_KEY) And _
                                                     (Not SYNCHRONIZE))
Public Const REG_BINARY                   As Long = 3
Public Const REG_CREATED_NEW_KEY          As Long = &H1
Public Const REG_DWORD                    As Long = 4
Public Const REG_DWORD_BIG_ENDIAN         As Long = 5
Public Const REG_DWORD_LITTLE_ENDIAN      As Long = 4
Public Const REG_EXPAND_SZ                As Long = 2
Public Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9
Public Const REG_LINK                     As Long = 6
Public Const REG_MULTI_SZ                 As Long = 7
Public Const REG_NONE                     As Long = 0
Public Const REG_SZ                       As Long = 1
Public Const REG_NOTIFY_CHANGE_ATTRIBUTES As Long = &H2
Public Const REG_NOTIFY_CHANGE_LAST_SET   As Long = &H4
Public Const REG_NOTIFY_CHANGE_NAME       As Long = &H1
Public Const REG_NOTIFY_CHANGE_SECURITY   As Long = &H8
Public Const REG_OPTION_BACKUP_RESTORE    As Long = 4
Public Const REG_OPTION_CREATE_LINK       As Long = 2
Public Const REG_OPTION_NON_VOLATILE      As Long = 0
Public Const REG_OPTION_RESERVED          As Long = 0
Public Const REG_OPTION_VOLATILE          As Long = 1
Public Const REG_LEGAL_CHANGE_FILTER      As Long = (REG_NOTIFY_CHANGE_NAME Or _
                                                     REG_NOTIFY_CHANGE_ATTRIBUTES Or _
                                                     REG_NOTIFY_CHANGE_LAST_SET Or _
                                                     REG_NOTIFY_CHANGE_SECURITY)
Public Const REG_LEGAL_OPTION             As Long = (REG_OPTION_RESERVED Or _
                                                     REG_OPTION_NON_VOLATILE Or _
                                                     REG_OPTION_VOLATILE Or _
                                                     REG_OPTION_CREATE_LINK Or _
                                                     REG_OPTION_BACKUP_RESTORE)
'===============================================================================
' Function declarations.
'===============================================================================
Public Declare Function RegCloseKey Lib "advapi32.dll" _
    ( _
        ByVal HKey As Long _
    ) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    ( _
        ByVal HKey As Long, _
        ByVal lpSubKey As String, _
        ByRef phkResult As Long _
    ) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    ( _
        ByVal HKey As Long, _
        ByVal lpSubKey As String _
    ) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    ( _
        ByVal HKey As Long, _
        ByVal lpValueName As String _
    ) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    ( _
        ByVal HKey As Long, _
        ByVal lpSubKey As String, _
        ByVal uloptions As Long, _
        ByVal samdesired As Long, _
        ByRef phkResult As Long _
    ) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    ( _
        ByVal HKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        ByRef lpData As Any, _
        ByRef lpcbData As Long _
    ) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    ( _
        ByVal HKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByRef lpData As Any, _
        ByVal cbData As Long _
    ) As Long
