VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSAApplicationOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.WSAApplicationOptions
'
' Description:
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
'   1.04.0003:
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
' Private Constants.
'===============================================================================
' Registry constants
Private Const ERROR_SUCCESS                      As Long = 0&

Private Const HKEY_CURRENT_USER                  As Long = &H80000001

Private Const KEY_QUERY_VALUE                    As Long = &H1
Private Const KEY_SET_VALUE                      As Long = &H2

Private Const REG_DWORD                          As Long = 4
Private Const REG_SZ                             As Long = 1

' Display device constants
Private Const DISPLAY_DEVICE_ATTACHED_TO_DESKTOP As Long = &H1
Private Const DISPLAY_DEVICE_MIRRORING_DRIVER    As Long = &H8

Private Const DISPLAY_DEVICE_ACTIVE              As Long = &H1

' WSAApplicationOptions constants
Private Const mlngPpControlIdToolsOptions        As Long = 522


'===============================================================================
' Public Type Declarations.
'===============================================================================

'===============================================================================
' Private Type Declarations.
'===============================================================================

' Display device types
Private Type DISPLAY_DEVICE
    cb           As Long
    DeviceName   As String * 32
    DeviceString As String * 128
    StateFlags   As Long
    DeviceID     As String * 128
    DeviceKey    As String * 128
End Type

Private Type OptionIndirectItem
    Name        As String
    Tab         As Long
    ShortcutKey As String
End Type

'===============================================================================
' Public Function Declarations.
'===============================================================================


'===============================================================================
' Private Function Declarations.
'===============================================================================
Private Declare Function RegCloseKey Lib "advapi32.dll" _
    ( _
        ByVal HKey As Long _
    ) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    ( _
        ByVal HKey As Long, _
        ByVal lpSubKey As String, _
        ByVal uloptions As Long, _
        ByVal samdesired As Long, _
        ByRef phkResult As Long _
    ) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    ( _
        ByVal HKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        lpData As Any, _
        ByRef lpcbData As Long _
    ) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    ( _
        ByVal HKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        lpData As Any, _
        ByVal cbData As Long _
    ) As Long

Private Declare Function EnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesA" _
    ( _
        ByVal lpDevice As String, _
        ByVal iDevNum As Long, _
        ByRef lpDisplayDevice As DISPLAY_DEVICE, _
        ByRef dwFlags As Long _
    ) As Long
    

'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
' The mvntOptionIndirect property uses the array.
' The Class_Initialize event handler initializes the array.
Private maudtOptionIndirectList() As OptionIndirectItem


'===============================================================================
' Public Properties and Methods.
'===============================================================================

Public Property Get DisplayMonitor() As String
    DisplayMonitor = mvntOption("DisplayMonitor")
End Property

Public Property Let DisplayMonitor(ByVal strData As String)
    mvntOption("DisplayMonitor") = strData
End Property

Public Property Get SaveAutoRecoveryInfo() As Long
    SaveAutoRecoveryInfo = mvntOption("SaveAutoRecoveryInfo")
End Property

Public Property Let SaveAutoRecoveryInfo(ByVal lngData As Long)
    mvntOptionIndirect("SaveAutoRecoveryInfo") = lngData
End Property

Public Property Get SSEndOnBlankSlide() As Long
    SSEndOnBlankSlide = mvntOption("SSEndOnBlankSlide")
End Property

Public Property Let SSEndOnBlankSlide(ByVal lngData As Long)
    mvntOptionIndirect("SSEndOnBlankSlide") = lngData
End Property

Public Property Get SSMenuButton() As Long
    SSMenuButton = mvntOption("SSMenuButton")
End Property

Public Property Let SSMenuButton(ByVal lngData As Long)
    mvntOptionIndirect("SSMenuButton") = lngData
End Property

Public Property Get SSRightMouse() As Long
    SSRightMouse = mvntOption("SSRightMouse")
End Property

Public Property Let SSRightMouse(ByVal lngData As Long)
    mvntOptionIndirect("SSRightMouse") = lngData
End Property

Public Property Get UseMonMgr() As Long
    UseMonMgr = mvntOption("UseMonMgr")
End Property

Public Property Let UseMonMgr(ByVal lngData As Long)
    mvntOption("UseMonMgr") = lngData
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get DisplayMonitorList() As String()
    Dim udtDisplay As DISPLAY_DEVICE
    Dim udtMonitor As DISPLAY_DEVICE
    Dim lngDisplay As Long
    Dim lngMonitor As Long
    Dim strMonitorName As String
    Dim lngMonitorIndex As Long
    Dim astrDisplayMonitorList() As String
    
    udtDisplay.cb = Len(udtDisplay)
    udtMonitor.cb = Len(udtMonitor)
    
    ReDim astrDisplayMonitorList(0)
    
    ' Loop through the displays.
    lngDisplay = 0
    Do While (EnumDisplayDevices(VBA.vbNullString, lngDisplay, udtDisplay, 0) <> 0)
        ' Skip displays that are not attached to the desktop.
        If ((udtDisplay.StateFlags And DISPLAY_DEVICE_ATTACHED_TO_DESKTOP) <> 0) Then
            ' Skip displays that are not real.
            If ((udtDisplay.StateFlags And DISPLAY_DEVICE_MIRRORING_DRIVER) = 0) Then
                ' Loop through monitors.
                lngMonitor = 0
                Do While (EnumDisplayDevices(udtDisplay.DeviceName, lngMonitor, udtMonitor, 0) <> 0)
                    If ((udtMonitor.StateFlags And DISPLAY_DEVICE_ACTIVE) <> 0) Then
                        Exit Do
                    End If
                    lngMonitor = lngMonitor + 1
                Loop
                If (udtMonitor.DeviceName <> "") Then
                    astrDisplayMonitorList(UBound(astrDisplayMonitorList)) = mstrC2VBConvert(udtMonitor.DeviceName)
                Else
                    astrDisplayMonitorList(UBound(astrDisplayMonitorList)) = mstrC2VBConvert(udtDisplay.DeviceName)
                End If
                ReDim Preserve astrDisplayMonitorList(UBound(astrDisplayMonitorList) + 1)
            End If
        End If
        lngDisplay = lngDisplay + 1
    Loop
    If (UBound(astrDisplayMonitorList) >= 1) Then
        ReDim Preserve astrDisplayMonitorList(UBound(astrDisplayMonitorList) - 1)
    End If

    DisplayMonitorList = astrDisplayMonitorList
End Property

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
Private Property Get mvntOption _
( _
    ByRef strOptionName As String _
) As Variant
    Dim lngResult As Long
    Dim lngKeyHandle As Long
    Dim strOptionPath As String
    Dim lngOptionType As Long
    Dim lngOptionData As Long
    Dim strOptionData As String
    Dim lngOptionSize As Long
    
    mvntOption = Null
    
    strOptionPath = _
        "Software\Microsoft\Office" & _
        "\" & PowerPoint.Application.Version & _
        "\" & "PowerPoint" & _
        "\" & "Options"
    
    lngResult = RegOpenKeyEx _
        (HKEY_CURRENT_USER, _
         strOptionPath, _
         0, _
         KEY_QUERY_VALUE, _
         lngKeyHandle)
    If (lngResult = ERROR_SUCCESS) Then
        lngResult = RegQueryValueEx _
            (lngKeyHandle, _
             strOptionName, _
             0&, _
             lngOptionType, _
             ByVal 0&, _
             lngOptionSize)
        If (lngResult = ERROR_SUCCESS) Then
            Select Case lngOptionType
                Case REG_DWORD
                    lngResult = RegQueryValueEx _
                        (lngKeyHandle, _
                         strOptionName, _
                         0&, _
                         lngOptionType, _
                         lngOptionData, _
                         lngOptionSize)
                    If (lngResult = ERROR_SUCCESS) Then
                        mvntOption = lngOptionData
                    End If
                Case REG_SZ
                    strOptionData = VBA.String$(lngOptionSize - 1, VBA.vbNullChar)
                    lngResult = RegQueryValueEx _
                        (lngKeyHandle, _
                         strOptionName, _
                         0&, _
                         lngOptionType, _
                         ByVal strOptionData, _
                         lngOptionSize)
                    If (lngResult = ERROR_SUCCESS) Then
                        mvntOption = mstrC2VBConvert(strOptionData)
                    End If
                Case Default
                    Err.Raise _
                        VBA.vbObjectError, _
                        "WorshipServiceAssistant.WSAApplicationOptions.mvntOption", _
                        "WorshipServiceAssistant.WSAApplicationOptions.mvntOption:" & _
                            VBA.vbCrLf & _
                            "Unknown option type " & _
                            "(" & lngOptionType & ") " & _
                            "for the option " & _
                            "(" & strOptionName & ")."
            End Select
        End If
        lngResult = RegCloseKey _
            (lngKeyHandle)
    End If
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Property Let mvntOption _
( _
    ByRef strOptionName As String, _
    ByVal vntOptionData As Variant _
)
    Dim lngResult As Long
    Dim lngKeyHandle As Long
    Dim strOptionPath As String
    Dim lngOptionType As VBA.VbVarType
    Dim lngOptionData As Long
    Dim strOptionData As String
    Dim lngOptionSize As Long
    
    strOptionPath = _
        "Software\Microsoft\Office" & _
        "\" & PowerPoint.Application.Version & _
        "\" & "PowerPoint" & _
        "\" & "Options"
    
    lngResult = RegOpenKeyEx _
        (HKEY_CURRENT_USER, _
         strOptionPath, _
         0, _
         KEY_SET_VALUE, _
         lngKeyHandle)
    If (lngResult = ERROR_SUCCESS) Then
        lngOptionType = VBA.VarType(vntOptionData)
        Select Case lngOptionType
            Case VBA.vbSingle, VBA.vbInteger, VBA.vbLong
                lngOptionData = VBA.CLng(vntOptionData)
                lngOptionSize = 4
                lngResult = RegSetValueEx _
                    (lngKeyHandle, _
                     strOptionName, _
                     0&, _
                     REG_DWORD, _
                     lngOptionData, _
                     lngOptionSize)
            Case VBA.vbString
                strOptionData = VBA.CStr(vntOptionData)
                lngOptionSize = VBA.Len(strOptionData) + 1
                lngResult = RegSetValueEx _
                    (lngKeyHandle, _
                     strOptionName, _
                     0&, _
                     REG_SZ, _
                     ByVal strOptionData, _
                     lngOptionSize)
            Case Default
                Err.Raise _
                    VBA.vbObjectError, _
                    "WorshipServiceAssistant.WSAApplicationOptions.mvntOption", _
                    "WorshipServiceAssistant.WSAApplicationOptions.mvntOption:" & _
                        VBA.vbCrLf & _
                        "Unknown option type " & _
                        "(" & lngOptionType & ") " & _
                        "for the option " & _
                        "(" & strOptionName & ")."
        End Select
        If (lngResult <> ERROR_SUCCESS) Then
            Err.Raise _
                VBA.vbObjectError, _
                "WorshipServiceAssistant.WSAApplicationOptions.mvntOption", _
                "WorshipServiceAssistant.WSAApplicationOptions.mvntOption:" & _
                    VBA.vbCrLf & _
                    "Set failed " & _
                    "for the option " & _
                    "(" & strOptionName & ")."
        End If
        lngResult = RegCloseKey _
            (lngKeyHandle)
    End If
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Property Let mvntOptionIndirect _
( _
    ByRef strOptionName As String, _
    ByVal lngOptionDataExpected As Long _
)
    Dim lngOptionType As Long
    Dim vntOptionData As Variant
    Dim lngOptionData As Long
    Dim lngOptionIndex As Long
    Dim lngOptionTab As Long
    Dim strOptionShortcutKey As String
    
    For lngOptionIndex = LBound(maudtOptionIndirectList) To UBound(maudtOptionIndirectList) Step 1
        If (maudtOptionIndirectList(lngOptionIndex).Name = strOptionName) Then
            Exit For
        End If
    Next
    If (lngOptionIndex > UBound(maudtOptionIndirectList)) Then
        Err.Raise _
        VBA.vbObjectError, _
        "WorshipServiceAssistant.WSAApplicationOptions.mvntOptionIndirect", _
        "WorshipServiceAssistant.WSAApplicationOptions.mvntOptionIndirect:" & _
            VBA.vbCrLf & _
            "Unknown option " & _
            "for the option " & _
            "(" & strOptionName & ")."
    End If

    If (lngOptionIndex <= UBound(maudtOptionIndirectList)) Then
        lngOptionTab = maudtOptionIndirectList(lngOptionIndex).Tab
        strOptionShortcutKey = maudtOptionIndirectList(lngOptionIndex).ShortcutKey
        
        vntOptionData = mvntOption(strOptionName)
        If (IsNull(vntOptionData) = True) Then
            mOptionDialogTabSet lngOptionTab
            PowerPoint.Application.CommandBars.FindControl(id:=mlngPpControlIdToolsOptions).Execute
            VBA.SendKeys "%" & strOptionShortcutKey, True
            VBA.SendKeys "{ENTER}", True
            vntOptionData = mvntOption(strOptionName)
        End If
        If ((IsNull(vntOptionData) = False) And (vntOptionData <> lngOptionDataExpected)) Then
            mOptionDialogTabSet lngOptionTab
            PowerPoint.Application.CommandBars.FindControl(id:=mlngPpControlIdToolsOptions).Execute
            VBA.SendKeys "%" & strOptionShortcutKey, True
            VBA.SendKeys "{ENTER}", True
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
Private Sub mOptionDialogTabSet _
( _
    lngIndex As Long _
)
    Dim vntOptionData As Variant
    
    vntOptionData = mvntOption("Options dialog current tab")
    Do While ((IsNull(vntOptionData) = False) And (vntOptionData <> lngIndex))
        PowerPoint.Application.CommandBars.FindControl(id:=mlngPpControlIdToolsOptions).Execute
        VBA.SendKeys "^{TAB}", True
        VBA.SendKeys "{ENTER}", True
        vntOptionData = mvntOption("Options dialog current tab")
    Loop
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mstrC2VBConvert(strC As String) As String
    Dim lngIndex As Long
    Dim strChar As String
    Dim strVB As String
    
    For lngIndex = 1 To Len(strC) Step 1
        strChar = Mid(strC, lngIndex, 1)
        If ((strChar <> VBA.vbNullChar) And (strChar <> VBA.vbNullString)) Then
            strVB = strVB & strChar
        End If
    Next
    
    mstrC2VBConvert = strVB
End Function

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
    ReDim maudtOptionIndirectList(0)
    With maudtOptionIndirectList(UBound(maudtOptionIndirectList))
        .Name = "SSRightMouse"
        .Tab = 0
        .ShortcutKey = "P"
    End With
    
    ReDim Preserve maudtOptionIndirectList(UBound(maudtOptionIndirectList) + 1)
    With maudtOptionIndirectList(UBound(maudtOptionIndirectList))
        .Name = "SSMenuButton"
        .Tab = 0
        .ShortcutKey = "S"
    End With
    
    ReDim Preserve maudtOptionIndirectList(UBound(maudtOptionIndirectList) + 1)
    With maudtOptionIndirectList(UBound(maudtOptionIndirectList))
        .Name = "SSEndOnBlankSlide"
        .Tab = 0
        .ShortcutKey = "E"
    End With
    
    ReDim Preserve maudtOptionIndirectList(UBound(maudtOptionIndirectList) + 1)
    With maudtOptionIndirectList(UBound(maudtOptionIndirectList))
        .Name = "SaveAutoRecoveryInfo"
        .Tab = 4
        .ShortcutKey = "S"
    End With
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

