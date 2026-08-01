VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSAApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.WSAApplication
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
'   1.04.0002:
'     (1) Added code to set the slide show window monitor using
'         by directly modifying the registry.
'     (2) Added code to disable PowerPoint's show presenter view.
'   1.04.0000:
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
Private Const mlngPpControlIdToolsOptions As Long = 522


'===============================================================================
' Private Variables.
'===============================================================================
Private mappDoppelganger As PowerPoint.Application

Private mudtBanner As New WorshipServiceAssistant.WSABanner
Private mudtPresentations As New WorshipServiceAssistant.WSAPresentationCollection
Private mlngSlideShowWindowMonitor As Long
Private mlngSlideShowWindowSize As Long

Private mblnWindowWindowStateMax As Boolean
Private mintWindowState As PowerPoint.PpWindowState
Private mlngTop As Long
Private mlngLeft As Long
Private mlngHeight As Long
Private mlngWidth As Long


'===============================================================================
' Public Properties and Methods.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Banner _
( _
) As WorshipServiceAssistant.WSABanner
    Set Banner = mudtBanner
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Presentations _
( _
) As WorshipServiceAssistant.WSAPresentationCollection
    Set Presentations = mudtPresentations
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get HasActivePresentation _
( _
) As Office.MsoTriState
    Dim lngIndex As Long
    
    HasActivePresentation = Office.msoFalse
    With mudtPresentations
        For lngIndex = .Count To 1 Step -1
            If (.Item(lngIndex).Active = Office.msoTrue) Then
                HasActivePresentation = Office.msoTrue
                Exit For
            End If
        Next
    End With
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get ActivePresentation _
( _
) As WorshipServiceAssistant.WSAPresentation
    Dim lngIndex As Long
    
    Set ActivePresentation = Nothing
    With mudtPresentations
        For lngIndex = .Count To 1 Step -1
            If (.Item(lngIndex).Active = Office.msoTrue) Then
                Set ActivePresentation = .Item(lngIndex)
                Exit For
            End If
        Next
    End With
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Initialize _
( _
    ByRef appDoppelganger As PowerPoint.Application _
)
    Dim dwCurrent As PowerPoint.DocumentWindow
    Dim udtPresentation As WorshipServiceAssistant.WSAPresentation
    Dim lngIndex As Long
    
    If (modActive.gblnActiveWindowExists) Then
        Set dwCurrent = PowerPoint.ActiveWindow
    Else
        Set dwCurrent = Nothing
    End If
    
    Set mappDoppelganger = appDoppelganger
    
    If ((modProject.glngSlideShowWindowMonitor >= 1) And _
        (modProject.glngSlideShowWindowMonitor <= mlngMonitorCountGet)) Then
        mlngSlideShowWindowMonitor = modProject.glngSlideShowWindowMonitor
    Else
        mlngSlideShowWindowMonitor = mlngMonitorCountGet
    End If
    
    mlngSlideShowWindowSize = modProject.glngSlideShowWindowSize
    
    mblnWindowWindowStateMax = False
    mintWindowState = 0
    mlngTop = 0
    mlngLeft = 0
    mlngHeight = 0
    mlngWidth = 0
    
    mConfigurationSave
    mConfigurationInitialize
    
    mudtBanner.Initialize Me
   
    mudtPresentations.Initialize
    
    With mappDoppelganger.Presentations
        For lngIndex = 1 To .Count Step 1
            If (((.Item(lngIndex) Is mudtBanner.Presentation.Doppelganger) = False) And _
                (.Item(lngIndex).Windows.Count >= 1) And _
                (.Item(lngIndex).Slides.Count >= 1)) Then
                Set udtPresentation = New WorshipServiceAssistant.WSAPresentation
                udtPresentation.Initialize Me, .Item(lngIndex)
                mudtPresentations.Add udtPresentation
            End If
        Next
    End With
    
    If ((dwCurrent Is Nothing) = False) Then
        dwCurrent.Activate
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Terminate _
( _
)
    Dim dwCurrent As PowerPoint.DocumentWindow
    
    If (modActive.gblnActiveWindowExists) Then
        Set dwCurrent = PowerPoint.ActiveWindow
    Else
        Set dwCurrent = Nothing
    End If
    
    With mudtPresentations
        While (.Count > 0)
            .Item(.Count).Terminate
            .Remove .Item(.Count)
        Wend
        .Terminate
    End With
    mudtPresentations.Terminate
    
    mudtBanner.Terminate
    
    mConfigurationRestore
    
    mlngWidth = 0
    mlngHeight = 0
    mlngLeft = 0
    mlngTop = 0
    mintWindowState = 0
    mblnWindowWindowStateMax = False
    
    mlngSlideShowWindowSize = 0
    mlngSlideShowWindowMonitor = 0
    
    Set mappDoppelganger = Nothing
    
    mlngSlideShowWindowMonitor = 0
    mlngSlideShowWindowSize = 0

    If ((dwCurrent Is Nothing) = False) Then
        dwCurrent.Activate
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Doppelganger _
( _
) As PowerPoint.Application
    Set Doppelganger = mappDoppelganger
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get SlideShowWindowMonitor _
( _
)
    SlideShowWindowMonitor = mlngSlideShowWindowMonitor
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get SlideShowWindowSize _
( _
)
    SlideShowWindowSize = mlngSlideShowWindowSize
End Property


'===============================================================================
' Private Properties and Methods.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub mConfigurationSave _
( _
)
    Dim lngIndex As Long
            
    With mappDoppelganger
        mintWindowState = .WindowState
        .WindowState = PowerPoint.ppWindowNormal
        mlngTop = .Top
        mlngLeft = .Left
        mlngHeight = .Height
        mlngWidth = .Width
        .WindowState = mintWindowState
    End With
    
    ' This should not be needed.
    ' However, while each has a separate WindowState, they seemed to be linked.
    With mappDoppelganger.Windows
        mblnWindowWindowStateMax = False
        For lngIndex = 1 To .Count Step 1
            If (.Item(lngIndex).WindowState = PowerPoint.ppWindowMaximized) Then
                mblnWindowWindowStateMax = True
                Exit For
            End If
        Next
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub mConfigurationRestore _
( _
)
    Dim lngIndex As Long
    
    With mappDoppelganger.Windows
        If (mblnWindowWindowStateMax = True) Then
            For lngIndex = 1 To .Count Step 1
                .Item(lngIndex).WindowState = PowerPoint.ppWindowMaximized
            Next
        End If
    End With
    
    With mappDoppelganger
        .WindowState = PowerPoint.ppWindowNormal
        .Height = 0
        .Width = 0
        .Top = mlngTop
        .Left = mlngLeft
        .Height = mlngHeight
        .Width = mlngWidth
        .WindowState = mintWindowState
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub mConfigurationInitialize _
( _
)
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngHeight As Long
    Dim lngWidth As Long
    Dim lngIndex As Long
    
    With mappDoppelganger
        .WindowState = PowerPoint.ppWindowMaximized
        lngTop = .Top + 3
        lngLeft = .Left + 3
        lngHeight = .Height - 6
        lngWidth = .Width - 6
        .WindowState = PowerPoint.ppWindowNormal
        .Height = 0
        .Width = 0
        .Top = lngTop
        .Left = lngLeft
        .Height = lngHeight
        .Width = lngWidth
    End With
    
    ' Set the slide show window monitor.
    mblnOptionSetSZ "DisplayMonitor", mstrMonitorNameGet(Me.SlideShowWindowMonitor)
    
    ' Hide presenter view.
    mblnOptionSetDWORD "UseMonMgr", 0
    
    ' Configure the PowerPoint options to the most appropiate values
    ' for Worship Service Assistant.
    mblnOptionSetDWORDIndirect "SSRightMouse", 0
    mblnOptionSetDWORDIndirect "SSMenuButton", 0
    mblnOptionSetDWORDIndirect "SSEndOnBlankSlide", 0
    mblnOptionSetDWORDIndirect "SaveAutoRecoveryInfo", 0

    ' Hide any floating, pop-up and other annoying command bars.
    With Me.Doppelganger.CommandBars
        For lngIndex = 1 To .Count Step 1
            With .Item(lngIndex)
                If ((.Position = Office.msoBarFloating) Or _
                    (.Position = Office.msoBarPopup)) Then
                    If (.Visible = True) Then
                        .Visible = False
                    End If
                End If
                If (.Name = "Task Pane") Then
                    If (.Visible = True) Then
                        .Visible = False
                    End If
                End If
            End With
        Next
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mblnOptionGetDWORD _
( _
    ByRef strOptionName As String, _
    ByRef lngOptionValue As Long _
) As Boolean
    Dim lngResult As Long
    Dim lngKeyHandle As Long
    Dim strOptionPath As String
    Dim lngOptionType As Long
    Dim lngOptionBuffer As Long
    Dim lngOptionBufferSize As Long
    
    mblnOptionGetDWORD = False
    strOptionPath = _
        "Software\Microsoft\Office" & _
        "\" & Me.Doppelganger.VERSION & _
        "\" & "PowerPoint" & _
        "\" & "Options"
    
    lngResult = modWin32AdvAPI32.RegOpenKeyEx _
        (modWin32AdvAPI32.HKEY_CURRENT_USER, _
         strOptionPath, _
         0, _
         modWin32AdvAPI32.KEY_QUERY_VALUE, _
         lngKeyHandle)
    If (lngResult = modWin32AdvAPI32.ERROR_SUCCESS) Then
        lngResult = modWin32AdvAPI32.RegQueryValueEx _
            (lngKeyHandle, _
             strOptionName, _
             0&, _
             lngOptionType, _
             ByVal 0&, _
             lngOptionBufferSize)
        If (lngResult = modWin32AdvAPI32.ERROR_SUCCESS) Then
            If (lngOptionType = modWin32AdvAPI32.REG_DWORD) Then
                lngResult = modWin32AdvAPI32.RegQueryValueEx _
                    (lngKeyHandle, _
                     strOptionName, _
                     0&, _
                     lngOptionType, _
                     lngOptionBuffer, _
                     lngOptionBufferSize)
                If (lngResult = modWin32AdvAPI32.ERROR_SUCCESS) Then
                    mblnOptionGetDWORD = True
                    lngOptionValue = lngOptionBuffer
                End If
            End If
        End If
        lngResult = modWin32AdvAPI32.RegCloseKey _
            (lngKeyHandle)
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mblnOptionSetDWORD _
( _
    ByRef strOptionName As String, _
    ByRef lngOptionValue As Long _
)
    Dim lngResult As Long
    Dim lngKeyHandle As Long
    Dim strOptionPath As String
    
    mblnOptionSetDWORD = False
    
    strOptionPath = _
        "Software\Microsoft\Office" & _
        "\" & PowerPoint.Application.VERSION & _
        "\" & "PowerPoint" & _
        "\" & "Options"
    
    lngResult = modWin32AdvAPI32.RegOpenKeyEx _
        (modWin32AdvAPI32.HKEY_CURRENT_USER, _
         strOptionPath, _
         0, _
         modWin32AdvAPI32.KEY_SET_VALUE, _
         lngKeyHandle)
    If (lngResult = modWin32AdvAPI32.ERROR_SUCCESS) Then
        lngResult = modWin32AdvAPI32.RegSetValueEx _
            (lngKeyHandle, _
             strOptionName, _
             0&, _
             modWin32AdvAPI32.REG_DWORD, _
             lngOptionValue, _
             4)
        If (lngResult = modWin32AdvAPI32.ERROR_SUCCESS) Then
            mblnOptionSetDWORD = True
        End If
        lngResult = modWin32AdvAPI32.RegCloseKey _
            (lngKeyHandle)
    End If
End Function


'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mblnOptionSetSZ _
( _
    ByRef strOptionName As String, _
    ByRef strOptionValue As String _
) As Boolean
    Dim lngResult As Long
    Dim lngKeyHandle As Long
    Dim strOptionPath As String
    
    mblnOptionSetSZ = False
    
    strOptionPath = _
        "Software\Microsoft\Office" & _
        "\" & PowerPoint.Application.VERSION & _
        "\" & "PowerPoint" & _
        "\" & "Options"
    
    lngResult = modWin32AdvAPI32.RegOpenKeyEx _
        (modWin32AdvAPI32.HKEY_CURRENT_USER, _
         strOptionPath, _
         0, _
         modWin32AdvAPI32.KEY_SET_VALUE, _
         lngKeyHandle)
    If (lngResult = modWin32AdvAPI32.ERROR_SUCCESS) Then
        lngResult = modWin32AdvAPI32.RegSetValueExSZ _
            (lngKeyHandle, _
             strOptionName, _
             0&, _
             modWin32AdvAPI32.REG_SZ, _
             strOptionValue, _
             Len(strOptionValue) + 1)
        If (lngResult = modWin32AdvAPI32.ERROR_SUCCESS) Then
            mblnOptionSetSZ = True
        End If
        lngResult = modWin32AdvAPI32.RegCloseKey _
            (lngKeyHandle)
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mblnOptionSetDWORDIndirect _
( _
    ByRef strOptionName As String, _
    ByRef lngOptionValue As Long _
) As Boolean
    Dim avntOptionList(3, 2) As Variant
    
    avntOptionList(0, 0) = "SSRightMouse":         avntOptionList(0, 1) = 0: avntOptionList(0, 2) = "P"
    avntOptionList(1, 0) = "SSMenuButton":         avntOptionList(1, 1) = 0: avntOptionList(1, 2) = "S"
    avntOptionList(2, 0) = "SSEndOnBlankSlide":    avntOptionList(2, 1) = 0: avntOptionList(2, 2) = "E"
    avntOptionList(3, 0) = "SaveAutoRecoveryInfo": avntOptionList(3, 1) = 4: avntOptionList(3, 2) = "S"
    
    Dim blnResultValid As Boolean
    Dim lngResultValue As Long
    Dim lngOptionIndex As Long
    Dim lngOptionTab As Long
    Dim strOptionShortcutKey As String
    
    mblnOptionSetDWORDIndirect = False
    For lngOptionIndex = LBound(avntOptionList, 1) To UBound(avntOptionList, 1) Step 1
        If (avntOptionList(lngOptionIndex, 0) = strOptionName) Then
            Exit For
        End If
    Next
    If (lngOptionIndex <= UBound(avntOptionList)) Then
        lngOptionTab = avntOptionList(lngOptionIndex, 1)
        strOptionShortcutKey = avntOptionList(lngOptionIndex, 2)
        
        blnResultValid = mblnOptionGetDWORD(strOptionName, lngResultValue)
        If (blnResultValid = False) Then
            mOptionDialogTabSet lngOptionTab
            Me.Doppelganger.CommandBars.FindControl(id:=mlngPpControlIdToolsOptions).Execute
            VBA.SendKeys "%" & strOptionShortcutKey, True
            VBA.SendKeys "{ENTER}", True
            blnResultValid = mblnOptionGetDWORD(strOptionName, lngResultValue)
        End If
        If ((blnResultValid = True) And (lngResultValue <> lngOptionValue)) Then
            mOptionDialogTabSet lngOptionTab
            Me.Doppelganger.CommandBars.FindControl(id:=mlngPpControlIdToolsOptions).Execute
            VBA.SendKeys "%" & strOptionShortcutKey, True
            VBA.SendKeys "{ENTER}", True
        End If
        mblnOptionSetDWORDIndirect = True
    End If
End Function

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
    Dim blnResult As Boolean
    Dim lngValue As Long
    
    blnResult = mblnOptionGetDWORD("Options dialog current tab", lngValue)
    If ((blnResult = True) And (lngValue <> lngIndex)) Then
        While ((blnResult = True) And (lngValue <> lngIndex))
            Me.Doppelganger.CommandBars.FindControl(id:=mlngPpControlIdToolsOptions).Execute
            VBA.SendKeys "^{TAB}", True
            VBA.SendKeys "{ENTER}", True
            blnResult = mblnOptionGetDWORD("Options dialog current tab", lngValue)
        Wend
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mlngMonitorCountGet _
( _
) As Long
    mlngMonitorCountGet = modWin32User32.GetSystemMetrics(modWin32User32.SM_CMONITORS)
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mstrMonitorNameGet _
( _
    ByRef lngMonitorNumber As Long _
) As String
    Dim udtDisplay As modWin32User32.DISPLAY_DEVICE
    Dim udtMonitor As modWin32User32.DISPLAY_DEVICE
    Dim lngDisplay As Long
    Dim lngMonitor As Long
    Dim strMonitorName As String
    Dim lngMonitorIndex As Long
    
    udtDisplay.cb = Len(udtDisplay)
    udtMonitor.cb = Len(udtMonitor)
    
    lngMonitorIndex = 0
    
    ' Loop through the displays.
    lngDisplay = 0
    Do While (EnumDisplayDevices(VBA.vbNullString, lngDisplay, udtDisplay, 0) <> 0)
        ' Skip displays that are not attached to the desktop.
        If ((udtDisplay.StateFlags And modWin32User32.DISPLAY_DEVICE_ATTACHED_TO_DESKTOP) <> 0) Then
            ' Skip displays that are not real.
            If ((udtDisplay.StateFlags And modWin32User32.DISPLAY_DEVICE_MIRRORING_DRIVER) = 0) Then
                ' Loop through monitors.
                lngMonitor = 0
                Do While (EnumDisplayDevices(udtDisplay.DeviceName, lngMonitor, udtMonitor, 0) <> 0)
                    If ((udtMonitor.StateFlags And modWin32User32.DISPLAY_DEVICE_ACTIVE) <> 0) Then
                        Exit Do
                    End If
                    lngMonitor = lngMonitor + 1
                Loop
                lngMonitorIndex = lngMonitorIndex + 1
                
                If (lngMonitorIndex = lngMonitorNumber) Then
                    If (udtMonitor.DeviceName <> "") Then
                        strMonitorName = mstrC2VBConvert(udtMonitor.DeviceName)
                    Else
                        strMonitorName = mstrC2VBConvert(udtDisplay.DeviceName)
                    End If
                End If
            End If
        End If
        lngDisplay = lngDisplay + 1
    Loop
    mstrMonitorNameGet = strMonitorName
End Function

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

