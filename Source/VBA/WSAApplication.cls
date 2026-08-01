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
'   1.04.0003:
'     (1) Moved routines associated with setting the PowerPoint options
'         to the WSAApplicationOptions class.
'     (2) Modified mConfigurationInitialize to set the PowerPoint options
'         using the WSAApplicationOptions class.
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
' Private Constants.
'===============================================================================
Private Const mlngPpControlIdToolsOptions As Long = 522


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private mappDoppelganger As PowerPoint.Application

Private mudtBanner As New WorshipServiceAssistant.WSABanner
Private mudtPresentations As New WorshipServiceAssistant.WSAPresentationCollection
Private mstrSlideShowWindowMonitor As String
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
    
    Set dwCurrent = mActiveWindow
    
    Set mappDoppelganger = appDoppelganger
    
    mstrSlideShowWindowMonitor = modProject.gstrSlideShowWindowMonitor
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
    
    Set dwCurrent = mActiveWindow
    
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
    mstrSlideShowWindowMonitor = ""
    
    Set mappDoppelganger = Nothing
    
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
) As String
    SlideShowWindowMonitor = mstrSlideShowWindowMonitor
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
) As Long
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
    Dim udtOptions As New WorshipServiceAssistant.WSAApplicationOptions
    
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
    
    With udtOptions
        .DisplayMonitor = Me.SlideShowWindowMonitor
        .SaveAutoRecoveryInfo = 0
        .SSEndOnBlankSlide = 0
        .SSMenuButton = 0
        .SSRightMouse = 0
        .UseMonMgr = 0
    End With
    
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
Private Function mActiveWindow _
( _
) As PowerPoint.DocumentWindow
    On Error GoTo mActiveWindow_Exit
    
    Set mActiveWindow = Nothing
    
    If ((PowerPoint.Application.ActiveWindow Is Nothing) = False) Then
        Set mActiveWindow = PowerPoint.Application.ActiveWindow
    End If
    
mActiveWindow_Exit:
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

