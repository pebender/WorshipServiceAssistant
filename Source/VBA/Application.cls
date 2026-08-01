VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.Application
'
' Description:
'
' Author:
'   Paul Bender <pbender@alumni.ucsd.edu>
'
' Copyright:
'   Copyright (c) 2000,2001,2002 Paul Bender
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
'   2.00.0000:
'     (1) Reset change history.
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

Public Enum WsaMode
    wsaModeOperator = 0
    wsaModeEditor = 1
    wsaModeDeveloper = 2
End Enum


'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private mblnInvariant As Boolean
Private mappDoppelganger As PowerPoint.Application
Attribute mappDoppelganger.VB_VarHelpID = -1

Private mudtMode As WorshipServiceAssistant.WsaMode
Private mstrSlideShowWindowMonitor As String
Private mlngSlideShowWindowSize As Long

Private mudtHelp As New WorshipServiceAssistant.Help
Private mudtOptions As New WorshipServiceAssistant.Options
Private mudtBanner As New WorshipServiceAssistant.Banner
Private mudtPresentations As New WorshipServiceAssistant.Presentations
Private mudtToolbar As New WorshipServiceAssistant.Toolbar
Private mudtEventHandlers As New WorshipServiceAssistant.ApplicationEventHandlers

Private mblnWindowWindowStateMax As Boolean
Private mlngWindowState As PowerPoint.PpWindowState
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
Public Sub Initialize _
( _
    ByRef appDoppelganger As PowerPoint.Application _
)
    mblnInvariant = False
    
    Set mappDoppelganger = appDoppelganger
    
    mudtHelp.Initialize Me, Me
    mudtOptions.Initialize Me, Me
    
    mudtMode = WsaMode.wsaModeOperator
    mstrSlideShowWindowMonitor = mudtOptions.DisplayMonitorList.Item(mudtOptions.DisplayMonitorList.Count)
    mlngSlideShowWindowSize = 1
    
    mudtBanner.Initialize Me, Me
    mudtPresentations.Initialize Me, Me, mappDoppelganger.Presentations
    mudtToolbar.Initialize Me, Me
    
    mudtEventHandlers.Initialize Me, Me
    
    mudtToolbar.Update
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
    mudtEventHandlers.Terminate
    mudtToolbar.Terminate
    mudtPresentations.Terminate
    mudtBanner.Terminate
    mudtOptions.Terminate
    mudtHelp.Terminate
    
    mblnInvariant = False
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Invariant _
( _
) As Boolean
    Invariant = mblnInvariant
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Let Invariant _
( _
    ByRef blnInvariant As Boolean _
)
    Dim dw As PowerPoint.DocumentWindow
    
    If (Not blnInvariant = mblnInvariant) Then
        Set dw = Nothing
        On Error Resume Next
        Set dw = mappDoppelganger.ActiveWindow
        On Error GoTo 0
        
        If (blnInvariant = True) Then
            mConfigurationSave
            mConfigurationInitialize
            mblnInvariant = blnInvariant
        End If
        
        Me.Banner.Invariant = blnInvariant
        Me.Presentations.Invariant = blnInvariant
        
        If (blnInvariant = False) Then
            mblnInvariant = blnInvariant
            mConfigurationRestore
        End If
        
        If (Not dw Is Nothing) Then dw.Activate
    End If
End Property

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
Public Property Get Help _
( _
) As WorshipServiceAssistant.Help
    Set Help = mudtHelp
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Options _
( _
) As WorshipServiceAssistant.Options
    Set Options = mudtOptions
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Toolbar _
( _
) As WorshipServiceAssistant.Toolbar
    Set Toolbar = mudtToolbar
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Banner _
( _
) As WorshipServiceAssistant.Banner
    Set Banner = mudtBanner
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
) As WorshipServiceAssistant.Presentation
    Set ActivePresentation = Presentations.Active
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get ActiveWindow _
( _
) As WorshipServiceAssistant.DocumentWindow
    Dim udtActivePresentation As WorshipServiceAssistant.Presentation
    
    Set udtActivePresentation = Me.ActivePresentation
    
    If (udtActivePresentation Is Nothing) Then
        Set ActiveWindow = Nothing
    Else
        Set ActiveWindow = udtActivePresentation.Windows.Active
    End If
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get ActiveSlide _
( _
) As WorshipServiceAssistant.Slide
    Dim udtActiveWindow As WorshipServiceAssistant.DocumentWindow
    Dim lng As Long
    
    Set udtActiveWindow = Me.ActiveWindow
    If (udtActiveWindow Is Nothing) Then
        Set ActiveSlide = Nothing
    Else
        Set ActiveSlide = udtActiveWindow.ActiveSlide
    End If
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get ActiveSlideSelection _
( _
) As WorshipServiceAssistant.SlideRange
    Dim udtActiveWindow As WorshipServiceAssistant.DocumentWindow
    Dim lng As Long
    
    Set udtActiveWindow = Me.ActiveWindow
    If (udtActiveWindow Is Nothing) Then
        Set ActiveSlideSelection = New WorshipServiceAssistant.SlideRange
    Else
        Set ActiveSlideSelection = udtActiveWindow.ActiveSlideSelection
    End If
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
) As WorshipServiceAssistant.Presentations
    Set Presentations = mudtPresentations
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Mode _
( _
) As WorshipServiceAssistant.WsaMode
    Mode = mudtMode
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Let Mode _
( _
    ByRef udtMode As WorshipServiceAssistant.WsaMode _
)
    mudtMode = udtMode
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
Public Property Let SlideShowWindowMonitor _
( _
    ByRef strSlideShowWindowMonitor As String _
)
    mstrSlideShowWindowMonitor = strSlideShowWindowMonitor
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

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Let SlideShowWindowSize _
( _
    ByRef lngSlideShowWindowSize As Long _
)
    mlngSlideShowWindowSize = lngSlideShowWindowSize
End Property

'-------------------------------------------------------------------------------
' Purpose:
'   Activate the slide navigator. In order to exit, the slide navigator must be
'   unloaded. If the slide navigator is only hidden, then this routine will
'   automatically refresh it and re-show it. This is a hack to work around
'   some focus problems resulting from activating slide shows and activating
'   new presentations.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub ActivateNavigator _
( _
)
    Dim objForm As Object
    Dim blnLoaded As Boolean
    
    
    ' Show the Navigator form until it is unloaded.
    blnLoaded = False
    Do
        If (blnLoaded) Then
            Navigator.gUpdate
        End If
        Navigator.Show
        blnLoaded = False
        For Each objForm In VBA.UserForms
            If (objForm.Name = "Navigator") Then
                blnLoaded = True
                Exit For
            End If
        Next
    Loop Until (Not blnLoaded)
End Sub


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
Private Sub mConfigurationSave _
( _
)
    Dim lngIndex As Long
            
    With mappDoppelganger
        mlngWindowState = .WindowState
        .WindowState = PowerPoint.ppWindowNormal
        mlngTop = .Top
        mlngLeft = .Left
        mlngHeight = .Height
        mlngWidth = .Width
        .WindowState = mlngWindowState
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
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mConfigurationRestore _
( _
)
    Dim lngIndex As Long
    
    With mappDoppelganger.Windows
        If (mblnWindowWindowStateMax) Then
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
        .WindowState = mlngWindowState
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
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
    
    With mudtOptions
        .DisplayMonitor = Me.SlideShowWindowMonitor
        .SaveAutoRecoveryInfo = 0
        .SSEndOnBlankSlide = 0
        .SSMenuButton = 0
        .SSRightMouse = 0
        .UseMonMgr = 0
    End With
    
    ' Hide any floating, pop-up and other annoying command bars.
    With mappDoppelganger.CommandBars
        For lngIndex = .Count To 1 Step -1
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
