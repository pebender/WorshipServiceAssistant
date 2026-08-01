VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSADocumentWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.WSADocumentWindow
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


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private mudtApplication As WorshipServiceAssistant.WSAApplication
Private mudtPresentation As WorshipServiceAssistant.WSAPresentation
Private mdwDoppelganger As PowerPoint.DocumentWindow

Private mintWindowState As PowerPoint.PpWindowState
Private mintViewType As PowerPoint.PpViewType
Private mintView_DisplaySlideMiniature As Office.MsoTriState
Private mintView_ZoomToFit As Office.MsoTriState


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
Public Property Get Active _
( _
) As Office.MsoTriState
    Active = mdwDoppelganger.Active
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Activate _
( _
)
    mdwDoppelganger.Activate
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   This property returns TRUE if there is a selection.
'   Otherwise, it returns FALSE.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get HasSelection _
( _
) As Boolean
    HasSelection = False
    If (mdwDoppelganger.Selection.Type = PowerPoint.ppSelectionNone) Then
        Exit Property
    End If
    HasSelection = True
End Property

'-------------------------------------------------------------------------------
' Purpose:
'    This property returns TRUE if the presentation has an active slide.
'    Otherwise, it returns FALSE.
' Assumptions:
' Effects:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get HasActiveSlide _
( _
) As Boolean
    On Error GoTo HasActiveSlide_Exit
    
    HasActiveSlide = False
    If (Me.HasSelection = True) Then
        HasActiveSlide = True
    ElseIf (VBA.IsNull(mdwDoppelganger.View.Slide) = False) Then
        HasActiveSlide = True
    End If

HasActiveSlide_Exit:
End Property

'-------------------------------------------------------------------------------
' Purpose:
'   This property returns the active slide, if there is an active slide.
'   Otherwise it returns Nothing.
' Assumptions:
' Effects:
' Returns:
'-------------------------------------------------------------------------------
Public Function ActiveSlide _
( _
) As PowerPoint.Slide
    If (Me.HasActiveSlide = True) Then
        If (Me.HasSelection = True) Then
            Set ActiveSlide = mdwDoppelganger.Selection.SlideRange(1)
        Else
            Set ActiveSlide = mdwDoppelganger.View.Slide
        End If
    Else
        Set ActiveSlide = Nothing
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Initialize _
( _
    ByRef udtApplication As WorshipServiceAssistant.WSAApplication, _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef dwDoppelganger As PowerPoint.DocumentWindow _
)
    Set mudtApplication = udtApplication
    Set mudtPresentation = udtPresentation
    Set mdwDoppelganger = dwDoppelganger
    
    mintWindowState = -1
    mintViewType = -1
    mintView_DisplaySlideMiniature = -1
    mintView_ZoomToFit = -1
    
    mConfigurationSave
    mConfigurationInitialize
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
    mConfigurationRestore
    
    mintView_ZoomToFit = -1
    mintView_DisplaySlideMiniature = -1
    mintViewType = -1
    mintWindowState = -1
    
    Set mdwDoppelganger = Nothing
    Set mudtPresentation = Nothing
    Set mudtApplication = Nothing
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
) As PowerPoint.DocumentWindow
    Set Doppelganger = mdwDoppelganger
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
    With mdwDoppelganger
        mintWindowState = .WindowState
        mintViewType = .ViewType
        If (mintViewType = PowerPoint.ppViewNormal) Then
            Select Case .Panes(2).ViewType
                Case PowerPoint.ppViewSlide
                    mintViewType = PowerPoint.ppViewNormal
                Case PowerPoint.ppViewSlideMaster
                    mintViewType = PowerPoint.ppViewSlideMaster
            End Select
        End If
        mintView_DisplaySlideMiniature = .View.DisplaySlideMiniature
        mintView_ZoomToFit = .View.ZoomToFit
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
    With mdwDoppelganger
        .WindowState = mintWindowState
        .ViewType = mintViewType
        .View.DisplaySlideMiniature = mintView_DisplaySlideMiniature
        .View.ZoomToFit = mintView_ZoomToFit
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
    With mdwDoppelganger
        .WindowState = PowerPoint.ppWindowMaximized
        .ViewType = PowerPoint.ppViewSlide
        .View.DisplaySlideMiniature = Office.msoFalse
        .View.ZoomToFit = Office.msoTrue
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
