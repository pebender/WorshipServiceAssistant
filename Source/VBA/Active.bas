Attribute VB_Name = "Active"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Active
'
' Description:
'
' Author:
'   Paul Bender <pebender@san.rr.com>
'
' Copyright:
'   Copyright (C) 2000, 2001 Paul Bender
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
'   1.00.0000:
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
' Public Constants.
'===============================================================================


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'   This function returns TRUE if there is an active document window and
'   the active document window is one of the following types: normal, outline,
'   notes, slide, or slide sorter.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function ActiveWindowExists() As Boolean
    On Error GoTo ActiveWindowExists_False
    
    ActiveWindowExists = True
    
    If ((Application.ActiveWindow.ViewType <> ppViewNormal) And _
        (Application.ActiveWindow.ViewType <> ppViewNotesPage) And _
        (Application.ActiveWindow.ViewType <> ppViewOutline) And _
        (Application.ActiveWindow.ViewType <> ppViewSlide) And _
        (Application.ActiveWindow.ViewType <> ppViewSlideSorter)) Then
        GoTo ActiveWindowExists_False
    End If
Exit Function

ActiveWindowExists_False:
    ActiveWindowExists = False
End Function

'-------------------------------------------------------------------------------
' Description:
'    This function returns TRUE if the document window, W, has an associated
'    active slide show window.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function ActiveSlideShowExists(ByVal W As DocumentWindow) As Boolean
    On Error GoTo ActiveSlideShowExists_False
    
    ActiveSlideShowExists = True
    
    If (IsNull(W.Presentation.SlideShowWindow) = True) Then
        GoTo ActiveSlideShowExists_False
    End If
    If (IsEmpty(W.Presentation.SlideShowWindow) = True) Then
        GoTo ActiveSlideShowExists_False
    End If
Exit Function

ActiveSlideShowExists_False:
    ActiveSlideShowExists = False
End Function

'-------------------------------------------------------------------------------
' Description:
'   This function returns TRUE if the document window, W, has at least
'   one slide.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function ActiveWindowSlideExists(ByVal W As DocumentWindow) As Boolean
    ActiveWindowSlideExists = False
    If (W.Presentation.slides.Count = 0) Then
        Exit Function
    End If
    ActiveWindowSlideExists = True
End Function

'-------------------------------------------------------------------------------
' Description:
'    This function returns TRUE if the document window, W, has a selection.
'    Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function ActiveSelectionExists(ByVal W As DocumentWindow) As Boolean
    ActiveSelectionExists = False
    If (W.Selection.Type = ppSelectionNone) Then
        Exit Function
    End If
    ActiveSelectionExists = True
End Function

'-------------------------------------------------------------------------------
' Description:
'    This function returns TRUE if the document window, W, has an active
'    slide.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function ActiveSlideExists(ByVal W As DocumentWindow) As Boolean
    ActiveSlideExists = False
    
    If (ActiveWindowSlideExists(W) = False) Then
        Exit Function
    End If
    If ((W.ViewType = ppViewSlideSorter) And _
        (ActiveSelectionExists(W) = False)) Then
        Exit Function
    End If
    ActiveSlideExists = True
End Function

'-------------------------------------------------------------------------------
' Description:
'   This function returns the active slide associated with the document
'   window, W.  This function does not test whether or not the document
'   window has an active slide.  If no active slide exists, then this
'   function will generate an error.
'-------------------------------------------------------------------------------
Public Function ActiveSlide(ByVal W As DocumentWindow) As Slide
    If (W.ViewType = ppViewSlideSorter) Then
        Set ActiveSlide = W.Selection.SlideRange(1)
    Else
        Set ActiveSlide = W.View.Slide
    End If
End Function


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
