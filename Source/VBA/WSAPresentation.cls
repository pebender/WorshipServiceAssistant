VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSAPresentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.WSAPresentation
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
Private mpreDoppelganger As PowerPoint.Presentation

Private mudtSlides As New WorshipServiceAssistant.WSASlideCollection
Private mudtWindows As New WorshipServiceAssistant.WSADocumentWindowCollection
Private mudtSlideShowWindow As New WorshipServiceAssistant.WSASlideShowWindow


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
Public Property Get Name _
( _
) As String
    Name = mpreDoppelganger.Name
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Slides _
( _
) As WorshipServiceAssistant.WSASlideCollection
    Set Slides = mudtSlides
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Windows _
( _
) As WorshipServiceAssistant.WSADocumentWindowCollection
    Set Windows = mudtWindows
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get SlideShowWindow _
( _
) As WorshipServiceAssistant.WSASlideShowWindow
    Set SlideShowWindow = mudtSlideShowWindow
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Active _
( _
) As Office.MsoTriState
    Dim lngIndex As Long
    
    Active = Office.msoFalse
    With mudtWindows
        For lngIndex = .Count To 1 Step -1
            If (.Item(lngIndex).Active = Office.msoTrue) Then
                Active = Office.msoTrue
                Exit For
            End If
        Next
    End With
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Activate _
( _
)
    With mudtWindows
        If (.Count >= 1) Then
            .Item(1).Activate
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Initialize _
( _
    ByRef udtApplication As WorshipServiceAssistant.WSAApplication, _
    ByRef preDoppelganger As PowerPoint.Presentation _
)
    Dim udtSlide As WorshipServiceAssistant.WSASlide
    Dim udtWindow As WorshipServiceAssistant.WSADocumentWindow
    Dim lngIndex As Long
    
    Set mudtApplication = udtApplication
    Set mpreDoppelganger = preDoppelganger
    
    mudtSlides.Initialize
    With preDoppelganger.Slides
        For lngIndex = 1 To .Count Step 1
            Set udtSlide = New WorshipServiceAssistant.WSASlide
            udtSlide.Initialize mudtApplication, Me, .Item(lngIndex)
            mudtSlides.Add udtSlide
        Next
    End With
        
    mudtWindows.Initialize
    With preDoppelganger.Windows
        For lngIndex = 1 To .Count Step 1
            Set udtWindow = New WorshipServiceAssistant.WSADocumentWindow
            udtWindow.Initialize mudtApplication, Me, .Item(lngIndex)
            mudtWindows.Add udtWindow
        Next
    End With
    
    mudtSlideShowWindow.Initialize mudtApplication, Me
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
    mudtSlideShowWindow.Terminate
    
    With mudtWindows
        While (.Count > 0)
            .Item(.Count).Terminate
            .Remove .Item(.Count)
        Wend
    End With
    mudtWindows.Terminate
    
    With mudtSlides
        While (.Count > 0)
            .Item(.Count).Terminate
            .Remove .Item(.Count)
        Wend
    End With
    mudtSlides.Terminate
    
    Set mpreDoppelganger = Nothing
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
) As PowerPoint.Presentation
    Set Doppelganger = mpreDoppelganger
End Property


'===============================================================================
' Private Properties and Methods.
'===============================================================================


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
