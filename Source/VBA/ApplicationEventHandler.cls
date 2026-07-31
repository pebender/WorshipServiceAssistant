VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.ApplicationEventHandlers
'
' Description:
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
'   1.03.0002:
'     (1) Made changes to the source code so that it follows Microsoft's
'         Visual Basic coding conventions.
'   1.00.0000:
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

'
' Application variable for setting application events.
'
Public WithEvents PPTApplication As Application
Attribute PPTApplication.VB_VarHelpID = -1


'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================


'===============================================================================
' Private Event Handlers.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPApplication_NewPresentation _
( _
    ByVal prePresentation As PowerPoint.Presentation _
)
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_PresentationClose _
( _
    ByVal prePresentation As PowerPoint.Presentation _
)
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_PresentationNewSlide _
( _
    ByVal sldSlide As PowerPoint.Slide _
)
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_PresentationOpen _
( _
    ByVal prePresentation As PowerPoint.Presentation _
)
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_PresentationPrint _
( _
    ByVal prePresentation As PowerPoint.Presentation _
)
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_PresentationSave _
( _
    ByVal prePresentation As PowerPoint.Presentation _
)
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_SlideShowBegin _
( _
    ByVal sswSlideShowWindow As PowerPoint.SlideShowWindow _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_SlideShowEnd _
( _
    ByVal prePresentation As PowerPoint.Presentation _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_SlideShowNextBuild _
( _
    ByVal sswSlideShowWindow As PowerPoint.SlideShowWindow _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_SlideShowNextSlide _
( _
    ByVal sswSlideShowWindow As PowerPoint.SlideShowWindow _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_WindowActivate _
( _
    ByVal prePresentation As PowerPoint.Presentation, _
    ByVal dwDocumentWindow As PowerPoint.DocumentWindow _
)
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_WindowBeforeDoubleClick _
( _
    ByVal selSelection As PowerPoint.Selection, _
    blnCancel As Boolean _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_WindowBeforeRightClick _
( _
    ByVal selSelection As PowerPoint.Selection, _
    blnCancel As Boolean _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_WindowDeactivate _
( _
    ByVal prePresentation As PowerPoint.Presentation, _
    ByVal dwDocumentWindow As PowerPoint.DocumentWindow _
)
    modToolbar.gDisable
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_WindowSelectionChange _
( _
    ByVal selSelection As PowerPoint.Selection _
)
    modToolbar.gRefresh
End Sub

