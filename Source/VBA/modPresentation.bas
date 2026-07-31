Attribute VB_Name = "modPresentation"
'===============================================================================
' Name:
'   WorshipServiceAssistant.modPresentation
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
'-------------------------------------------------------------------------------
Public Function gblnExists _
( _
) As Boolean
    Dim prePresentation As PowerPoint.Presentation
    
    gblnExists = False
    For Each prePresentation In Application.Presentations
        If (modBanner.gblnIsBanner(prePresentation) = False) Then
            If (prePresentation.Windows.Count > 0) Then
                If (prePresentation.Slides.Count > 0) Then
                    gblnExists = True
                End If
            End If
        End If
    Next
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function gblnIsPresentation _
( _
    ByRef prePresentation As PowerPoint.Presentation _
) As Boolean
    gblnIsPresentation = False
    If (modBanner.gblnIsBanner(prePresentation) = False) Then
        If (prePresentation.Windows.Count > 0) Then
            If (prePresentation.Slides.Count > 0) Then
                gblnIsPresentation = True
            End If
        End If
    End If
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function gblnSlideShowExists _
( _
    ByVal prePresentation As PowerPoint.Presentation _
) As Boolean
    On Error GoTo gblnSlideShowExists_False
    
    gblnSlideShowExists = True
    
    If (VBA.IsNull(prePresentation.SlideShowWindow) = True) Then
        GoTo gblnSlideShowExists_False
    End If
    If (VBA.IsEmpty(prePresentation.SlideShowWindow) = True) Then
        GoTo gblnSlideShowExists_False
    End If
Exit Function

gblnSlideShowExists_False:
    gblnSlideShowExists = False
End Function


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
