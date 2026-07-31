Attribute VB_Name = "Presentation"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Presentation
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
Public Function Exists() As Boolean
    Dim P As PowerPoint.Presentation
    
    Exists = False
    For Each P In Application.Presentations
        If (Banner.IsBanner(P) = False) Then
            If (P.Windows.Count > 0) Then
                If (P.slides.Count > 0) Then
                    Exists = True
                End If
            End If
        End If
    Next
End Function

Public Function IsPresentation(ByVal P As PowerPoint.Presentation) As Boolean
    IsPresentation = False
    If (Banner.IsBanner(P) = False) Then
        If (P.Windows.Count > 0) Then
            If (P.slides.Count > 0) Then
                IsPresentation = True
            End If
        End If
    End If
End Function

Public Function SlideShowExists(ByVal P As PowerPoint.Presentation) As Boolean
    On Error GoTo SlideShowExists_False
    
    SlideShowExists = True
    
    If (IsNull(P.SlideShowWindow) = True) Then
        GoTo SlideShowExists_False
    End If
    If (IsEmpty(P.SlideShowWindow) = True) Then
        GoTo SlideShowExists_False
    End If
Exit Function

SlideShowExists_False:
    SlideShowExists = False
End Function


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
