Attribute VB_Name = "Presentation"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Presentation
'
' Description:
'
' Author:
'   Paul Bender <pebender@san.rr.com>
'
' Copyright:
'   Copyright (C) 2001 Paul Bender
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
