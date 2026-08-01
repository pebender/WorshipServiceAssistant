VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Presentations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.Presentations
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
Private mudtApplication As WorshipServiceAssistant.Application
Private mudtParent As WorshipServiceAssistant.Application
Private mpreDoppelganger As PowerPoint.Presentations
Private mcpreRemoved As New VBA.Collection
Private mcudtPresentations As New VBA.Collection


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
    ByRef udtApplication As WorshipServiceAssistant.Application, _
    ByRef udtParent As WorshipServiceAssistant.Application, _
    ByRef preDoppelganger As PowerPoint.Presentations _
)
    Dim lng As Long
    
    mblnInvariant = False
    
    Set mudtApplication = udtApplication
    Set mudtParent = udtParent
    Set mpreDoppelganger = preDoppelganger
    
    With mcpreRemoved
        Do While (.Count > 0)
            .Remove .Count
        Loop
    End With
    With mcudtPresentations
        Do While (.Count > 0)
            .Remove .Count
        Loop
    End With
    
    mUpdate
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
    With mcudtPresentations
        Do While (.Count > 0)
            .Item(.Count).Terminate
            .Remove .Count
        Loop
    End With
    With mcpreRemoved
        Do While (.Count > 0)
            .Remove .Count
        Loop
    End With
    
    Set mpreDoppelganger = Nothing
    Set mudtParent = Nothing
    Set mudtApplication = Nothing
    
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
    Dim lng As Long
    
    If (Not blnInvariant = mblnInvariant) Then
        If (blnInvariant = True) Then
            mUpdate
            mblnInvariant = blnInvariant
        End If
        With mcudtPresentations
            For lng = 1 To .Count Step 1
                .Item(lng).Invariant = blnInvariant
            Next
        End With
        If (blnInvariant = False) Then
            mblnInvariant = blnInvariant
        End If
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
) As PowerPoint.Presentations
    Set Doppelganger = mpreDoppelganger
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Application _
( _
) As WorshipServiceAssistant.Application
    Set Application = mudtApplication
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Parent _
( _
) As WorshipServiceAssistant.Application
    Set Parent = mudtParent
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Count _
( _
) As Long
    If (Not mblnInvariant) Then mUpdate
    
    Count = mcudtPresentations.Count
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function Item _
( _
    ByVal lngIndex As Long _
) As WorshipServiceAssistant.Presentation
Attribute Item.VB_UserMemId = 0
    If (Not mblnInvariant) Then mUpdate
    
    Set Item = mcudtPresentations.Item(lngIndex)
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Add _
( _
    ByRef preAdd As PowerPoint.Presentation _
)
    Dim lng As Long
    
    If (mblnValid(preAdd)) Then
    ' preAdd is a valid WorshipServiceAssistant presentation.
        With mcudtPresentations
            ' Find out if preAdd is in the WorshipServiceAssistant presentations.
            For lng = 1 To .Count Step 1
                If (.Item(lng).Doppelganger Is preAdd) Then
                    Exit For
                End If
            Next
            If (lng > .Count) Then
            ' preAdd is not in the WorshipServiceAssistant presentations.
                ' Add preAdd to the WorshipServiceAssistant presentations.
                .Add New WorshipServiceAssistant.Presentation
                .Item(.Count).Initialize mudtApplication, mudtParent, preAdd
            End If
        End With
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Remove _
( _
    ByRef preRemove As PowerPoint.Presentation _
)
    Dim lng As Long
    
    ' Remove preRemove from the WorshipServiceAssistant presentations.
    With mcudtPresentations
        For lng = .Count To 1 Step -1
            If (.Item(lng).Doppelganger Is preRemove) Then
                .Item(lng).Terminate
                .Remove lng
            End If
        Next
    End With
    
    ' Track all yet-to-be-removed presentations until they have been
    ' removed from the PowerPoint Presentations collection. the mUpdate
    ' method removes presentations from the mcpreRemoved collection after
    ' PowerPoint removes the presentations from the PowerPoint Presentations
    ' collection.
    mcpreRemoved.Add preRemove
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Active _
( _
) As WorshipServiceAssistant.Presentation
    Dim lng As Long

    If (Not mblnInvariant) Then mUpdate
    
    Set Active = Nothing
    With mcudtPresentations
        For lng = 1 To .Count Step 1
            If (.Item(lng).Active = Office.msoTrue) Then
                Set Active = .Item(lng)
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
Public Function Index _
( _
    ByRef udtItem As WorshipServiceAssistant.Presentation _
) As Long
    Dim lng As Long
    
    If (Not mblnInvariant) Then mUpdate
    
    With mcudtPresentations
        For lng = 1 To .Count Step 1
            If (.Item(lng) Is udtItem) Then
                Exit For
            End If
        Next
        Index = lng
    End With
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function Find _
( _
    ByRef preDoppelganger As PowerPoint.Presentation _
) As WorshipServiceAssistant.Presentation
    Dim lng As Long
    
    If (Not mblnInvariant) Then mUpdate
    
    Set Find = Nothing
    With mcudtPresentations
        For lng = 1 To .Count Step 1
            If (.Item(lng).Doppelganger Is preDoppelganger) Then
                Exit For
            End If
        Next
        If (lng <= .Count) Then Set Find = .Item(lng)
    End With
End Function


'===============================================================================
' Private Properties and Methods.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
'   This method updates mcudtPresentations.
'
'   While not invariant, any Presentations property or method that uses
'   mcudtPresentations should call this method first.
'
'   Ideally, the updating of mcudtPresentations would be triggered by
'   PowerPoint Application events resulting from adding and removing
'   presentations and slides. If this were done, then this method would not
'   be needed. However, PowerPoint events are not reliable. PowerPoint removes
'   windowless presentations without triggering any events. PowerPoint has no
'   reliable method for detecting the adding or removing of slides through
'   PowerPoint Application events.
'
'   When PowerPoint triggers the PowerPoint Application PresentationClose
'   the presentation being closed is still part of the PowerPoint Presentations
'   collection. These yet-to-be-removed presentations are tracked using the
'   mcpreRemoved collection. PowerPoint presentations in this collection are
'   considered already removed and are ignored.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mUpdate _
( _
)
    Dim lng As Long
    Dim pre As PowerPoint.Presentation
    
    ' Update removed PowerPoint presentations.
    With mcpreRemoved
        For lng = .Count To 1 Step -1
            For Each pre In mpreDoppelganger
                If (.Item(lng) Is pre) Then
                    Exit For
                End If
            Next
            If (Not .Item(lng) Is pre) Then
                .Remove lng
            End If
        Next
    End With
    
    ' Remove invalid WorshipServiceAssistant presentations.
    With mcudtPresentations
        For lng = .Count To 1 Step -1
            If (Not mblnValid(.Item(lng).Doppelganger)) Then
                .Item(lng).Terminate
                .Remove lng
            End If
        Next
    End With
    
    ' Add valid WorshipServiceAssistant presentations.
    lng = 1
    For Each pre In mpreDoppelganger
        If (mblnValid(pre)) Then
            With mcudtPresentations
                If (lng <= .Count) Then
                    If (Not pre Is .Item(lng).Doppelganger) Then
                        .Add New WorshipServiceAssistant.Presentation, Before:=lng
                        .Item(lng).Initialize mudtApplication, mudtParent, pre
                    End If
                Else
                    .Add New WorshipServiceAssistant.Presentation
                    .Item(.Count).Initialize mudtApplication, mudtParent, pre
                End If
            End With
            lng = lng + 1
        End If
    Next
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mblnValid _
( _
    ByRef pre As PowerPoint.Presentation _
) As Boolean
    Dim preRemoved As PowerPoint.Presentation

    mblnValid = (pre.Windows.Count > 0) And (pre.Slides.Count > 0)
    For Each preRemoved In mcpreRemoved
        If (pre Is preRemoved) Then
            mblnValid = False
            Exit For
        End If
    Next
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
