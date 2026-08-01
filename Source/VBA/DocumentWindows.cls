VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.DocumentWindows
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
Private mudtParent As WorshipServiceAssistant.Presentation
Private mdwDoppelganger As PowerPoint.DocumentWindows
Private mcudtWindows As New VBA.Collection


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
    ByRef udtParent As WorshipServiceAssistant.Presentation, _
    ByRef dwDoppelganger As PowerPoint.DocumentWindows _
)
    Dim lng As Long
    
    mblnInvariant = False
    
    Set mudtApplication = udtApplication
    Set mudtParent = udtParent
    Set mdwDoppelganger = dwDoppelganger
    
    With mcudtWindows
        Do While (.Count > 0)
            .Remove .Count
        Loop
    End With
    
    With mdwDoppelganger
        For lng = 1 To .Count Step 1
            mcudtWindows.Add New WorshipServiceAssistant.DocumentWindow
            mcudtWindows.Item(mcudtWindows.Count).Initialize mudtApplication, mudtParent, .Item(lng)
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
Public Sub Terminate _
( _
)
    With mcudtWindows
        Do While (.Count > 0)
            .Item(.Count).Terminate
            .Remove .Count
        Loop
    End With
    
    Set mdwDoppelganger = Nothing
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
            mblnInvariant = blnInvariant
        End If
        With mcudtWindows
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
) As PowerPoint.DocumentWindows
    Set Doppelganger = mdwDoppelganger
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
) As WorshipServiceAssistant.Presentation
    Set Parent = mudtParent
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Presentation _
( _
) As WorshipServiceAssistant.Presentation
    Set Presentation = mudtParent
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
)
    Count = mcudtWindows.Count
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
    ByRef lngIndex As Long _
) As WorshipServiceAssistant.DocumentWindow
Attribute Item.VB_UserMemId = 0
    Set Item = mcudtWindows.Item(lngIndex)
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Active _
( _
) As WorshipServiceAssistant.DocumentWindow
    Dim lng As Long

    Set Active = Nothing
    With mcudtWindows
        For lng = 1 To .Count Step 1
            If (.Item(lng).Active = Office.msoTrue) Then
                Set Active = .Item(lng)
                Exit For
            End If
        Next
    End With
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

