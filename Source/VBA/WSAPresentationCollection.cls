VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSAPresentationCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.WSAPresentationCollection
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
Private mudtCollection As New WorshipServiceAssistant.WSACollection


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
Public Property Get Count _
( _
) As Long
    Count = mudtCollection.Count
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
) As WorshipServiceAssistant.WSAPresentation
    Set Item = mudtCollection.Item(lngIndex)
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function Index _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
) As Long
    Index = mudtCollection.Index(udtPresentation)
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
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    mudtCollection.Add udtPresentation
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
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    mudtCollection.Remove udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Initialize _
( _
)
    mudtCollection.Initialize
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
    mudtCollection.Terminate
End Sub


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
