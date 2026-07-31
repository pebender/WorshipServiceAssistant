Attribute VB_Name = "Navigator"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Navigator
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
'
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
'   Launch the slide navigator.  In order to exit, the slide navigator must be
'   unloaded.  If the slide navigator is only hidden, then this routine will
'   automatically refresh it and re-show it.  This is a hack to work around
'   some focus problems resulting from activating slide shows and activating
'   new presentations.
'-------------------------------------------------------------------------------
Public Sub Navigator_Run()
    '
    ' Show the Navigator form until it is unloaded.
    '
    Do
        If (Navigator_Loaded = True) Then
            NavigatorForm.Refresh
        End If
        NavigatorForm.show
    Loop Until (Navigator_Loaded = False)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function Navigator_Refresh() As Boolean
    If (Navigator_Loaded = True) Then
        NavigatorForm.Refresh
    End If
End Function

'-------------------------------------------------------------------------------
' Description:
'   Determine if the slide navigator is loaded.
'-------------------------------------------------------------------------------
Public Function Navigator_Loaded() As Boolean
    Dim f As Object
    
    Navigator_Loaded = False
    For Each f In UserForms
        If (f.Name = "NavigatorForm") Then
            Navigator_Loaded = True
            Exit For
        End If
    Next
End Function


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
