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
Private Sub PPTApplication_NewPresentation _
( _
    ByVal P As PowerPoint.Presentation _
)
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_PresentationClose _
( _
    ByVal P As PowerPoint.Presentation _
)
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_PresentationNewSlide _
( _
    ByVal S As PowerPoint.Slide _
)
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_PresentationOpen _
( _
    ByVal P As PowerPoint.Presentation _
)
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_PresentationPrint _
( _
    ByVal P As PowerPoint.Presentation _
)
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_PresentationSave _
( _
    ByVal P As PowerPoint.Presentation _
)
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_SlideShowBegin _
( _
    ByVal SSW As PowerPoint.SlideShowWindow _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_SlideShowEnd _
( _
    ByVal P As PowerPoint.Presentation _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_SlideShowNextBuild _
( _
    ByVal SSW As PowerPoint.SlideShowWindow _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_SlideShowNextSlide _
( _
    ByVal SSW As PowerPoint.SlideShowWindow _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_WindowActivate _
( _
    ByVal P As PowerPoint.Presentation, _
    ByVal W As PowerPoint.DocumentWindow _
)
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_WindowBeforeDoubleClick _
( _
    ByVal Sel As PowerPoint.Selection, _
    Cancel As Boolean _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_WindowBeforeRightClick _
( _
    ByVal Sel As PowerPoint.Selection, _
    Cancel As Boolean _
)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_WindowDeactivate _
( _
    ByVal P As PowerPoint.Presentation, _
    ByVal W As PowerPoint.DocumentWindow _
)
    Menu_Disable
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PPTApplication_WindowSelectionChange _
( _
    ByVal Sel As PowerPoint.Selection _
)
    Menu_Refresh
End Sub

