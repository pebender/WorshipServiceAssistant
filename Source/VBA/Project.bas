Attribute VB_Name = "Project"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Project
'
' Description:
'
' Author:
'   Paul Bender <pbender@alumni.ucsd.edu>
'
' Copyright:
'   Copyright (c) 2000,2001,2002 Paul Bender
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
Public Const ProjectName As String = "WorshipServiceAssistant"
Public Const ProjectNamePretty As String = "Worship Service Assistant"
Public Const ProjectVersion As String = "1.02.0000"
Public Const ProjectAuthor As String = "Paul Bender"
Public Const ProjectCopyright As String = "Copyright (c) 2000,2001,2002 Paul Bender"
Public Const ProjectHomepage As String = "http://home.san.rr.com/benderfamily/software/wsa/"
Public Const ProjectEmail As String = "mailto:pbender@alumni.ucsd.edu"


'===============================================================================
' Public Variables.
'===============================================================================

'
' This variable is true if the project is running.
'
Public Project_Loaded As Boolean

'
' This variable contains the list of categories in the 'Set Category' menu.
'
Public Project_Categories() As String

'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private AEH As New ApplicationEventHandler


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'   Load the project.
'-------------------------------------------------------------------------------
Public Sub Project_Load()
    Project_Loaded = True
    Menu_Load
    SlideShow_Initialize
    
    Set AEH.PPTApplication = Application
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Unload Project.
'-------------------------------------------------------------------------------
Public Sub Project_Unload()
    Menu_Unload
    Project_Loaded = False
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
