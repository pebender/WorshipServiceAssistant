Attribute VB_Name = "Project"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Project
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
Option Private Module
Option Explicit
Option Compare Text
Option Base 0


'===============================================================================
' Public Constants.
'===============================================================================
Public Const ProjectName As String = "WorshipServiceAssistant"
Public Const ProjectNamePretty As String = "Worship Service Assistant"
Public Const ProjectVersion As String = "1.00.0000"
Public Const ProjectAuthor As String = "Paul Bender"
Public Const ProjectCopyright As String = "Copyright (c) 2000, 2001 Paul Bender"
Public Const ProjectHomepage As String = "http://home.san.rr.com/benderfamily/software/wsa/"
Public Const ProjectEmail As String = "mailto:pebender@san.rr.com"


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
