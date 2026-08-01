Attribute VB_Name = "ToolbarAction"
'===============================================================================
' Name:
'   WorshipServiceAssistant.ToolbarAction
'
' Description:
'   The routines in this module provide a wrapper around each of the Worship
'   Service Assistant toolbar OnAction routines. Each wrapper checks whether
'   or not Worship Service Assistant has crashed and restarts Worship Service
'   Assistant when necessary.
'
' Author:
'   Paul Bender <pbender@alumni.ucsd.edu>
'
' Copyright:
'   Copyright (c) 2002,2003 Paul Bender
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
'   2.01.0000:
'     (1) Add support for "Set Template" menu item in "Song Edit" menu.
'   2.00.0000:
'     (1) Reset change history.
'===============================================================================


'===============================================================================
' Options.
'===============================================================================
Option Private Module
Option Explicit
Option Compare Text
Option Base 0


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' mNavigator*_KeyDown subroutines.
'
' Purpose:
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Public Sub ModeButton _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_ModeValue_Action
End Sub

Public Sub Navigator _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_Navigator_Action
End Sub

Public Sub SongEditSetCategoryButton _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_SongEditSetCategoryValue_Action
End Sub

Public Sub SongEditSetTemplate _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_SongEditSetTemplate_Action
End Sub

Public Sub SongEditSort _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_SongEditSort_Action
End Sub

Public Sub SongEditIndex _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_SongEditIndex_Action
End Sub

Public Sub DebugSSWMonitorButton _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_DebugSSWMonitorValue_Action
End Sub

Public Sub DebugSSWSizeButton _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_DebugSSWSizeValue_Action
End Sub

Public Sub HelpTopicApplication _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_HelpTopicApplication_Action
End Sub

Public Sub HelpTopicHowTo _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_HelpTopicHowTo_Action
End Sub

Public Sub HelpTopicToolbar _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_HelpTopicToolbar_Action
End Sub

Public Sub HelpTopicKnownIssues _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_HelpTopicKnownIssues_Action
End Sub

Public Sub HelpTopicCopyrightPermission _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_HelpTopicCopyrightPermission_Action
End Sub

Public Sub HelpVisitHomepage _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_HelpVisitHomepage_Action
End Sub

Public Sub HelpEmailAuthor _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_HelpEmailAuthor_Action
End Sub

Public Sub HelpAbout _
( _
)
    If (WorshipServiceAssistant.Application Is Nothing) Then
        WorshipServiceAssistant.Auto_Open
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Toolbar.Control_HelpAbout_Action
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
