Attribute VB_Name = "Description"
'===============================================================================
' Naming Conventations
'===============================================================================
' Prefixes for constants, variables and routines:
'   Scope:
'     'g': Global
'     'm': Module
'     '' : Local
'   Size:
'     '' : Scalar
'     'a': Array
'   Type:
'     ''   : Routine with no return value.
'     'int': Integer.
'     'lng': Long.
'     'str': String.
'     'vnt': Variant.
'     'obj': Object.
'     'mod': Module.
'     'frm': Form.
'     'tab': Form multipage or Form page.
'     'fra': Form frame.
'     'lbl': Form label.
'     'lst': Form list box.
'     'tlb': Office command bar (aka a toolbar).
'     'mnu': Office command bar menu.
'     'cmd': Form command button or
'            Office command bar button.
'     'ctr': Office command bar control
'     'cbo': Office command bar combo box.
'     'pp' : PowerPoint type.
'     'pre': PowerPoint presentation.
'     'dw' : PowerPoint document window.
'     'ssw': PowerPoint slide show window.
'     'sld': PowerPoint slide.
'     'shp': PowerPoint shape.
'     'sel': PowerPoint selection.
'
' The project does not require that names of globals be globally unique.
' Therefore, when a global is referenced, the referece includes both the
' name of the global and the name of parent of the global.
'===============================================================================

'===============================================================================
' WSA Object Model
'===============================================================================
' WSA                                   WSAApplication
'   Banner                              WSABanner
'   Presentations                       WSAPresentationCollection
'     Presentation                      WSAPresentation
'       Windows                         WSAWindowCollection
'         Window                        WSAWindow
'       Slides                          WSASlideCollection
'         Slide                         WSASlide
'     SlideShowWindow                   WSASlideShowWindow
'===============================================================================

