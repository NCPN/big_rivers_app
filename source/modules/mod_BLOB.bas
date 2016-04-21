Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_BLOB
' Level:        Framework module
' Version:      1.00
' Description:  BLOB functions & procedures
'
' Source/date:  Bonnie Campbell, 11/25/2015
' Revisions:    BLC, 11/25/2015 - 1.00 - initial version
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  Methods
' ---------------------------------
' ---------------------------------
' SUB:          CreateAppTempImages
' Description:  Create the temporary folder & files for the application
' Assumptions:  Folder & files will remain as long as the user doesn't delete them
'
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Requires:     -
' Source/date:
'   Michael Carey, December 21, 2006
'   http://www.codeproject.com/Articles/16851/Uploading-and-Downloading-BLOBs-to-Microsoft-Acces
' Adapted:      Bonnie Campbell, November 25, 2015 - for NCPN tools
' Revisions:
'   BLC - 11/25/2015  - initial version
' ---------------------------------
Public Sub CreateAppTempImages()
On Error GoTo Err_Handler
    
    'determine if application images folder exists
    If FolderExists(APP_IMAGES_DIR) Then
    
    
        'if app images folder exists, check for images/icons
    
    Else
    'if no images folder, create it & add images
    
        If CreateFolder(APP_IMAGES_DIR) Then
    
        Else
            'exit
            MsgBox ""
        End If
    End If
    
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateAppTempImages[mod_BLOB])"
    End Select
    Resume Exit_Handler
End Sub