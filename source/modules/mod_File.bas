Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_File
' Level:        Framework module
' Version:      1.03
' Description:  File and directory related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 6/13/2016 - 1.01 - adapted ParseFileName() for big rivers
'               BLC, 6/24/2016 - 1.02 - replaced Exit_Function > Exit_Handler
'               BLC, 8/30/2016 - 1.03 - add BrowseFolder(), GetSpecialFolderPath()
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
'   Peter Thornton, March 4, 2009
'   http://dailydoseofexcel.com/archives/2009/02/26/get-the-path-to-my-documents-in-vba/#comment-38217
'   https://msdn.microsoft.com/en-us/library/windows/desktop/bb762494%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
'   Public Const CSIDL_PERSONAL As Long = &H5 'my documents
'   was CSIDL constants, now KNOWNFOLDERIDs
'   https://msdn.microsoft.com/en-us/library/windows/desktop/dd378457(v=vs.85).aspx
'
'   KNOWNFOLDERIDs
'       FOLDERID_AccountPictures        FOLDERID_AppsFolder     FOLDERID_ChangeRemovePrograms
'       FOLDERID_AddNewPrograms         FOLDERID_AppUpdates     FOLDERID_CommonAdminTools
'       FOLDERID_AdminTools             FOLDERID_CameraRoll     FOLDERID_CommonOEMLinks
'       FOLDERID_ApplicationShortcuts   FOLDERID_CDBurning      FOLDERID_CommonPrograms
'       FOLDERID_CommonStartMenu        FOLDERID_ComputerFolder FOLDERID_CommonTemplates
'       FOLDERID_CommonStartup          FOLDERID_ConflictFolder FOLDERID_ConnectionsFolder
'       FOLDERID_Contacts               FOLDERID_Cookies        FOLDERID_ControlPanelFolder
'       FOLDERID_Desktop                FOLDERID_Documents      FOLDERID_DeviceMetadataStore
'       FOLDERID_DocumentsLibrary       FOLDERID_Downloads      FOLDERID_Favorites
'       FOLDERID_Fonts                  FOLDERID_Games          FOLDERID_GameTasks
'       FOLDERID_History                FOLDERID_HomeGroup      FOLDERID_HomeGroupCurrentUser
'       FOLDERID_ImplicitAppShortcuts   FOLDERID_InternetCache  FOLDERID_InternetFolder
'       FOLDERID_Libraries              FOLDERID_Links          FOLDERID_LocalAppData
'       FOLDERID_LocalAppDataLow        FOLDERID_Music          FOLDERID_LocalizedResourcesDir
'       FOLDERID_MusicLibrary           FOLDERID_NetHood        FOLDERID_NetworkFolder
'       FOLDERID_OriginalImages         FOLDERID_PhotoAlbums    FOLDERID_PicturesLibrary
'       FOLDERID_Pictures               FOLDERID_Playlists      FOLDERID_PrintersFolder
'       FOLDERID_PrintHood              FOLDERID_Profile        FOLDERID_ProgramData
'       FOLDERID_ProgramFiles           FOLDERID_Programs       FOLDERID_PublicDocuments
'       FOLDERID_ProgramFilesX86        FOLDERID_Public         FOLDERID_PublicDownloads
'       FOLDERID_ProgramFilesX64        FOLDERID_PublicDesktop  FOLDERID_PublicGameTasks
'       FOLDERID_ProgramFilesCommon     FOLDERID_PublicMusic    FOLDERID_PublicLibraries
'       FOLDERID_ProgramFilesCommonX64  FOLDERID_PublicPictures FOLDERID_PublicRingtones
'       FOLDERID_ProgramFilesCommonX86  FOLDERID_PublicVideos   FOLDERID_PublicUserTiles
'       FOLDERID_QuickLaunch            FOLDERID_Recent         FOLDERID_RecordedTV
'       FOLDERID_RecordedTVLibrary      FOLDERID_ResourceDir    FOLDERID_RecycleBinFolder
'       FOLDERID_Ringtones              FOLDERID_RoamingAppData FOLDERID_RoamedTileImages
'       FOLDERID_RoamingTiles           FOLDERID_SampleMusic    FOLDERID_SamplePictures
'       FOLDERID_SamplePlaylists        FOLDERID_SampleVideos   FOLDERID_SavedGames
'       FOLDERID_SavedPictures          FOLDERID_SavedSearches  FOLDERID_SavedPicturesLibrary
'       FOLDERID_Screenshots            FOLDERID_SearchCSC      FOLDERID_SearchHistory
'       FOLDERID_SearchHome             FOLDERID_SEARCH_MAPI    FOLDERID_SearchTemplates
'       FOLDERID_SendTo                 FOLDERID_SidebarParts   FOLDERID_SidebarDefaultParts
'       FOLDERID_SkyDrivePictures       FOLDERID_SkyDrive       FOLDERID_SkyDriveDocuments
'       FOLDERID_SkyDriveCameraRoll     FOLDERID_StartMenu      FOLDERID_SyncManagerFolder
'       FOLDERID_SyncResultsFolder      FOLDERID_Startup        FOLDERID_SyncSetupFolder
'       FOLDERID_System                 FOLDERID_SystemX86      FOLDERID_Templates
'       FOLDERID_TreeProperties         FOLDERID_UserPinned     FOLDERID_UserProfiles
'       FOLDERID_UserProgramFiles       FOLDERID_UsersFiles     FOLDERID_Videos
'       FOLDERID_UserProgramFilesCommon FOLDERID_UsersLibraries FOLDERID_VideosLibrary
'       FOLDERID_Windows
'
'   WshScript Special Folders
'        AllUsersDesktop        Desktop         NetHood        SendTo
'        AllUsersStartMenu      Favorites       PrintHood      StartMenu
'        AllUsersPrograms       Fonts           Programs       Startup
'        AllUsersStartup        MyDocuments     Recent         Templates


' ---------------------------------
'  DIRECTORY RELATED
' ---------------------------------
' =================================
' FUNCTION:     CreateFolder
' Description:  Creates a folder with the specified path
' Parameters:   strPath as a string
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 1/9/2009
' Revisions:    JRB, 1/9/2009 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function CreateFolder(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    CreateFolder = False    ' Default in case of error

    Dim fs As Variant

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(strPath) = False Then
        fs.CreateFolder (strPath)
        CreateFolder = True
    End If

Exit_Handler:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateFolder[mod_File])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     FolderExists
' Description:  Indicates whether or not the indicated folder exists
' Parameters:   strPath as a string
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 1/9/2009
' Revisions:    JRB, 1/9/2009 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function FolderExists(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    FolderExists = False    ' Default in case of error

    Dim fs As Variant

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(strPath) Then FolderExists = True

Exit_Handler:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FolderExists[mod_File])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     BrowseFolder
' Description:  file dialog browsing actions
' Assumptions:  -
' Parameters:   Title - display name of the file dialog (string)
'               ButtonTitle - name of OK button (string)
'               InitialFolder - folder to begin display (string)
'               InitialView - desired file dialog view (string, MsoFileDialogView options)
'               DialogType - desired dialog type (string, MsoFileDialogType options)
'               AllowMultiples - allows multiple directories to be selected (boolean)
' Returns:      fully-qualified folder name selected by the user
'               or an empty string if the user cancelled the dialog (string)
' Throws:       none
' References:
'   Chip Pearson, July 5, 2007
'   http://www.cpearson.com/excel/browsefolder.aspx
' Source/date:
' Adapted:      Bonnie Campbell, August 30, 2016 - for NCPN tools
' Revisions:
'   BLC - 8/30/2016 - initial version
' ---------------------------------
Public Function BrowseFolder(Title As String, _
        Optional ButtonTitle As String = "Confirm", _
        Optional InitialFolder As String = vbNullString, _
        Optional InitialView As Office.MsoFileDialogView = msoFileDialogViewList, _
        Optional DialogType As MsoFileDialogType = msoFileDialogFolderPicker, _
        Optional AllowMultiples As Boolean = False) As String
'----------------------
' Dialog Options:
'   (MsoFileDialogType Constants)
'   msoFileDialogFilePicker     Allows user to select a file.
'   msoFileDialogFolderPicker   Allows user to select a folder.
'   msoFileDialogOpen           Allows user to open a file.
'   msoFileDialogSaveAs         Allows user to save a file.
'
' View Options:
'   msoFileDialogViewDetails    2   Files displayed in a list with detail information.
'   msoFileDialogViewLargeIcons 6   Files displayed as large icons.
'   msoFileDialogViewList       1   Files displayed in a list without details.
'   msoFileDialogViewPreview    4   Files displayed in a list with a preview pane showing
'                                   the selected file.
'   msoFileDialogViewProperties 3   Files displayed in a list with a pane showing the
'                                   selected file's properties.
'   msoFileDialogViewSmallIcons 7   Files displayed as small icons.
'   msoFileDialogViewThumbnail  5   Files displayed as thumbnails.
'   msoFileDialogViewTiles      9   Files displayed as tiled icons.
'   msoFileDialogViewWebView    8   Files displayed in Web view.
'----------------------
    
On Error GoTo Err_Handler
    
    'Dim V As Variant
    Dim InitFolder As String
    Dim SelectedFolder As String
    
    'prepare file dialog box
    With Application.FileDialog(DialogType)
        .Title = Title
        .ButtonName = ButtonTitle
        .InitialView = InitialView

        If Len(InitialFolder) > 0 Then
            
            If dir(InitialFolder, vbDirectory) <> vbNullString Then
                InitFolder = InitialFolder
                If Right(InitFolder, 1) <> "\" Then
                    InitFolder = InitFolder & "\"
                End If
                .InitialFileName = InitFolder
            End If
            
        End If
        '.Show
        
        'set directory if OK clicked
        If .Show = True Then
            SelectedFolder = .SelectedItems(1)
        End If
        
'        On Error Resume Next
'        Err.Clear
'
'        V = .SelectedItems(1)
'        If Err.Number <> 0 Then
'            V = vbNullString
'        End If
    End With
    
    BrowseFolder = SelectedFolder 'CStr(V)
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - BrowseFolder[mod_File])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetSpecialFolderPath
' Description:  retrieve full path of specified special folder
' Assumptions:  -
' Notes:
'   WshScript special folders include
'        AllUsersDesktop        Desktop         NetHood        SendTo
'        AllUsersStartMenu      Favorites       PrintHood      StartMenu
'        AllUsersPrograms       Fonts           Programs       Startup
'        AllUsersStartup        MyDocuments     Recent         Templates
'
' Parameters:   SpecialFolder - special folder name (string)
' Returns:      fully-qualified folder name or desired special folder (string)
' Throws:       none
' References:
'   Mike Alexander, February 27, 2009
'   http://dailydoseofexcel.com/archives/2009/02/26/get-the-path-to-my-documents-in-vba/
'   bradxlsure, March 17, 2008
'   http://www.pcreview.co.uk/threads/re-wscript-object-not-found.947405/
' Source/date:
' Adapted:      Bonnie Campbell, August 30, 2016 - for NCPN tools
' Revisions:
'   BLC - 8/30/2016 - initial version
' ---------------------------------
Function GetSpecialFolderPath(SpecialFolder As String)
On Error GoTo Err_Handler

    Dim arySpecials() As String, strPath As String
    
    arySpecials = Split("desktop,allusersdesktop,sendto,startmenu,recent,favorites,mydocuments" _
                    & "" _
                        , ",")
    
    
    'default
    strPath = ""
        
 '   If IsInArray(SpecialFolder, arySpecials) Then
    
        Dim oWshShell As Object
        Dim oFolders As Object
        
        Set oWshShell = CreateObject("WScript.Shell")
        
        Set oFolders = oWshShell.SpecialFolders
    

        strPath = oFolders(SpecialFolder)
    
  '  End If
    
    GetSpecialFolderPath = strPath

Exit_Handler:
    Set oWshShell = Nothing
    Set oFolders = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSpecialFolderPath[mod_File])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  FILE RELATED
' ---------------------------------

' =================================
' FUNCTION:     GetFile
' Description:  Opens the open/save file dialog and returns the file name selected by the user
' Parameters:   strInitialDir - the directory to start searching in (optional)
'               strFileType, varFileExt - file type and extension (optional)
'               strTitle - title of the dialog box (optional)
' Returns:      name of the file to open/import; or Null if user cancels
' Throws:       none
' References:   adhAddFilterItem, adhCommonFileOpenSave
' Source/date:  Susan Huse, fall 2004
' Revisions:    John R. Boetsch, May 17, 2006 - updated documentation and error trap
'               JRB, 6/22/2009 - revised from fxnGetLinkFile; added file type/ext variables
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function GetFile(Optional ByVal strInitialDir As String, _
    Optional ByVal strFileType As String, _
    Optional ByVal varFileExt As Variant, _
    Optional ByVal strTitle As String = "Select File to Open") As Variant

    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim lngFlags As Long

    ' Use the open file dialog to interactively browse to and select the desired file
    strFilter = adhAddFilterItem(strFilter, strFileType, varFileExt)

    lngFlags = adhOFN_HIDEREADONLY Or _
        adhOFN_HIDEREADONLY Or adhOFN_NOCHANGEDIR

    GetFile = adhCommonFileOpenSave( _
        InitialDir:=strInitialDir, _
        OpenFile:=True, _
        Filter:=strFilter, _
        flags:=lngFlags, _
        DialogTitle:=strTitle)

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetFile[mod_File])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     SaveFile
' Description:  Opens the open/save file dialog and returns the file name selected by the user
' Parameters:   strFileName, strFileType, strFileExt - file name/path, type and extension
'               strTitle - title of the dialog box (optional)
' Returns:      name of the file to save; or Null if user cancels
' Throws:       none
' References:   adhAddFilterItem, adhCommonFileOpenSave
' Source/date:  Susan Huse, fall 2004
' Revisions:    John R. Boetsch, May 2005 - minor revisions and documentation
' Revisions:    JRB, 5/16/2006 - updated documentation, error traps
'               JRB, 6/22/2009 - added strTitle to parameters
'               BLC, 4/30/2015 - move from mod_Utilities to mod_File
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function SaveFile(ByVal strFilename As String, ByVal strFileType As String, _
    ByVal strFileExt As String, Optional ByVal strTitle As String = "Save As") As Variant

    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim lngFlags As Long

    ' Use the save file dialog to interactively browse to and select the desired file
    strFilter = adhAddFilterItem(strFilter, strFileType, strFileExt)

    lngFlags = adhOFN_HIDEREADONLY Or adhOFN_OVERWRITEPROMPT Or _
        adhOFN_HIDEREADONLY Or adhOFN_NOCHANGEDIR

    SaveFile = adhCommonFileOpenSave( _
        OpenFile:=False, _
        Filter:=strFilter, _
        flags:=lngFlags, _
        DialogTitle:=strTitle, _
        FileName:=strFilename)

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SaveFile[mod_File])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     FileExists
' Description:  Indicates whether or not the indicated file exists
' Parameters:   strPath as a string
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 5/8/2006
' Revisions:    JRB, 5/8/2006 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function FileExists(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    FileExists = False    ' Default in case of error

    Dim fs As Variant

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(strPath) Then FileExists = True

Exit_Handler:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FileExists[mod_File])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     DeleteFile
' Description:  Deletes the specified file; this is preferred over the Kill command
'               because it works for hidden files and read-only files
' Parameters:   strPath - the path and file name to be deleted
' Returns:      True if deleted, or False if error
' Throws:       none
' References:   FileExists
' Source/date:  John R. Boetsch, 5/19/2006
' Revisions:    JRB, 5/19/2006 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function DeleteFile(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    DeleteFile = False    ' Default in case of error

    Dim fs As Variant

    Set fs = CreateObject("Scripting.FileSystemObject")
    If FileExists(strPath) Then
        fs.DeleteFile strPath, True
        DeleteFile = True
    Else
        MsgBox "Unable to delete the specified file", vbCritical, _
            "File delete error (DeleteFile)"
    End If

Exit_Handler:
    On Error Resume Next
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DeleteFile[mod_File])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     ParseFileName
' Description:  Parses an input path string to return only the name, if present
' Parameters:   strFullPath - string for the full file path
' Returns:      string including only the file name
' Throws:       none
' References:   none
' Source/date:  From Front-end Application Builder v1.1, Simon Kingston, date unknown
' Revisions:    John R. Boetsch, 6/17/2009 - error trapping, documentation
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 6/13/2016 - adapted for big rivers
' =================================
Public Function ParseFileName(ByVal strFullPath As String) As String
    On Error GoTo Err_Handler

    Dim strTemp As String

    Do While (InStr(strFullPath, "\") > 0)
        strTemp = strTemp & Left(strFullPath, InStr(strFullPath, "\"))
        strFullPath = Mid(strFullPath, InStr(strFullPath, "\") + 1)
    Loop
    
    ParseFileName = strFullPath

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ParseFileName[mod_File])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     ParseFileExt
' Description:  Parses an input path string to return only the file extension, if present
' Parameters:   strFullPath - string for the full file path
'               blnIncludeDot - flag to include the dot (".") in the return (default is True)
' Returns:      string including only the file extension, or an empty string ("") if missing
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/22/2009
' Revisions:    JRB, 6/22/2009 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function ParseFileExt(ByVal strFullPath As String, _
    Optional blnIncludeDot As Boolean = True) As String

    On Error GoTo Exit_Handler

    Dim arrPath() As String
    Dim strFile As String
    Dim strTemp As String
    Dim varPosition As Variant

    ' Split into an array based on the "\" delimiter; file name should be the uppermost segment
    arrPath = Split(strFullPath, "\")
    strFile = arrPath(UBound(arrPath))

    ' Get the position in the string of the dot
    varPosition = InStr(1, strFile, ".")
    If varPosition > 0 Then
        If blnIncludeDot = False Then varPosition = varPosition + 1
        strTemp = Mid(strFile, varPosition)
    Else
        strTemp = ""
    End If

    ParseFileExt = strTemp

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ParseFileExt[mod_File])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     OpenExcelFile
' Description:  Opens file in Excel - assumes that the file exists and can be opened by Excel
' Parameters:   strPath - full path of the file to be opened
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/22/2009
' Revisions:    JRB, 3/7/12 - fixed function header to indicate 'Public'
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function OpenExcelFile(ByVal strPath As String)
    On Error GoTo Err_Handler

    Dim objExcel As Object

    ' Create a new instance of Excel
    Set objExcel = CreateObject("Excel.Application")
    objExcel.UserControl = True

    ' Open the file
    With objExcel
        .visible = True
        .Workbooks.Open (strPath)
    End With
    
Exit_Handler:
    On Error Resume Next
    Set objExcel = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - OpenExcelFile[mod_File])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     ParsePath
' Description:  Parses an input path string to return only the path without the file name
' Parameters:   strFullPath - string for the full file path
' Returns:      string including only the file path, or an empty string ("") if missing
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/22/2009
' Revisions:    JRB, 6/22/2009 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_File
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function ParsePath(ByVal strFullPath As String) As String
    On Error GoTo Exit_Handler

    Dim arrPath() As String
    Dim strFile As String

    ' Split into an array based on the "\" delimiter; file name should be the uppermost segment
    arrPath = Split(strFullPath, "\")
    strFile = arrPath(UBound(arrPath))

    ' Path is the full string minus length of the file name
    ParsePath = Left(strFullPath, Len(strFullPath) - Len(arrPath(UBound(arrPath))))

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ParsePath[mod_File])"
    End Select
    Resume Exit_Handler
End Function