Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Zip_Files
' Level:        Framework module
' Version:      1.01
' Description:  Standard module for compressing files using Windows XP's built-in
'                   'Compressed (Zipped) Folders' feature
' Source/date:  Alan Williams at nps.gov, 7/20/2007 - collected code bits and cleaned
'                   them up for MSAccess
' Revisions:    John R. Boetsch, 1/8/2009 - minor reformatting and name changes
'               JRB, 10/8/2009 - added fxnPause to allow a delay in code execution until
'                   the zip file is created
'               -------------------------------------
'               BLC, 5/26/2015 - 1.00 - included in NCPN invasives reporting tool &
'                   moved sapiSleep & fxnPause(Delay) to mod_Time
'               BLC, 4/4/2016 - 1.01 - changed Exit_Procedure > Exit_Handler
' =================================

' ---------------------------------
'  Functions
' ---------------------------------
Public Declare Function GetVersionExA Lib "kernel32" _
               (LpVersionInformation As OSVERSIONINFO) As Integer

' ---------------------------------
'  Types
' ---------------------------------
Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

' ---------------------------------
'  Constants
' ---------------------------------
Public Const VER_PLATFORM_WIN32s = 0        ' Win32s on Windows 3.1
Public Const VER_PLATFORM_WIN32_WINDOWS = 1 ' Windows 95, Windows 98, or Windows Me
Public Const VER_PLATFORM_WIN32_NT = 2      ' Windows NT/2000/XP, or Windows Server 2003 family

' ---------------------------------
'  Functions
' ---------------------------------

' =================================
' FUNCTION:     ZipFiles
' Description:  Creates a Zip file using WinXP Compressed (Zipped) Folders (Won't work on Win2K)
' Parameters:   varSourceFiles = Individual file path or a Directory path
'               varZipFileName = Name of Zip file to create or add file(s) to
'               Optional AppendToZip = True, If you want to append to an exsisiting zip file,
'                   False or missing, to create a new file or kill existing file
' Returns:      False if error occurs
' Throws:       none
' References:   fxnNewZip (Make empty Zip File)
'               fxnGetVersion (Checks for to see if user is running atleast WinXP)
' KnownIssues:  It is not possible to hide the copy dialog while copying to a zip folder.
'               Also there is no way to avoid that someone can cancel the CopyHere operation
'               or that your VBA code will be notified that the operation has been cancelled.
'            ** Also the Shell.NameSpace().CopyHere operation fails if the paths are passed
'               as a strings. Converting path to variants just before NameSpace().CopyHere
'               seems to keep it happy.
' Source/date:  Alan Williams, 7/19/2007 - help from http://www.tek-tips.com/faqs.cfm?fid=4599
' Revisions:    Alan Williams, 7/20/2007 - added OS check to help evaluate compatability
'               John R. Boetsch, 1/8/2009 - updated error handling and naming conventions
'               BLC, 5/19/2015 - renamed, removed fxn prefix
' =================================
Public Function ZipFiles(strSourceFiles As String, strZipFileName As String, _
    Optional AppendToZip As Boolean = False) As Boolean
    On Error GoTo Err_Handler

    ' Default in case of fail
    ZipFiles = False

    Dim OSVersion
    Dim objShell As Object
    Dim strProcName As String
    strProcName = "ZipFiles"

    'Check that the user has atleast WinXP
    Call GetVersion(OSVersion)
    If OSVersion < 5.1 Then
        MsgBox "This function is not supported by " & strProcName & _
        " with an OS version of: " & OSVersion, vbCritical, strProcName & " Error"
        ZipFiles = False
        GoTo Exit_Handler
    End If

    If (AppendToZip = False) Or (Len(Dir(strZipFileName)) = 0) Then
        Call NewZip(strZipFileName)
    End If

    'Check Source Files
    If (Len(Dir(strSourceFiles)) = 0) Then
        MsgBox "The source file(s): " & vbNewLine & strSourceFiles & vbNewLine & _
        "Are missing...", vbCritical, strProcName & " Error"
        ZipFiles = False
        GoTo Exit_Handler
    End If
    
    'Copy files
    Set objShell = CreateObject("Shell.Application")
    objShell.NameSpace(CVar(Trim(strZipFileName))).CopyHere CVar(Trim(strSourceFiles))
    ZipFiles = True
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ZipFiles[mod_Zip])"
    End Select
    Resume Exit_Handler

End Function

' =================================
' FUNCTION:     NewZip
' Description:  Makes new, empty zip file (kills any file with the same name before)
' Parameters:   sPath = Full directory path and file name for new zip file
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  http://www.tek-tips.com/faqs.cfm?fid=4599 (Changed by keepITcool Dec-12-2005)
' Revisions:    Alan Williams, 7/19/2007 - added error handling
'               John R. Boetsch, 1/8/2009 - updated error handling and formatting
'               BLC, 5/19/2015 - renamed, removed fxn prefix
' =================================
Public Function NewZip(sPath)
    On Error GoTo Err_Handler

    Dim strProcName As String
    strProcName = "NewZip"

    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewZip[mod_Zip_Files])"
    End Select
    Resume Exit_Handler

End Function

' =================================
' FUNCTION:     GetVersion
' Description:  Extracts OS Version info and returns it
' Parameters:   Optional VersionNumber
' Returns:      operating system version information (string)
' Throws:       none
' References:   none
' Source/date:  http://www.tek-tips.com/faqs.cfm?fid=4599
' Revisions:    Alan Williams, 7/19/2007 - added optional parameter 'VersionNumber'
'                   to help evaluate compatability
'               John R. Boetsch, 1/8/2009 - added error handling and updated formatting
'               BLC, 5/19/2015 - renamed, removed fxn prefix
' =================================
Private Function GetVersion(Optional VersionNumber) As String
    On Error GoTo Err_Handler

    Dim OSInfo As OSVERSIONINFO
    Dim retvalue As Integer

    OSInfo.dwOSVersionInfoSize = 148
    OSInfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(OSInfo)

    With OSInfo
        Select Case .dwPlatformId
          Case VER_PLATFORM_WIN32s           ' Win32s on Windows 3.1
            GetVersion = "Windows 3.1"

          Case VER_PLATFORM_WIN32_WINDOWS    ' Windows 95, Windows 98,
            Select Case .dwMinorVersion   ' or Windows Me
              Case 0
                GetVersion = "Windows 95"
              Case 10
                If (OSInfo.dwBuildNumber And &HFFFF&) = 2222 Then
                    GetVersion = "Windows 98SE"
                Else
                    GetVersion = "Windows 98"
                End If
              Case 90
                GetVersion = "Windows Me"
            End Select

          Case VER_PLATFORM_WIN32_NT         ' Windows NT, Windows 2000, Windows XP,
            Select Case .dwMajorVersion   ' or Windows Server 2003 family.
              Case 3
                GetVersion = "Windows NT 3.51"
              Case 4
                GetVersion = "Windows NT 4.0"
              Case 5
                Select Case .dwMinorVersion
                  Case 0
                    GetVersion = "Windows 2000"
                  Case 1
                    GetVersion = "Windows XP"
                  Case 2
                    GetVersion = "Windows Server 2003"
                  End Select
            End Select

          Case Else
            GetVersion = "Failed"
        End Select
        If Not IsMissing(VersionNumber) Then
            VersionNumber = CSng(.dwMajorVersion & "." & .dwMinorVersion)
            GetVersion = GetVersion & " (" & VersionNumber & ")"
        End If
    End With

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetVersion[mod_Zip_Files])"
    End Select
    Resume Exit_Handler

End Function