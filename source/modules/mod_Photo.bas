Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Photo
' Level:        Framework module
' Version:      1.00
' Description:  photo functions & procedures
'
' Source/date:  Bonnie Campbell, 8/30/2016
' Revisions:    BLC, 8/30/2016 - 1.00 - initial version
' =================================

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  Subroutines & Functions
' ---------------------------------

' ---------------------------------
' SUB:          IngestPhotos
' Description:  photo ingestion actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
'   SJ, December 7, 2014
'   http://excel-macro.tutorialhorizon.com/excel-vba-insert-multiple-images-from-a-folder-to-excel-cells/
' Source/date:  Bonnie Campbell, August 30, 2016 - for NCPN tools
' Adapted:  -
' Revisions:
'   BLC - 8/30/2015 - initial version
' ---------------------------------
Public Sub IngestPhotos(strPath As String, category As String)
On Error GoTo Err_Handler

    Dim fso As FileSystemObject
    Dim iFile As File
    Dim NumFiles As Integer, i As Integer, iProg As Integer
    Dim ListFiles As Files
    Dim aryExtensions() As String
    Dim strFullPath As String, strProgForm As String
    Dim varReturn As Variant
    Dim frm As Form

    aryExtensions = Split(PHOTO_EXT_ALLOWED, ",")

    'exit if no path given
    If Len(strPath) = 0 Then GoTo Exit_Handler

    'determine if directory exists
    If DirExists(strPath) Then
     
        Set fso = CreateObject("Scripting.FileSystemObject")
    
        NumFiles = fso.GetFolder(strPath).Files.Count
        
        'retrieve files
        Set ListFiles = fso.GetFolder(strPath).Files
        
        'present system progress bar
        varReturn = SysCmd(acSysCmdInitMeter, "Uploading photos", NumFiles)
        iProg = 0
        
        'present hourglass
        DoCmd.Hourglass True
        
        'present custom progress form
        strProgForm = "ProgressMeter"
        DoCmd.OpenForm strProgForm
        Set frm = Forms!ProgressMeter
        frm.Caption = " Uploading photos"
        frm!tbxProgress = ""
        frm!tbxPercent = 0
        
        'iterate through files w/in directory
        For Each iFile In ListFiles
        
            For i = 0 To UBound(aryExtensions)
                
                'check for valid images
                If InStr(1, iFile, aryExtensions(i), vbTextCompare) > 1 Then
                        
                    'prepare for insert
                    Dim Params(0 To 4) As Variant
                    
                    Params(0) = "i_usys_temp_photo"
                    Params(1) = strPath
                    Params(2) = iFile.Name 'filename
                    Params(3) = iFile.DateCreated 'file date
                    Params(4) = "U"
'Debug.Print "-----------"
'Debug.Print iFile.Name
'Debug.Print iFile.DateCreated
'Debug.Print iFile.DateLastModified
'Debug.Print iFile.Attributes
'Debug.Print iFile.Type

                    'insert photos
                    SetRecord "i_usys_temp_photo", Params
                        
                    'update system progress bar
                    iProg = iProg + 1
                    varReturn = SysCmd(acSysCmdUpdateMeter, iProg)

                    'update progress meter
                    frm.tbxMsg = "processing " & iFile.Name
                    frm.tbxPercent = (iProg / NumFiles) * 100
                    'font Terminal, character 'Û' (Alt+0219)
                    frm.tbxProgress = String(CInt(frm.tbxPercent / 100 * (frm.tbxProgress.Width / 144)), "Û") 'frm.tbxProgress & "Û" 'Û = color box

                End If
                
            Next
            
        Next
    
    Else
        MsgBox "Sorry, the directory is not valid. Please re-select it.", vbOKOnly, "Invalid Directory"
    End If
    
Exit_Handler:
    'cleanup
    varReturn = SysCmd(acSysCmdRemoveMeter)
    DoCmd.Hourglass False
    If Len(strProgForm) > 0 Then _
        DoCmd.Close acForm, strProgForm
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IngestPhotos[mod_Photo])"
    End Select
    Resume Exit_Handler
End Sub