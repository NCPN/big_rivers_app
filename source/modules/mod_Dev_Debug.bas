Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_Debug
' Level:        Development module
' Version:      1.01
'
' Description:  Debugging related functions & procedures for version control
'
' Source/date:  Bonnie Campbell, 2/12/2015
' Revisions:    BLC - 5/27/2015 - 1.00 - initial version
'               BLC - 7/7/2016  - 1.01 - added GetErrorTrappingOption()
' =================================

' ===================================================================================
'  NOTE:
'       Functions and subroutines within this module are for debugging and test
'       purposes.
'
'       When the application is ready for release, this module can be
'       removed without negative impact to the application.
'
'       All mod_Debug_XX (debugging) and VCS_XX (version control system) modules can also be removed.
' ===================================================================================

' ---------------------------------
' SUB:          ChangeMSysConnection
' Description:  Change connection value for a table w/in MSys_Objects (which cannot/shouldn't be directly edited)
' Assumptions:  -
' Parameters:   strTable - table name to change (string)
'               strConn - new connection string (e.g. "ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=Publishers")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:
' Joe Kendall, 8/25/2003
' http://www.experts-exchange.com/Database/MS_Access/Q_20615117.html
' Adapted:      Bonnie Campbell, May 27, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/27/2015 - initial version
' ---------------------------------
Public Sub ChangeMSysConnection(ByVal strTable As String, ByVal strConn As String)
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb()
    Set tdf = db.tabledefs(TableName)

    'Change the connect value
    tdf.Connect = strConn '"ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=Publishers"
    
Exit_Sub:
    Set tdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ChangeMSysConnection[mod_Dev_Debug])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          ChangeMSysDb
' Description:  Change database value for a table w/in MSys_Objects (which cannot/shouldn't be directly edited)
' Assumptions:  -
' Parameters:   strTable - table name to change (string)
'               strDbPath - new database path (string) (e.g. "C:\__TEST_DATA\mydb.accdb")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:
' Joe Kendall, 8/25/2003
' http://www.experts-exchange.com/Database/MS_Access/Q_20615117.html
' Adapted:      Bonnie Campbell, May 27, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/27/2015 - initial version
' ---------------------------------
Public Sub ChangeMSysDb(ByVal strTable As String, ByVal strDbPath As String)
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb()
    Set tdf = db.tabledefs(strTable)

    'Change the database value
    tdf.Connect = ";DATABASE=" & strDbPath
    
    tdf.RefreshLink
    
Exit_Sub:
    Set tdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ChangeMSysDb[mod_Dev_Debug])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          ChangeTSysDb
' Description:  Change database value for a table w/in tsys_Link_Files & tsys_Link_Dbs
' Assumptions:  Tables (tsys_Link_Files & tsys_Link_Dbs) exist with fields as noted
' Parameters:   strDbPath - new database path (string) (e.g. "C:\__TEST_DATA\mydb.accdb")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 27, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2015 - initial version
' ---------------------------------
Public Sub ChangeTSysDb(ByVal strDbPath As String)
On Error GoTo Err_Handler
    
    Dim strDbFile As String, strSQL As String
    
    'get db file name
    strDbFile = ParseFileName(strDbPath)
    
    DoCmd.SetWarnings False
    
    'update tsys_Link_Files
    strSQL = "UPDATE tsys_Link_Files SET Link_file_path = '" & strDbPath & "' WHERE Link_file_name = '" & strDbFile & "';"
    DoCmd.RunSQL (strSQL)
    
   'update tsys_Link_Dbs
    strSQL = "UPDATE tsys_Link_Dbs SET File_path = '" & strDbPath & "' WHERE Link_db = '" & strDbFile & "';"
    DoCmd.RunSQL (strSQL)
    
    DoCmd.SetWarnings True

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ChangeTSysDb[mod_Dev_Debug])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          SetDebugDbPaths
' Description:  Change database paths for debugging in MSys_Objects, tsys_Link_Files, & tsys_Link_Dbs
' Assumptions:  Tables (tsys_Link_Files & tsys_Link_Dbs) exist with fields as noted
'               tsys_Link_Tables exists and includes desired tables
' Parameters:   strDbPath - new database path (string) (e.g. "C:\__TEST_DATA\mydb.accdb")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 27, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2015 - initial version
' ---------------------------------
Public Sub SetDebugDbPaths(ByVal strDbPath As String)
On Error GoTo Err_Handler
    Dim rs As DAO.Recordset
    Dim strDb As String, strTable As String
    
    'change the tsys_Link_Files & tsys_Link_Dbs tables
    ChangeTSysDb strDbPath
    
    'get db name
    strDb = ParseFileName(strDbPath)
    
    'iterate through linked tables w/in tsys_Link_Tables
    Set rs = CurrentDb.OpenRecordset("tsys_Link_Tables", dbOpenDynaset)
    
    If Not (rs.BOF And rs.EOF) Then
    
        Do Until rs.EOF
            
            'match table source db
            If rs!Link_db = strDb Then
                
                strTable = rs!Link_table
                ChangeMSysDb strTable, strDbPath
            
            End If
        
            rs.MoveNext
        Loop
        
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetDebugDbPaths[mod_Dev_Debug])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          DebugTest
' Description:  Run debug testing routines as noted within the subroutine.
' Assumptions:  This subroutine will be modified as needed during testing.
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 27, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2015 - initial version
' ---------------------------------
Public Sub DebugTest()
On Error GoTo Err_Handler

    Dim strDbPath As String, strDb As String

    'invasives be
'    strDbPath = "C:\___TEST_DATA\test\Invasives_be.accdb"
    strDbPath = "Z:\_____LIB\dev\git_projects\TEST_DATA\test2\Invasives_be.accdb"
    strDb = ParseFileName(strDbPath)
    
    SetDebugDbPaths strDbPath
    
    'NCPN master plants
'    strDbPath = "C:\___TEST_DATA\NCPN_Master_Species.accdb"
    strDbPath = "Z:\_____LIB\dev\git_projects\TEST_DATA\test2\NCPN_Master_Species.accdb"
    strDb = ParseFileName(strDbPath)

    SetDebugDbPaths strDbPath


    'progress bar test
    DoCmd.OpenForm "frm_ProgressBar", acNormal
    
    For i = 1 To 10
        
        Forms("frm_ProgressBar").Increment i * 10, "Preparing report..."
    Next

    'test parsing
    ParseFileName ("C:\___TEST_DATA\test_BE_new\Invasives_be.accdb")


Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DebugTest[mod_Dev_Debug])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' FUNCTION:     GetErrorTrappingOption
' Description:  Determine the error trapping option setting.
' Assumptions:  -
' Parameters:   -
' Returns:      String representing the IDE's error trapping setting.
' Throws:       none
' References:   -
' Source/date:  Luke Chung, date unknown
'               http://www.fmsinc.com/tpapers/vbacode/debug.asp
' Adapted:      Bonnie Campbell, July 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/7/2015 - initial version
' ---------------------------------
Function GetErrorTrappingOption() As String
On Error GoTo Err_Handler

  Dim strSetting As String
  
  Select Case Application.GetOption("Error Trapping")
    Case 0
      strSetting = "Break on All Errors"
    Case 1
      strSetting = "Break in Class Modules"
    Case 2
      strSetting = "Break on Unhandled Errors"
  End Select
  GetErrorTrappingOption = strSetting

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetErrorTrappingOption[mod_Dev_Debug])"
    End Select
    Resume Exit_Function
End Function