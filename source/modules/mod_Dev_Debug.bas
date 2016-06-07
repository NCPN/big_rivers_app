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
'               BLC - 7/7/2015  - 1.01 - added GetErrorTrappingOption()
'               BLC - 7/24/2015 - 1.02 - added RemoveVCSModules()
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
'   BLC - 8/10/2015 - fixed bug referencing TableName vs strTable
' ---------------------------------
Public Sub ChangeMSysConnection(ByVal strTable As String, ByVal strConn As String)
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb()
    Set tdf = db.tabledefs(strTable) 'TableName)

    'Change the connect value
    tdf.Connect = strConn '"ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=Publishers"
    
Exit_Handler:
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
    Resume Exit_Handler
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
    
Exit_Handler:
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
    Resume Exit_Handler
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

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ChangeTSysDb[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
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
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetDebugDbPaths[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
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
'   BLC - 8/10/2015 - fixed compile bug defining i
' ---------------------------------
Public Sub DebugTest()
On Error GoTo Err_Handler

'    Dim strDbPath As String, strDb As String
'
'    'invasives be
''    strDbPath = "C:\___TEST_DATA\test\Invasives_be.accdb"
'    strDbPath = "Z:\_____LIB\dev\git_projects\TEST_DATA\test2\Invasives_be.accdb"
'    strDb = ParseFileName(strDbPath)
'
'    SetDebugDbPaths strDbPath
'
'    'NCPN master plants
''    strDbPath = "C:\___TEST_DATA\NCPN_Master_Species.accdb"
'    strDbPath = "Z:\_____LIB\dev\git_projects\TEST_DATA\test2\NCPN_Master_Species.accdb"
'    strDb = ParseFileName(strDbPath)
'
'    SetDebugDbPaths strDbPath
'
'    'progress bar test
'    DoCmd.OpenForm "frm_ProgressBar", acNormal
'    Dim i As Integer
'
'    For i = 1 To 10
'
'        Forms("frm_ProgressBar").Increment i * 10, "Preparing report..."
'    Next
'
'    'test parsing
'    ParseFileName ("C:\___TEST_DATA\test_BE_new\Invasives_be.accdb")

    Dim p_oTask As Task
    
    Set p_oTask = New Task
    With p_oTask
        .TaskType = "TaskType.Photo"
        .Task = "Testing description"
        .Status = Status.Opened
        .Priority = Priority.High
        .RequestedByID = 3
        .CompletedByID = 1
        .AddTask
    End With

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DebugTest[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
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

' ---------------------------------
' SUB:          DeleteModule
' Description:  Remove a module from the current database.
' Assumptions:  -
' Parameters:   strModule - module name (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  OurManInBananas, November 16, 2014
'               http://stackoverflow.com/questions/26948789/ms-access-use-vba-to-delete-a-module-from-access-file
' Adapted:      Bonnie Campbell, July 24, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/24/2015 - initial version
' ---------------------------------
Sub DeleteModule(strModule As String)
On Error GoTo Err_Handler

    Dim vbCom As Object

    Set vbCom = Application.VBE.ActiveVBProject.VBComponents

    vbCom.Remove VBComponent:=vbCom.item(strModule)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DeleteModule[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          RemoveModules
' Description:  Remove VCS modules from the current database.
' Assumptions:  Use the following strType for removing VCS or dev modules:
'                   VCS => "VCS_", Dev => "mod_Dev_"
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Unknown, unknown date
'   http://www.java2s.com/Code/VBA-Excel-Access-Word/Access/IteratethroughallmoduleslocatedinthedatabasereferencedbytheCurrentProjectobject.htm
' Source/date:  Bonnie Campbell, July 24, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/24/2015 - initial version
'   BLC - 8/10/2015 - adjusted to handle all (not just open) modules
' ---------------------------------
Sub RemoveModules(strType As String)
On Error GoTo Err_Handler
   
'    Dim i As Integer
'    Dim modOpenModules As Modules
    Dim modl As Variant
    
'    Set modOpenModules = Application.Modules
'
'    For i = 0 To modOpenModules.count - 1
    
'        Debug.Print modOpenModules(i).name
'
'        If Left(modOpenModules(i).name, Len(strType)) = strType Then
'            DeleteModule (modOpenModules(i).name)
'        End If
'
'    Next

    With CurrentProject
    
        For Each modl In .AllModules
            
            Debug.Print modl.Name
            
            If Left$(modl.Name, Len(strType)) = strType Then
                
                DeleteModule modl.Name
                
                Debug.Print modl.Name & " DELETED!"
            
            End If
            
        Next
    
    End With


Exit_Handler:
    'NOTE: Watch for Automation error - Unspecified error # -2147467259
    '      on exit sub, cause currently unknown 8/10/2015
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveVCSModules[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          AddModules
' Description:  Add VCS or Debug modules to the current database.
' Assumptions:  Only modules to be added are found within the directory path passed into the function
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 24, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/24/2015 - initial version
' ---------------------------------
Sub AddModules(strPath As String)
On Error GoTo Err_Handler

    Dim ModuleFilePath As String, ModuleFile As String
   
    ModuleFile = Dir(strPath, vbNormal)
    
    While ModuleFile <> ""
    
        ModuleFilePath = strPath & ModuleFile
        
        If Right(ModuleFilePath, 4) = ".bas" Then
            'add the module (file)
            Application.VBE.ActiveVBProject.VBComponents.Import ModuleFilePath
        End If
        
        'call Dir without params to get the next file in strPath
        ModuleFile = Dir
    Wend

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddModules[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
End Sub


Public Sub moduletest()

    'AddModules ("C:\__git-projects\dev_modules - Copy\")
    AddModules "C:\__git-projects\vcs_modules\"
    'AddModules "Z:\_____LIB\dev\git_projects\vcs_modules\"
    
    'RemoveModules "VCS_"
    
End Sub

Public Sub Testing()

''create event
'Dim myEvent As New EventVisit
'With myEvent
'   .LocationID = 4
'   .ProtocolID = 1
'   .SiteID = 2
'   .StartDate = Now()
'   .Save
'End With
'
'MsgBox myEvent.ID

''test regex
'Dim strName As String, strEmail As String
'strName = "O'Malley"
'strEmail = "bonnie_campbell@nps.gov"
'strEmail = "a@b.c"
'strEmail = "abc+23@a23.ca"
'
'MsgBox IsName(strName)
'MsgBox IsEmail(strEmail)

''create action
'Dim myAction As New Action
'With myAction
'    .Action = "DE"
'    .ActionDate = Date
'    .RefID = 2
'    .RefTable = "Event"
'    .ContactID = 3
'    .Save
'    MsgBox .ID
'End With

'test
'IsBetween 0.4, 0, 360, True
'IsBetween -0.4, 0, 360, True
'IsBetween -0.4, 0, 360, False
'IsBetween 0.4, 0, 360, False
'IsBetween 0, 0, 360, False
'IsBetween 40, 0, 360, False
'IsBetween 0, 0, 360, True

'test save to db
' location
'Dim myLocation As New Location
'With myLocation
'    .LocationName = "xyz"
'    .LocationType = "F"
'    .CollectionSourceName = "1"
'    .HeadtoOrientDistance = 22
'    .HeadtoOrientBearing = 361
'    .CreatedByID = 3
'    .LastModifiedByID = 3
'    .SaveToDb
'End With

' river
'Dim myRiver As New Waterway
'With myRiver
'    .ParkID = 3
'    .Name = "Green"
'    .Segment = "GAC"
'    .SaveToDb
'End With

''line distance
'Dim myTagline As New tagline
'With myTagline
'    .LineDistSource = "T"
'    .LineDistSourceID = 4
'    .LineDistType = "SC"
'    .LineDistance = 24
'    .HeightType = "V"
'    .Height = 44
'
'    .SaveToDb
'End With

''feature
'Dim myFeature As New Feature
'With myFeature
'    .Name = "G"
'    .LocationID = 5
'    .Description = "This is the place..."
'    .Directions = "These are directions..."
'    .SaveToDb
'End With
'
''site
'Dim mySite As New Site
'With mySite
''    .Code = "RR"
''    .Name = "Red Rocks"
'    .Code = "EP"
'    .Name = "East Portal"
'    .Park = "DINO"
'    .ObserverID = 5
'    .RecorderID = 4
'    .SaveToDb
'End With

''photo
'Dim myPhoto As New Photo
'With myPhoto
'    .PhotoDate = Now()
'    .PhotoType = "R"
'    .DirectionFacing = "US"
'    .PhotogLocation = ""
'    .SubjectLocation = ""
'    .PhotographerID = 5
'    .Filename = ""
'    .IsCloseup = False
'    .IsReplacement = False
'    .IsSkipped = False
'
'
'End With

'attribute directive --> give class default property
'   Attribute Value.VB_UserMemId = 0
'   Chip Pearson May 2, 2008
'   http://www.cpearson.com/excel/DefaultMember.aspx


''transducer
'Dim myTransducer As New Transducer
'With myTransducer
'    .EventID = 1
'    .TransducerType = "A"
'    .TransducerNumber = "abc123"
'    .SerialNumber = "def1342"
'    .Timing = "BD"
'    .ActionDate = Date
'    .ActionTime = Now()
'    .IsSurveyed = True
'    .SaveToDb
'End With

'CreateEnums

'Dim myVegPlot As New VegPlot
'With myVegPlot
'    .EventID = 2
'    .SiteID = 2
'    .FeatureID = 3
'    .VegTransectID = 4
'    .PlotNumber = 2
'    .PlotDistance = 22
'    .ModalSedimentSize = "S"
'    .PercentFines = 30
'    .PercentWater = 10
'    .UnderstoryRootedPctCover = 24
'    .PlotDensity = 2
'    .NoCanopyVeg = False
'    .NoRootedVeg = False
'    .HasSocialTrail = False
'    .FilamentousAlgae = True
'    .NoIndicatorSpecies = False
'    .SaveToDb
'End With

''veg transects - park required before transectnumber is set!
'Dim myVegTransect As New VegTransect
'With myVegTransect
'    .Park = "CANY"
'    .LocationID = 3
'    .EventID = 4
''    .TransectNumber = 9
''    .TransectNumber = 8
'    .TransectNumber = 1
'    .SampleDate = Date
'    .SaveToDb
'End With

'Dim myCoverSpecies As New CoverSpecies
'
'With myCoverSpecies
'    .ID = 3
'    .COfamily = ""
'    .COspecies = ""
'    .luCode = "ABC"
'    .UTfamily = ""
'    .PercentCover = 4
'    .VegPlotID = 3
'End With

'Dim mySpecies As New Species
'
'With mySpecies
'    .Init "EPHVIR" '"JUNOST" '"EPHEDRA" '"EPIGLABERRIMUM" '"PICEA"
'
'    Debug.Print "MasterCode= " & .MasterCode & vbCrLf
'    Debug.Print "UTcode= " & .UTcode & vbCrLf
'    Debug.Print "UTfamily= " & .UTfamily & vbCrLf
'    Debug.Print "UTspecies= " & .UTspecies & vbCrLf
'    Debug.Print "Nativity= " & .Nativity & vbCrLf
'    Debug.Print "Lifeform= " & .Lifeform & vbCrLf
'    Debug.Print "CommonName= " & .MasterCommonName & vbCrLf
'End With

'Dim myCoverSpecies As New CoverSpecies
'
'With myCoverSpecies
'    .Init "EPHVIR"
'    .VegPlotID = 4
'    .PercentCover = 44
'    Debug.Print "MasterCode= " & .MasterCode & vbCrLf
'    Debug.Print "UTcode= " & .UTcode & vbCrLf
'    Debug.Print "UTfamily= " & .UTfamily & vbCrLf
'    Debug.Print "UTspecies= " & .UTspecies & vbCrLf
'    Debug.Print "Nativity= " & .Nativity & vbCrLf
'    Debug.Print "Lifeform= " & .Lifeform & vbCrLf
'    Debug.Print "CommonName= " & .MasterCommonName & vbCrLf
'End With

'Dim myUnderstoryCoverSpecies As New UnderstoryCover
'With myUnderstoryCoverSpecies
'    .Init "EPHVIR"
'    .VegPlotID = 4
'    .PercentCover = 200
'    .IsSeedling = False
'    Debug.Print "MasterCode= " & .MasterCode & vbCrLf
'    Debug.Print "UTcode= " & .UTcode & vbCrLf
'    Debug.Print "UTfamily= " & .UTfamily & vbCrLf
'    Debug.Print "UTspecies= " & .UTspecies & vbCrLf
'    Debug.Print "Nativity= " & .Nativity & vbCrLf
'    Debug.Print "Lifeform= " & .Lifeform & vbCrLf
'    Debug.Print "CommonName= " & .MasterCommonName & vbCrLf
'    .SaveToDb
'End With
'
'Debug.Print TypeName(myUnderstoryCoverSpecies)

'Dim myWoodyCanopyCoverSpecies  As New WoodyCanopy
'With myWoodyCanopyCoverSpecies
'    .Init "EPHVIR"
'    .VegPlotID = 4
'    .PercentCover = 200
'    Debug.Print "MasterCode= " & .MasterCode & vbCrLf
'    Debug.Print "UTcode= " & .UTcode & vbCrLf
'    Debug.Print "UTfamily= " & .UTfamily & vbCrLf
'    Debug.Print "UTspecies= " & .UTspecies & vbCrLf
'    Debug.Print "Nativity= " & .Nativity & vbCrLf
'    Debug.Print "Lifeform= " & .Lifeform & vbCrLf
'    Debug.Print "CommonName= " & .MasterCommonName & vbCrLf
'    .SaveToDb
'    Debug.Print "ID= " & .ID & vbCrLf
'End With
'
'Debug.Print TypeName(myWoodyCanopyCoverSpecies)


Dim vw As New VegWalk
With vw
    .EventID = 3
    .CollectionPlaceID = 2
    .CollectionType = "S"
    .StartDate = Date
    .CreatedByID = 4
    .LastModifiedByID = .CreatedByID
    .SaveToDb
End With

Dim vws As New VegWalkSpecies
With vws
    .Init ("EPHVIR")
    .VegWalkID = vw.ID
    .SaveToDb
    Debug.Print "MasterCode= " & .MasterCode & vbCrLf
    Debug.Print "UTcode= " & .UTcode & vbCrLf
    Debug.Print "UTfamily= " & .UTfamily & vbCrLf
    Debug.Print "UTspecies= " & .UTspecies & vbCrLf
    Debug.Print "Nativity= " & .Nativity & vbCrLf
    Debug.Print "Lifeform= " & .Lifeform & vbCrLf
    Debug.Print "CommonName= " & .MasterCommonName & vbCrLf
End With


End Sub


Public Sub doit()
' Mark K, 10/6/2011
' http://www.access-programmers.co.uk/forums/showthread.php?t=216531

'    GetADCommonName

GetTemplates

End Sub

Public Sub testme()
Dim rs As Recordset
Set rs = VirtualDAORecordset(10, "tbl")

Debug.Print rs("RecCount")

Set rs = Nothing
DBEngine.Rollback

End Sub