Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_Debug
' Level:        Development module
' Version:      1.03
'
' Description:  Debugging related functions & procedures for version control
'
' Source/date:  Bonnie Campbell, 2/12/2015
' Revisions:    BLC - 5/27/2015 - 1.00 - initial version
'               BLC - 7/7/2015  - 1.01 - added GetErrorTrappingOption()
'               BLC - 7/24/2015 - 1.02 - added RemoveVCSModules()
'               BLC - 6/24/2016 - 1.03 - replaced Exit_Function > Exit_Handler
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
    Set tdf = db.TableDefs(strTable) 'TableName)

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
    Set tdf = db.TableDefs(strTable)

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

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetErrorTrappingOption[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
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
   
    ModuleFile = dir(strPath, vbNormal)
    
    While ModuleFile <> ""
    
        ModuleFilePath = strPath & ModuleFile
        
        If Right(ModuleFilePath, 4) = ".bas" Then
            'add the module (file)
            Application.VBE.ActiveVBProject.VBComponents.Import ModuleFilePath
        End If
        
        'call Dir without params to get the next file in strPath
        ModuleFile = dir
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
Dim i As Integer
For i = 0 To 4
Dim myVegPlot As New VegPlot
With myVegPlot
    .EventID = 2
    .SiteID = 2
    .FeatureID = 3
    .VegTransectID = 4
    .PlotNumber = 3
    .PlotDistance = Int((10 - 0 + 1) * Rnd + 0)
    .ModalSedimentSize = "S"
    .PercentFines = Int((100 - 0 + 1) * Rnd + 0)
    .PercentWater = Int((100 - 0 + 1) * Rnd + 0)
    .UnderstoryRootedPctCover = Int((100 - 0 + 1) * Rnd + 0)
    .PlotDensity = 4
    .NoCanopyVeg = 0
    .NoRootedVeg = 0
    .HasSocialTrail = 0
    .PctFilamentousAlgae = 1
    .NoIndicatorSpecies = 0
    .SaveToDb
End With
Next

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


'Dim vw As New VegWalk
'With vw
'    .EventID = 3
'    .CollectionPlaceID = 2
'    .CollectionType = "S"
'    .StartDate = Date
'    .CreatedByID = 4
'    .LastModifiedByID = .CreatedByID
'    .SaveToDb
'End With
'
'Dim vws As New VegWalkSpecies
'With vws
'    .Init ("EPHVIR")
'    .VegWalkID = vw.id
'    .SaveToDb
'    Debug.Print "MasterCode= " & .MasterCode & vbCrLf
'    Debug.Print "UTcode= " & .UTcode & vbCrLf
'    Debug.Print "UTfamily= " & .UTfamily & vbCrLf
'    Debug.Print "UTspecies= " & .UTspecies & vbCrLf
'    Debug.Print "Nativity= " & .Nativity & vbCrLf
'    Debug.Print "Lifeform= " & .Lifeform & vbCrLf
'    Debug.Print "CommonName= " & .MasterCommonName & vbCrLf
'End With


'Dim p As New Person
'
'With p
'    .FirstName = "Maija"
'    .LastName = "Tester"
'    .MiddleInitial = "X"
'    .Username = Username
'    .Email = "a@bc.com"
'    .Role = "d"
'    .AccessRole = "admin"
'    .SaveToDb
'
'    Debug.Print .ID
'    Debug.Print .AccessLevel
'
'End With

'add multiple species for WCC, URC, ARS
'Dim ary As Variant
'ary = Array("YUCCA", "SALIX", "RIBES", "POPFRE", "TAMARIX", "SALEXI", _
'            "GNAPAL", "ERAHYP", "CALCAN", "BROTEC", "ASTER", "CYMCON", "ERODIUM", "HELHOO")
'Dim i As Integer
'Dim str As String
'
'For i = LBound(ary) To UBound(ary)
'    str = CStr(ary(i))
''    Dim wcc As New WoodyCanopySpecies
''    With wcc
''        .Init (str)
''        .PercentCover = Int((100 - 0 + 1) * Rnd + 0)
''        .VegPlotID = 4
''        .SaveToDb
''    End With
'
'    Dim urc As New UnderstoryCoverSpecies
'    With urc
'         .Init (str)
'        .PercentCover = Int((100 - 0 + 1) * Rnd + 0)
'        .VegPlotID = 4
'        .SaveToDb
'    End With
''    Dim ars As New RootedSpecies
''    With ars
''        .Init (str)
''        .PercentCover = Int((100 - 0 + 1) * Rnd + 0)
''        .VegPlotID = 4
''        .SaveToDb
''    End With
'Next

End Sub


Public Sub DoIt()
' Mark K, 10/6/2011
' http://www.access-programmers.co.uk/forums/showthread.php?t=216531

'    GetADCommonName

'GetTemplates

'Dim strSQL As String
'
'
'strSQL = GetTemplate("s_event_by_park_river_w_location", _
'                    "ParkCode" & PARAM_SEPARATOR & TempVars("ParkCode") & _
'                    "|waterway" & PARAM_SEPARATOR & TempVars("River"))

'Debug.Print strSQL

'GetTemplate "i_event"

GetTemplate ("s_app_releases")

End Sub

Public Sub testme()
'Dim rs As Recordset
'Set rs = VirtualDAORecordset(10, "tbl")
'
'Debug.Print rs("RecCount")
'
'Set rs = Nothing
'DBEngine.Rollback

'    Dim aryDensity() As Variant
'    aryDensity = SplitInt(PLOT_DENSITIES, ",")
'Dim ID As Integer
'ID = 5
'
'    DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
'                "Tagline" & PARAM_SEPARATOR & ID & _
'                "|Type" & PARAM_SEPARATOR & "info"

'    DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
'                "msg" & PARAM_SEPARATOR & "Please select a river segment." & _
'                "|Type" & PARAM_SEPARATOR & "alert"


'    DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
'                "msg" & PARAM_SEPARATOR & "Please select river segment." & _
'                "|Type" & PARAM_SEPARATOR & "info"
'    Dim mi As Variant
'    mi = "P"
'    DoCmd.RunSQL GetTemplate("TEST2")
'    Dim qdf As DAO.QueryDef
'    Set qdf = CurrentDb.QueryDefs("usys_temp_qdf")
'    qdf.SQL = GetTemplate("TEST2")
'    qdf.Parameters("mi").Value = "P"
'    'qdf.Parameters("mi").value = NullStr
'    'qdf.Parameters("mi").value = 7
'    qdf.Execute
'Dim id As Integer
'Dim tbl As String
'tbl = "Contact"
'id = 2
'
'    DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
'            tbl & PARAM_SEPARATOR & id & _
'            "|Type" & PARAM_SEPARATOR & "info"


'getting IDs

'    Debug.Print GetParkID("BLCA")
'    Debug.Print GetRiverSegmentID("CBC")
'    Debug.Print GetSiteID("BLCA", "EP")

'wia
Dim img 'As ImageFile
Dim s 'As String
Dim v 'As Vector

Set img = CreateObject("WIA.ImageFile")

'Img.LoadFile "Z:\_____LIB\pics\IMAG5402.jpg"
'XResolution(282) = 72
'YResolution(283) = 72
'ResolutionUnit(296) = 2
'YCbCrPositioning(531) = 1
'ExifISOSpeed(34855) = 580
'ExifDTOrig (36867) =2016:01:03 14:20:16
'ExifDTDigitized (36868) =2016:01:03 14:20:16
'ExifFlash(37385) = 0
'ExifFocalLength(37386) = 3.69
'ExifColorSpace(40961) = 1
'ExifPixXDim(40962) = 2592
'ExifPixYDim(40963) = 1456
'20545 (20545) =R98
'GpsLatitudeRef(1) = N
'GpsLongitudeRef(3) = W
'GpsAltitudeRef(5) = 0
'GpsAltitude(6) = 0
'29 (29) =2016:01:03
'ThumbnailCompression(20515) = 6
'ThumbnailResolutionX(20525) = 72
'ThumbnailResolutionY(20526) = 72
'ThumbnailResolutionUnit(20528) = 2
'JPEGInterFormat(513) = 854
'JPEGInterLength(514) = 28534
'Width = 2592
'Height = 1456
'Depth = 24
'HorizontalResolution = 72
'VerticalResolution = 72
'FrameCount = 1
'Img.LoadFile "C:\Users\Public\Pictures\internet\royaltyfree\MH900185150.JPG"

'Img.LoadFile "Z:\_____LIB\pics\IMAG5402.jpg"
'XXImg.LoadFile "Z:\_____LIB\pics\Zumba\2013-03-01_18-25-29_1.3gp"
'Img.LoadFile "Z:\_____LIB\pics\MDY\523781_4249158158136_1966400092_n.jpg"
'ExifUserComment (37510) =*
'Width = 960
'Height = 720
'Depth = 24
'HorizontalResolution = 96
'VerticalResolution = 96
'FrameCount = 1

'Img.LoadFile "Z:\_____LIB\pics\MDY\003.JPG"
'SoftwareUsed(305) = Picasa
'ExifDTOrig (36867) =2012:09:26 17:02:26
'ExifDTDigitized (36868) =2012:09:26 17:02:26
'ExifDTOrigSS(37521) = 0
'ExifDTDigSS(37522) = 0
'ExifPixXDim(40962) = 1600
'ExifPixYDim(40963) = 1200
'ThumbnailCompression(20515) = 6
'ThumbnailResolutionX(20525) = 72
'ThumbnailResolutionY(20526) = 72
'ThumbnailResolutionUnit(20528) = 2
'JPEGInterFormat(513) = 4416
'JPEGInterLength(514) = 4664
'Width = 1600
'Height = 1200
'Depth = 24
'HorizontalResolution = 96
'VerticalResolution = 96
'FrameCount = 1

s = "Width = " & img.Width & vbCrLf & _
    "Height = " & img.Height & vbCrLf & _
    "Depth = " & img.PixelDepth & vbCrLf & _
    "HorizontalResolution = " & img.HorizontalResolution & vbCrLf & _
    "VerticalResolution = " & img.VerticalResolution & vbCrLf & _
    "FrameCount = " & img.FrameCount & vbCrLf

If img.IsIndexedPixelFormat Then
    s = s & "Pixel data contains palette indexes" & vbCrLf
End If

If img.IsAlphaPixelFormat Then
    s = s & "Pixel data has alpha information" & vbCrLf
End If

If img.IsExtendedPixelFormat Then
    s = s & "Pixel data has extended color information (16 bit/channel)" & vbCrLf
End If

If img.IsAnimated Then
    s = s & "Image is animated" & vbCrLf
End If

If img.Properties.Exists("40091") Then
    Set v = img.Properties("40091").Value
    s = s & "Title = " & v.String & vbCrLf
End If

If img.Properties.Exists("40092") Then
    Set v = img.Properties("40092").Value
    s = s & "Comment = " & v.String & vbCrLf
End If

If img.Properties.Exists("40093") Then
    Set v = img.Properties("40093").Value
    s = s & "Author = " & v.String & vbCrLf
End If

If img.Properties.Exists("40094") Then
    Set v = img.Properties("40094").Value
    s = s & "Keywords = " & v.String & vbCrLf
End If

If img.Properties.Exists("40095") Then
    Set v = img.Properties("40095").Value
    s = s & "Subject = " & v.String & vbCrLf
End If

Dim vecProperty As WIA.Vector
Dim propEach As WIA.Property

With img
    For Each propEach In .Properties
            Select Case propEach.Name
                Case "40091"
                    Set vecProperty = propEach.Value
                    Debug.Print "Title = " & vecProperty.String

                Case "40092"
                    Set vecProperty = propEach.Value
                    Debug.Print "Comment = " & vecProperty.String

                Case "40093"
                    Set vecProperty = propEach.Value
                    Debug.Print "Author = " & vecProperty.String

                Case "40094"
                    Set vecProperty = propEach.Value
                    Debug.Print "Keywords = " & vecProperty.String

                Case "40095"
                    Set vecProperty = propEach.Value
                    Debug.Print "Subject = " & vecProperty.String

                Case Else
                'Bob77, May 9, 2011
                'http://stackoverflow.com/questions/5927828/extract-properties-from-the-image-file
                    
                    If Not (propEach.Name = "ChrominanceTable" Or _
                            propEach.Name = "LuminanceTable") Then
                    If Not varType(propEach.Value) = vbObject Then _
                    Debug.Print propEach.Name & " (" & propEach.PropertyID & ") ="; CStr(propEach.Value)
                    End If
            End Select
        Next
End With

Debug.Print s


End Sub

'Function b2d(bstr)
''convert binary string to decimal number
'    numbits = Len(bstr)
'    asum = 0
'    For i = 1 To numbits
'        asum = asum + Mid(bstr, i, 1) * 2 ^ (numbits - i)
'    Next
'    b2d = asum
'End Function

Public Sub DoItAgain()

'    Dim a As New ExtArray
'
'    a.Name = "my new array"

    TempVars("ContactID") = 1
    Dim c As New Location 'Person 'AppComment
    
    With c
''        .Comment = "test comment from dev_debug"
''        .CommentorID = 1
''        .CommentType = "test"
'        .FirstName = "Tsmeer"
'        .LastName = "Mytest"
'        .Email = "abcd@def.com"
'        .Organization = "NCPN"
'        .Username = "mylogin"
'        .IsActive = 1
'        .AccessLevel = 2
        .LocationType = "P"
        .LocationName = "X"
        .HeadtoOrientDistance = 3
        .HeadtoOrientBearing = 22
        .SaveToDb False
    End With

End Sub

Public Sub ExecuteIt()

    'getDbUserAccess
    'Debug.Print DisplayIcons("uDocument|uPDF", "|")
    
'    Dim StartFolder As String, strPics As String, strPath As String
'
'    StartFolder = GetSpecialFolderPath("FOLDERID_Pictures")  '"desktop"
'
'    strPath = BrowseFolder("Photo Directory Selection", StartFolder)
'
'    IngestPhotos strPath, "U"


''---------------------
'' Declarations
''---------------------
'Private m_show As Long
''---------------------
'' Event Declarations
''---------------------
'
''---------------------
'' Properties
''---------------------
'Public Property Let Show(Value As Long)
'    m_show = Value
'End Property
'
'Public Property Get Show() As Long
'    Show = m_show
'End Property
'Dim p As New Person
'
'p.Show = 1

''---------------------
'' Properties
''---------------------
'Private m_node As Long
'Public Property Let Node(Value As Long)
'    m_node = Value
'End Property
'
'Public Property Get Node() As Long
'    Node = m_node
'End Property
'
'
'Dim p As New Person
'p.Node = 1
'
'Dim frm As Form
'Dim Params(0 To 1) As Variant
'
''params(0) = "s_photo_data"
'Params(0) = "s_tsys_temp_photo_data"
'
'Set frm = Forms("Tree")
''LoadTree frm, frm.Controls("tvwTree").Object, "s_photo_data", params
'LoadTree frm, frm.Controls("tvwTree").Object, "s_usys_temp_photo_data", Params

''---------------------
'' Declarations
''---------------------
'Private m_visible As Long
''---------------------
'' Event Declarations
''---------------------
'
''---------------------
'' Properties
''---------------------
'Public Property Let Visible(Value As Long)
'    m_visible = Value
'End Property
'
'Public Property Get Visible() As Long
'    Visible = m_visible
'End Property
'Dim p As New Person
'
'p.Visible = 1

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim aryData() As Variant
    Dim aryFields() As Variant
    
    Set db = CurrentDb
    
    Set rs = db.OpenRecordset("SELECT '';")
    
    aryData = FetchDbTableFieldInfo("SurveyDataSourceFile")
    'use Array("",...) vs. Split() since array ==> variant array, split ==> string array
    aryFields = Array("Column|" & CLng(dbText) & "|25", _
                    "ColType|" & CLng(dbLong), _
                    "IsReqd|" & CLng(dbByte), _
                    "Length|" & CLng(dbInteger), _
                    "AllowZLS|" & CLng(dbByte))
    'Split("Column|Text|25,ColType|Long,IsReqd|Byte,Length|Integer,AllowZLS|Byte", ",")
    'Set rs = ArrayToRecordset(ary, Array("A", "B", "C"))
    
    Debug.Print ""
    
    'don't use Array("A","B","C"...) this creates Variant array
    'Split("a,b,c,d,e",",") creates String array
'    aryFields = Split("a,b,c,d,e,f,g,h,i", ",")

    Set rs = ArrayToRecordset(aryFields, aryData, "|")

    Debug.Print ""
End Sub

Public Sub DoIt2()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim Template As String
    
    Template = "i_site"
    
    Set db = CurrentDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
        
            'check if record exists in site
            .sql = GetTemplate(Template)
        
            Debug.Print .sql
        End With
    End With
End Sub

Public Sub DoIt3()
'    DoCmd.OpenForm "TableFieldList", acNormal, , , , , "Site"
'    DoCmd.OpenForm "TableFieldList", acNormal, , , , , "Photo"
    
    DoCmd.OpenForm "Task", acNormal, , , , , "Site|Site|12"

End Sub


Public Sub CheckRS()
    'populate w/ table data
    Dim rs As DAO.Recordset
    Dim aryRecord() As String
    Dim i As Integer
    
    Set rs = CurrentDb.OpenRecordset("usys_temp_rs2", dbOpenDynaset)
'    Dim rs As Object
'
'    Set rs = CreateObject("Word.application")

    Debug.Print IsRecordset(rs)

End Sub

'Public Sub PrepCSV()
'
'    Dim plots As New Collection
'
'    With ActiveSheet
'        Dim col As Long
'        For col = 1 To 5
'            Dim current As Plot
'            Set current = New Plot
'
'            With current
'                .PlotID = .Cells(10, col).Value
'                .VisitDate = .Cells(1, col).Value
'                .LocationID = .Cells(4, col).Value
'                .EventID = .Cells(3, col).Value
'                .SiteID = .Cells(7, col).Value
'                .ModalSedimentSize = .Cells(12, col).Value
'                .PercentWater = .Cells(14, col).Value
'                .Litter = .Cells(15, col).Value
'
'            Dim r As Long
'            For r = 16 To 155
'                Dim cover As String
'                Dim seedling As Byte
'                cover = .Cells(r, col).Value
'                seedling = .Cells(r, col + 1).Value
'                If cover <> vbNullString Then
'                    current.AddSpeciesCover .Cells(r, 1).Value, cover, seedling
'                End If
'            Next
'            plots.Add current
'        Next
'
'    End With
'
'    For Each current In plots
'        Debug.Print current.CsvRows
'    Next
'
'End Sub

Public Sub FixID()
'   HansUp, January 15, 2014
'   http://stackoverflow.com/questions/20738596/how-to-reset-an-access-tables-autonumber-field-it-didnt-start-from-1
Dim strSQL As String

'requires ADO, DAO fails w/ CurrentDb.Execute strSQL
'Added MSFT ADO 6.1 library
strSQL = "ALTER TABLE SOP ALTER COLUMN ID COUNTER(1,1);"
CurrentProject.Connection.Execute strSQL

    
End Sub



Public Function FixQDF(Template As String)
Dim db As DAO.Database
Dim qdf As DAO.QueryDef
Dim strSQL As String

    Set db = CurrentDb

    With db
        Set qdf = .QueryDefs("usys_temp_qdf")

        With qdf

            'check if record exists in site
            '.sql = GetTemplate(Template)

            strSQL = "PARAMETERS vers FLOAT, sflag INTEGER, contxt TEXT(255), syntx TEXT(10)," _
                     & "tname TEXT(255), prms TEXT(255), tmpl TEXT(500), rmks TEXT(255), effdate DATE, CID LONG, LMID LONG;" _
                     & "INSERT INTO tsys_Db_templates (Version, IsSupported, Context, Syntax, TemplateName, Params, Template, Remarks, EffectiveDate, RetireDate, CreateDate, CreatedBy_ID, LastModified, LastModifiedBy_ID)" _
                     & "VALUES" _
                     & "([vers],[sflag],[contxt],[syntx],[tname],[prms],[tmpl],[rmks],[effdate]," _
                     & "NULL, NOW, [CID],NOW, [LMID]);"
    'Debug.Print strSQL
            .sql = strSQL
        End With

    End With
End Function

Public Function RetParams() As String
Dim strSQL As String

strSQL = "PARAMETERS csn TEXT(25), ltype TEXT(1), lname TEXT(100)," _
& "dist INTEGER, brg INTEGER, lnotes TEXT(1500), CID LONG, LMID LONG;" _
& "insert INTO" _
& "Location" _
& "(CollectionSourceName, LocationType, LocationName, HeadtoOrientDistance_m," _
& "HeadtoOrientBearing, LocationNotes, CreateDate, CreatedBy_ID," _
& "LastModified, LastModifiedBy_ID)" _
& "Values" _
& "([csn], [ltype], [lname], [dist], [brg], [lnotes], Now, [CID], Now, [LMID]);"

GetParamsFromSQL (strSQL)

End Function

Public Function getvals() As String
'showImageFileProperties ("C:\Users\Public\Pictures\Assets\117000000000248708_1080x1920.jpg")
'showImageFileProperties ("C:\Users\indigonw\Pictures\GoPro\GOPR0729.MP4")
'ShowImageFileProperties ("C:\Users\indigonw\Documents\HTC\Gallery\ht26ys318579\phone storage\IMAG4760.jpg")
'ShowImageFileProperties ("E:\Big_Rivers\2012\CURE\P7211545.JPG")

    Dim d As Dictionary
    
    Set d = GetFileExifInfo("Z:\_____LIB\dev\git_projects\big_rivers_app\Data\P7231618.JPG")

End Function