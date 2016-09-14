Option Compare Database
Option Explicit

' =================================
' Module:       ContextMenu
' Level:        Framework form
' Version:      1.02
'
' Description:  Context menu object related properties, events, functions & procedures for UI display
'
' Requires:     Microsoft Office 14.0 Object Library for custom context menus
' Source/date:  Bonnie Campbell, 11/3/2015
' References:
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
'               BLC - 6/28/2016 - 1.01 - revised s_site_list to s_site_list_active
'               BLC - 9/7/2016  - 1.02 - revised s_feature_list to s_feature_list_by_site
' =================================

'---------------------
' Declarations
'---------------------

'---------------------
' Menus
'---------------------

' ---------------------------------
' Sub:          CreateMenu
' Description:  Creates various right click context menus
' Notes:
'   By running this code once, you can create commandbar "MyContext" in the database.
'   Can then go into your target form in design view:
'       Properties->Other->Shortcut Menu=Yes
'       Properties->Other->Shortcut Menu Bar=MyContext
'   Add an AutoExec macro can be created to run this sub & create context menus
'   Deletes any existing MyContext and builds it from scratch.
'
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
'   http://www.excelbanter.com/showthread.php?t=397624
'   http://www.experts-exchange.com/articles/12904/Understanding-and-using-CommandBars-Part-II-Creating-your-own.html
'   http://spreadsheetpage.com/index.php/site/tip/identifying_commandbar_images/
'   http://supportingtech.blogspot.com/2011/03/microsoft-faceid-numbers-for-vba.html
'   https://msdn.microsoft.com/en-us/library/office/ff194247.aspx
'   https://bytes.com/topic/access/answers/949589-how-do-i-create-custom-right-click-menu
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Public Sub CreateMenu(Context As String)
On Error GoTo Err_Handler
    
    Dim cbar As CommandBar
    Dim mnuItem As CommandBarControl
    Dim mnu As String

    Select Case Context
        Case "park"
            mnu = "park"
        Case "river"
            mnu = "river"
        Case "site"
            mnu = "site"
        Case "feature"
            mnu = "feature"
        Case "level"
            mnu = "setlevel"
        Case "transect"
            mnu = "dataentry"
        Case "plot"
        Case "plotestablish"
        
        '-- observations --
        Case "obs-photos"
        Case "obs-Transducers"
        
        '-- trip prep --
        Case "vegplot"
            mnu = "vegplot"
        Case "photos"
            mnu = "photos"
        Case "vegwalk"
            mnu = "vegwalk"
        Case "transducer"
            mnu = "transducer"
        Case "task"
            mnu = "task"
            
        '-- xx --
        Case "comment"
    End Select

    'check for existing menu
    For Each cbar In Application.CommandBars
        If cbar.Name = mnu Then
            CommandBars(mnu).Delete
            Exit For
        End If
    Next cbar
    
    With CommandBars.Add(Name:=mnu, Position:=msoBarPopup)
 
        Select Case mnu
            
            Case "park"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="park")
                mnuItem.Caption = "Set &Park"
                mnuItem.OnAction = "mnuSetPark"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="data_entry")
                mnuItem.Caption = "Set &User"
                mnuItem.OnAction = "mnuSetDataEntryUser"
                
            Case "river"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="river")
                mnuItem.Caption = "Set &River"
                mnuItem.OnAction = "mnuSetRiver"
            
            Case "site"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="site")
                mnuItem.Caption = "Set &Site"
                mnuItem.OnAction = "mnuSetSite"
            
            Case "feature"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="feature")
                mnuItem.Caption = "Set &Feature"
                mnuItem.OnAction = "mnuSetFeature"
            
            Case "setlevel"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="park")
                mnuItem.Caption = "Set &Park"
                mnuItem.OnAction = "mnuSetPark"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="river", Before:=1)
                mnuItem.Caption = "Set &River"
                mnuItem.OnAction = "mnuSetRiver"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="site", Before:=2)
                mnuItem.Caption = "Set &Site"
                mnuItem.OnAction = "mnuSetSite"
            
            Case "dataentry"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="task")
                mnuItem.Caption = "Add &Task"
                mnuItem.OnAction = "mnuAddTask"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="comment", Before:=1)
                mnuItem.Caption = "Comment"
                mnuItem.OnAction = "mnuComment"
            'Case "report"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="export", Before:=1)
                mnuItem.Caption = "Export to PDF/Excel"
                mnuItem.OnAction = "mnuExport"
                Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="openPDF", Before:=1)
                mnuItem.Caption = "Open as PDF"
                mnuItem.OnAction = "mnuOpenPDF"
            Case "form"
        End Select

    End With
 
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateMenu[mod_ContextMenu])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          CreateDynamicMenu
' Description:  Creates various right click context menus from database records
' Notes:
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
'   Dave Applegate, January 30, 2005
'   http://vbcity.com/forums/t/88686.aspx
'   Tony Jollans, October 16, 2014
'   http://www.tek-tips.com/viewthread.cfm?qid=1739131
' Source/date:  Bonnie Campbell, May 22, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/22/2016 - initial version
'   BLC - 6/28/2016 - revised s_site_list to s_site_list_active
'   BLC - 9/7/2016  - revised s_feature_list to s_feature_list_by_site
' ---------------------------------
Public Sub CreateDynamicMenu(Context As String)
On Error GoTo Err_Handler
    
    Dim cbar As CommandBar
    Dim mnuItem As CommandBarControl
    Dim mnu As String, action As String
    Dim rs As DAO.Recordset

    Select Case Context
        Case "park"
            mnu = "park"
        Case "river"
            mnu = "river"
        Case "site"
            mnu = "site"
        Case "feature"
            mnu = "feature"
        Case "level"
            mnu = "setlevel"
        Case "transect"
            mnu = "dataentry"
        Case "plot"
        Case "plotestablish"
        
        '-- observations --
        Case "obs-photos"
        Case "obs-Transducers"
        
        '-- trip prep --
        Case "vegplot"
            mnu = "vegplot"
        Case "photos"
            mnu = "photos"
        Case "vegwalk"
            mnu = "vegwalk"
        Case "transducer"
            mnu = "transducer"
        Case "task"
            mnu = "task"
            
        '-- xx --
        Case "comment"
    End Select

    'check for existing menu
    For Each cbar In Application.CommandBars
        If cbar.Name = mnu Then
            CommandBars(mnu).Delete
            Exit For
        End If
    Next cbar
    
    With CommandBars.Add(Name:=mnu, Position:=msoBarPopup)
 
        Select Case mnu
            
            Case "park"
                Set rs = CurrentDb.OpenRecordset(GetTemplate("s_get_parks"), dbOpenDynaset)
                
                If Not (rs.BOF And rs.EOF) Then
                
                    rs.MoveFirst
                    
                    Do Until rs.EOF
                        Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="park")
                        mnuItem.Caption = rs.fields("ParkCode") '"Set &Park"
                        
                        action = "mnuSetPark"
                        mnuItem.Parameter = rs.fields("ParkCode")
                        mnuItem.OnAction = "mnuSetPark"
                        
                        rs.MoveNext
                    Loop
                
                Else
                    'clear menu
                
                End If
                
            Case "river"
                If Not IsNull(TempVars("ParkCode")) Then
                    Set rs = CurrentDb.OpenRecordset(GetTemplate("s_river_list", "ParkCode" & PARAM_SEPARATOR & TempVars.item("ParkCode")), dbOpenDynaset)
                    
                    If Not (rs.BOF And rs.EOF) Then
                        
                        rs.MoveFirst
                        
                        Do Until rs.EOF
                            Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="river")
                            mnuItem.Caption = rs.fields("Segment")
                            mnuItem.Parameter = rs.fields("Segment")
                            mnuItem.OnAction = "mnuSetRiver"
                            
                            rs.MoveNext
                        Loop
                        
                    Else
                        'clear menu
                    
                    End If
                End If
                
            Case "site"
                If Not IsNull(TempVars("ParkCode")) Then
                    Set rs = CurrentDb.OpenRecordset(GetTemplate("s_site_list_active", "ParkCode" & PARAM_SEPARATOR & TempVars.item("ParkCode")), dbOpenDynaset)
                    
                    If Not (rs.BOF And rs.EOF) Then
                    
                        rs.MoveFirst
                        
                        Do Until rs.EOF
                            Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="site")
                            mnuItem.Caption = rs.fields("Site")
                            mnuItem.Parameter = Left(Right(rs.fields("Site"), 3), 2)
                            mnuItem.OnAction = "mnuSetSite"
                            
                            rs.MoveNext
                        Loop
                        
                    Else
                        'clear menu
                    
                    End If
                End If
                
            Case "feature"
                If Not IsNull(TempVars("ParkCode")) And Not IsNull(TempVars("SiteCode")) Then
                    'Set rs = CurrentDb.OpenRecordset(GetTemplate("s_feature_list", "ParkCode" & PARAM_SEPARATOR & TempVars.item("ParkCode")), dbOpenDynaset)
                    Set rs = GetRecords("s_feature_list_by_site")
                    
                    If Not (rs.BOF And rs.EOF) Then
                    
                        rs.MoveFirst
                        
                        Do Until rs.EOF
                            Set mnuItem = .Controls.Add(Type:=msoControlButton, Parameter:="feature")
                            mnuItem.Caption = rs.fields("Feature")
                            mnuItem.Parameter = rs.fields("Feature")
                            mnuItem.OnAction = "mnuSetFeature"
                            
                            rs.MoveNext
                        Loop
                    
                    Else
                        'clear menu
                        
                    End If
                End If
        End Select
        
    End With

 
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 3078 'Can't find template/template by that name
        'This error could be caused by GetTemplates() finding a
        'duplicate template -- tsys_Db_Templates finds more than one w/ same name
        'Best option is to notify them to contact database manager
        MsgBox "Application cannot find the template." & vbCrLf & vbCrLf & _
            "Context: " & Context & vbCrLf & vbCrLf & _
            "Please contact your data manager to resolve this issue." & vbCrLf & vbCrLf & _
            "Error #" & Err.Number & " - CreateDynamicMenu[mod_ContextMenu]:" & vbCrLf & _
            Err.Description, vbExclamation, "Db Template for Creating ContextMenu Not Found!"

            '********** PREVENT FORM FROM OPENING ***********
            'Main form will not have correct context menus if duplicate templates exist
            'Close it if it's open
            If FormIsOpen("Main") Then
                DoCmd.Close acForm, "Main"
            End If
            
            '********** FATAL ERROR ****************
            'terminate *ALL* VBA code to prevent other popups
            End
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateMenu[mod_ContextMenu])"
    End Select
    Resume Exit_Handler
End Sub


' =================================
'   Context Menu Actions
' =================================

' ---------------------------------
' Sub:          VARIOUS
' Description:  Accomplishes various right click context menu actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
'   BLC - 6/28/2016 - added setting of park & river segment IDs
' ---------------------------------
Public Sub mnuAddTask()
    
    DoCmd.OpenForm "Tile", acNormal
    
End Sub

Public Sub mnuComment()
    
    DoCmd.OpenForm "Comment", acNormal, OpenArgs:="comment"
    
End Sub

Public Sub mnuSetPark()
    'DoCmd.OpenForm "_SelectSingle", acNormal, OpenArgs:="park"
    
    Dim ParkCode As String
    Dim iClearBelow As Integer
    
    iClearBelow = 0
    ParkCode = CommandBars.ActionControl.Parameter
    
    'set global
    TempVars.Add "ParkCode", ParkCode
    TempVars.Add "ParkID", GetParkID(ParkCode)
    
    'update dynamic menus
    CreateDynamicMenu "river"
    CreateDynamicMenu "site"
    CreateDynamicMenu "feature"
    
    'update calling form
    Call Forms("Main").UpdateBreadcrumb(iClearBelow)
    
End Sub

Public Sub mnuSetRiver()
    'DoCmd.OpenForm "_SelectSingle", acNormal, OpenArgs:="river"
    
    Dim segment As String
    Dim iClearBelow As Integer
    
    iClearBelow = 1
    segment = CommandBars.ActionControl.Parameter
    
    'set global
    TempVars.Add "River", segment
    TempVars.Add "RiverID", GetRiverSegmentID(segment)
    
    'update dynamic menus
    CreateDynamicMenu "site"
    CreateDynamicMenu "feature"
    
    'update calling form
    Call Forms("Main").UpdateBreadcrumb(iClearBelow)
    
    
End Sub

Public Sub mnuSetSite()
    'DoCmd.OpenForm "_SelectSingle", acNormal, OpenArgs:="site"

    Dim SiteCode As String
    Dim iClearBelow As Integer
    
    iClearBelow = 2
    SiteCode = CommandBars.ActionControl.Parameter
    
    'set global
    TempVars.Add "SiteCode", SiteCode
    TempVars.Add "SiteID", GetSiteID(TempVars("ParkCode"), SiteCode)
    
    'update dynamic menus
    CreateDynamicMenu "feature"
    
    'update calling form
    Call Forms("Main").UpdateBreadcrumb(iClearBelow)
End Sub

Public Sub mnuSetFeature()
    'DoCmd.OpenForm "_SelectSingle", acNormal, OpenArgs:="feature"
    Dim Feature As String
    Dim iClearBelow As Integer
    
    iClearBelow = 3
    Feature = CommandBars.ActionControl.Parameter
    
    'set global
    TempVars.Add "Feature", Feature
    TempVars.Add "SiteID", GetFeatureID(TempVars("ParkCode"), Feature)
    
    'update calling form
    Call Forms("Main").UpdateBreadcrumb(iClearBelow)

End Sub

Public Sub mnuSetDataEntryUser()
    
    DoCmd.OpenForm "_SelectSingle", acNormal, OpenArgs:="data_entry"
    
End Sub




'---------------------------------
Sub CreateSimpleShortcutMenu()
    Dim btnShortcutMenu As Office.CommandBar
     
    ' Create a shortcut menu named "SimpleShortcutMenu.
    Set btnShortcutMenu = CommandBars.Add("SimpleShortcutMenu", msoBarPopup, False, True)
     
    ' Add the Remove Filter/Sort command.
    btnShortcutMenu.Controls.Add Type:=msoControlButton, ID:=605
 
    ' Add the Filter By Selection command.
    btnShortcutMenu.Controls.Add Type:=msoControlButton, ID:=640
     
    Set btnShortcutMenu = Nothing
     
End Sub


Sub CreateShortcutMenuWithGroups()
    Dim btnRightClick As Office.CommandBar
 
 ' Create the shortcut menu.
    Set btnRightClick = CommandBars.Add("cmdFormFiltering", msoBarPopup, False, True)
     
    With btnRightClick
        ' Add the Find command.
        .Controls.Add msoControlButton, 141, , , True
         
        ' Start a new grouping and add the Sort Ascending command.
        .Controls.Add(msoControlButton, 210, , , True).BeginGroup = True
         
        ' Add the Sort Descending command.
        .Controls.Add msoControlButton, 211, , , True
         
        ' Start a new grouping and add the Remove Filer/Sort command.
        .Controls.Add(msoControlButton, 605, , , True).BeginGroup = True
         
        ' Add the Filter by Selection command.
        .Controls.Add msoControlButton, 640, , , True
         
        ' Add the Filter Excluding Selection command.
        .Controls.Add msoControlButton, 3017, , , True
         
        ' Add the Between... command.
        .Controls.Add msoControlButton, 10062, , , True
    End With
 
Set btnRightClick = Nothing
End Sub

Sub CreateReportShortcutMenu()
    Dim btnRightClick As Office.CommandBar
    Dim btnControl As Office.CommandBarControl
 
   ' Create the shortcut menu.
    Set btnRightClick = CommandBars.Add("cmdReportRightClick", msoBarPopup, False, True)
 
    With btnRightClick
         
        ' Add the Print command.
        Set btnControl = .Controls.Add(msoControlButton, 2521, , , True)
        ' Change the caption displayed for the control.
        btnControl.Caption = "Quick Print"
         
        ' Add the Print command.
        Set btnControl = .Controls.Add(msoControlButton, 15948, , , True)
        ' Change the caption displayed for the control.
        btnControl.Caption = "Select Pages"
         
        ' Add the Page Setup... command.
        Set btnControl = .Controls.Add(msoControlButton, 247, , , True)
        ' Change the caption displayed for the control.
        btnControl.Caption = "Page Setup"
         
        ' Add the Mail Recipient (as Attachment)... command.
        Set btnControl = .Controls.Add(msoControlButton, 2188, , , True)
        ' Start a new group.
        btnControl.BeginGroup = True
        ' Change the caption displayed for the control.
        btnControl.Caption = "Email Report as an Attachment"
         
        ' Add the PDF or XPS command.
        Set btnControl = .Controls.Add(msoControlButton, 12499, , , True)
        ' Change the caption displayed for the control.
        btnControl.Caption = "Save as PDF/XPS"
         
        ' Add the Close command.
        Set btnControl = .Controls.Add(msoControlButton, 923, , , True)
        ' Start a new group.
        btnControl.BeginGroup = True
        ' Change the caption displayed for the control.
        btnControl.Caption = "Close Report"
    End With
     
    Set btnControl = Nothing
    Set btnRightClick = Nothing
End Sub

' http://www.experts-exchange.com/Database/MS_Access/Q_27830781.html
Public Function CreateCMenu()
On Error Resume Next

    CommandBars("MyContext").Delete

    Dim cmb As CommandBar 'Object
    Dim cmbBtn1 As CommandBarButton 'Object
    Dim cmbBtn2 As CommandBarButton 'Object

    Set cmb = CommandBars.Add("MyContext", _
               msoBarPopup, False, False)    ' msoBarPopup = 5
        With cmb
              ' add cut, copy, and paste buttons with the "magic number" technique that assigns
              ' appearance and behavior. The magic number goes in as the second parameter

            .Controls.Add msoControlButton, _
                  21, , , True  ' 21=Cut, msoControlButton=1
            .Controls.Add msoControlButton, _
                      19, , , True  '19= Copy
            .Controls.Add msoControlButton, _
                      22, , , True  ' 22=Paste

' add customized buttons with our caption and function name -- second param is blank
            Set cmbBtn1 = .Controls.Add(msoControlButton, _
                                    , , , True)
            With cmbBtn1
                .BeginGroup = True
                .Caption = "Create New"
                .OnAction = "=CreateNewOrder()"
                .FaceId = 59  'smiley face
            End With
           
            Set cmbBtn2 = .Controls.Add(msoControlButton, _
                                    , , , True)
            With cmbBtn2
                .Caption = "Reset"
                .OnAction = "=ClearOrder()"
            End With
        End With

   
End Function

Public Function reposit()
'    For Each ctrl In Me.Controls
'        If ctrl.ControlType = acCommandButton Then
'            ctrl.BackStyle = acNormal
'            ctrl.HoverColor = lngYelLime
'            ctrl.HoverForeColor = lngPurple
'            ctrl.PressedColor = lngLtBlue
'        End If
'    Next
    
'    Me.Refresh
End Function