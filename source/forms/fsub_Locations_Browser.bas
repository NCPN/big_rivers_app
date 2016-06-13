Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DefaultView =2
    ViewsAllowed =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4260
    ItemSuffix =30
    Left =615
    Top =3975
    Right =11670
    Bottom =5670
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4088da2ab6c4e440
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT tbl_Locations.* FROM tbl_Locations ORDER BY tbl_Locations.Location_code; "
    Caption ="fsub_Locations_Browser"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowFormView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    DatasheetForeColor12 =33554432
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Section
            Height =5220
            BackColor =13025979
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =120
                    ColumnWidth =945
                    Name ="cmbPark_code"
                    ControlSource ="Park_code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.Park_code FROM tlu_Parks ORDER BY tlu_Parks.Park_code; "
                    StatusBarText ="Park code (optional except for incidental observations not associated with sites"
                        ")"
                    BeforeUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =780
                            Top =120
                            Width =900
                            Height =240
                            Name ="labPark_code"
                            Caption ="Park_code"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListRows =20
                    ListWidth =4608
                    Left =2100
                    Top =480
                    ColumnWidth =885
                    TabIndex =1
                    Name ="cmbSite_ID"
                    ControlSource ="Site_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Sites.Site_ID, tbl_Sites.Site_code, tbl_Sites.Park_code, tbl_Sites.Si"
                        "te_status, tbl_Sites.Site_name FROM tbl_Sites ORDER BY tbl_Sites.Site_code, tbl_"
                        "Sites.Site_status; "
                    ColumnWidths ="0;720;1008;1152;1728"
                    StatusBarText ="Site membership of the sample location (transect)"
                    BeforeUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =480
                            Width =645
                            Height =240
                            Name ="labSite_ID"
                            Caption ="Site_ID"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =840
                    ColumnWidth =1245
                    TabIndex =2
                    Name ="txtLocation_code"
                    ControlSource ="Location_code"
                    StatusBarText ="Alphanumeric code for the sample location (e.g., NN1, or TO for transect origin)"
                    BeforeUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =840
                            Width =1185
                            Height =240
                            Name ="labLocation_code"
                            Caption ="Location_code"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5040
                    Left =2100
                    Top =1200
                    ColumnWidth =1095
                    TabIndex =3
                    Name ="cmbLocation_type"
                    ControlSource ="Location_type"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Location_Type.Location_type, tlu_Location_Type.Loc_type_desc FROM tlu"
                        "_Location_Type ORDER BY tlu_Location_Type.Sort_order; "
                    ColumnWidths ="1008;4032"
                    StatusBarText ="Indicates the type of sample location"
                    BeforeUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =1200
                            Width =1125
                            Height =240
                            Name ="labLocation_type"
                            Caption ="Loc_type"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =3720
                    ColumnWidth =1275
                    TabIndex =10
                    Name ="txtLocation_name"
                    ControlSource ="Location_name"
                    StatusBarText ="Brief colloquial name of the sample location (generally only used as a landmark "
                        "name for incidental observations)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =3720
                            Width =1215
                            Height =240
                            Name ="labLocation_name"
                            Caption ="Location_name"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =4080
                    ColumnWidth =1095
                    TabIndex =11
                    Name ="txtUTME_public"
                    ControlSource ="UTME_public"
                    StatusBarText ="UTM easting (zone 10N, meters).  Note: in addition to any measurement error, the"
                        "se coordinates may have been offset up to 2 km from their actual position."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =4080
                            Width =1095
                            Height =240
                            Name ="labUTME_public"
                            Caption ="UTME_public"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =4440
                    ColumnWidth =1110
                    TabIndex =12
                    Name ="txtUTMN_public"
                    ControlSource ="UTMN_public"
                    StatusBarText ="UTM northing (zone 10N, meters).  Note: in addition to any measurement error, th"
                        "ese coordinates may have been offset up to 2 km from their actual position."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =4440
                            Width =1110
                            Height =240
                            Name ="labUTMN_public"
                            Caption ="UTMN_public"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =4800
                    ColumnWidth =1425
                    TabIndex =13
                    Name ="txtPublic_offset"
                    ControlSource ="Public_offset"
                    StatusBarText ="Type of processing performed to make coordinates publishable"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =4800
                            Width =1035
                            Height =240
                            Name ="labPublic_offset"
                            Caption ="Public_offset"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =1920
                    ColumnWidth =1140
                    TabIndex =5
                    Name ="cmbTrail_or_road"
                    ControlSource ="Trail_or_road"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Trail_Or_Road.Trail_code FROM tlu_Trail_Or_Road ORDER BY tlu_Trail_Or"
                        "_Road.Sort_order; "
                    StatusBarText ="Indicates whether or not the sample location is along a road or trail"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =1920
                            Width =1050
                            Height =240
                            Name ="labTrail_or_road"
                            Caption ="Trail_or_road"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =2280
                    ColumnWidth =1125
                    TabIndex =6
                    Name ="txtAzimuth_to_point"
                    ControlSource ="Azimuth_to_point"
                    StatusBarText ="Azimuth (degrees, declination corrected) to the sampling point from the previous"
                        " point, to facilitate relocating the position; 999 signifies points along the tr"
                        "ail"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =2280
                            Width =1335
                            Height =240
                            Name ="labAzimuth_to_point"
                            Caption ="Azimuth_to_point"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2100
                    Top =2640
                    TabIndex =7
                    Name ="cmbDirection_changed"
                    ControlSource ="Direction_changed"
                    RowSourceType ="Value List"
                    RowSource ="Yes;No"
                    StatusBarText ="Indicates whether the azimuth to the point was changed to accommodate navigation"
                    Format ="Yes/No"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =660
                            Top =2640
                            Width =1470
                            Height =240
                            Name ="labDirection_changed"
                            Caption ="Direction_changed"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =3000
                    ColumnWidth =1365
                    TabIndex =8
                    Name ="txtLoc_established"
                    ControlSource ="Loc_established"
                    Format ="yyyy mmm dd"
                    StatusBarText ="Date the sample location was established"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =3000
                            Width =1275
                            Height =240
                            Name ="labLoc_established"
                            Caption ="Loc_established"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =3360
                    ColumnWidth =1455
                    TabIndex =9
                    Name ="txtLoc_discontinued"
                    ControlSource ="Loc_discontinued"
                    Format ="yyyy mmm dd"
                    StatusBarText ="Date the sample location was discontinued"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =3360
                            Width =1380
                            Height =240
                            Name ="labLoc_discontinued"
                            Caption ="Loc_discontinued"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =2100
                    Top =1560
                    ColumnWidth =1005
                    TabIndex =4
                    Name ="cmbLocation_status"
                    ControlSource ="Location_status"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Site_Status.Site_status, tlu_Site_Status.Site_status_desc FROM tlu_Si"
                        "te_Status ORDER BY tlu_Site_Status.Sort_order; "
                    ColumnWidths ="1008;4752"
                    StatusBarText ="Status of the sample location (blank for incidental locations)"
                    BeforeUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =1560
                            Width =1245
                            Height =240
                            Name ="labLocation_status"
                            Caption ="Loc_status"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' FORM NAME:    fsub_Locations_Browser
' Description:  Standard data browser subform for viewing the list of sample locations
'                   associated with a site
' Data source:  In-line query based on tbl_Locations
' Data access:  edit, add; no delete (use main form to delete)
' Pages:        none
' Functions:    none
' References:   fxnSwitchboardIsOpen, fxnGUIDGen
' Source/date:  John R. Boetsch, October 2008
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    <name, date, desc>
'               --------------------------------------------------------------------------------------
'               BLC, 6/3/2014 - Adapted for NCPN WQ Utilities tool
'               BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' =================================

' ---------------------------------
' SUB:          Form_Current
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' ---------------------------------
Private Sub Form_Current()
    On Error GoTo Err_Handler

    ' Allow power users/admins to add records
    If TempVars.item("UserAccessLevel") = "admin" Or _
       TempVars.item("UserAccessLevel") = "power user" Then
        Me.AllowAdditions = True
    Else
        Me.AllowAdditions = False
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          Form_Dirty
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' ---------------------------------
Private Sub Form_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Note: this event is ignored on inserting a new record if BeforeInsert code exists

    If TempVars.item("UserAccessLevel") <> "admin" And _
        TempVars.item("UserAccessLevel") <> "power user" Then
        ' Edits not allowed
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          Form_BeforeInsert
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    If TempVars.item("UserAccessLevel") <> "admin" And _
        TempVars.item("UserAccessLevel") <> "power user" Then
        ' Insert not allowed
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

    ' Create the GUID primary key value
    Me.Location_ID = fxnGUIDGen
    Me.Location_status = "active"

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          Form_Delete
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' ---------------------------------
Private Sub Form_Delete(Cancel As Integer)
    On Error GoTo Err_Handler

    If TempVars.item("UserAccessLevel") <> "admin" And _
        TempVars.item("UserAccessLevel") <> "power user" Then
        ' Delete not allowed
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          Form_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
'               BLC, 7/29/2014 - updated to use TempVars.Item("User") vs. cUser
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Validate the record and cancel updates if not valid
    If IsNull(Me.cmbPark_code) Then
        MsgBox "Please enter the park", vbOKOnly, "Validation error"
        Me.cmbPark_code.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtLocation_code) Then
        MsgBox "Please enter the location code", vbOKOnly, "Validation error"
        Me.txtLocation_code.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.cmbLocation_type) Then
        MsgBox "Please indicate the location type", vbOKOnly, "Validation error"
        Me.cmbLocation_type.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.cmbLocation_status) Then
        MsgBox "Please fill in the location status", vbOKOnly, "Validation error"
        Me.cmbLocation_status.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf Me.cmbLocation_type <> "Incidental" And IsNull(Me.cmbSite_ID) Then
        ' Site ID required for all except incidental/rare bird obs locations
        MsgBox "Please enter the site", vbOKOnly, "Validation error"
        Me.cmbSite_ID.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ' Check that the park matches the park in the sites table
    ElseIf Not IsNull(Me.cmbSite_ID) And Me.Parent!Park_code <> Me.cmbPark_code Then
        MsgBox "The park does not match the park in the site record", vbOKOnly, _
            "Validation error"
        Me.cmbPark_code.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ' Make sure that the record is not a duplicate prior to saving
    ' ... Site ID and location code are unique for all except incidental/rare bird obs locations
    ElseIf Me.cmbLocation_type <> "Incidental" And DCount("*", "tbl_Locations", _
        "[Site_ID]=""" & Me.cmbSite_ID & """ AND Location_code=""" & Me.txtLocation_code & _
        """ AND [Location_ID] <> """ & Me.Location_ID & """") > 0 Then
        MsgBox "A record with the same site and location code already exists.", _
        vbOKOnly, "Duplicate record found"
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf Me.txtLoc_discontinued < Me.Loc_established Then
        MsgBox "The discontinued date cannot be before the establisment date", _
            vbOKOnly, "Validation error"
        Me.txtLoc_discontinued.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtLoc_discontinued) = False Then
        If IsNull(Me.txtLoc_established) = False And _
            (Me.txtLoc_established > Me.txtLoc_discontinued) Then
            MsgBox "The discontinued date must be after the establishment date", _
                vbOKOnly, "Validation error"
            Me.txtLoc_discontinued.SetFocus
            DoCmd.CancelEvent
            GoTo Exit_Procedure
        End If
        If Me.cmbLocation_status = "Active" Then
            MsgBox "This location has a discontinued date. If the location" & vbCrLf & _
                "is discontinued, please change the status", vbOKOnly, "Validation error"
            Me.cmbLocation_status.SetFocus
            DoCmd.CancelEvent
            GoTo Exit_Procedure
        End If
    End If
    If IsNull(Me.txtLoc_established) = False And Me.cmbLocation_status = "proposed" Then
        MsgBox "This location has an establishment date," & vbCrLf & _
            "but its status still indicates proposed", vbOKOnly, "Validation error"
        Me.cmbLocation_status.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

    ' Prior to saving, include a timestamp for edits
    If Me.NewRecord = False Then Me.Loc_updated = Now()
    ' Add the current user name to updated by
    If fxnSwitchboardIsOpen Then
        If IsNull(TempVars.item("User")) = False Then
            Me.Loc_updated_by = TempVars.item("User")
        Else
            Me.Loc_updated_by = Environ("Username")
        End If
    Else
        Me.Loc_updated_by = Environ("Username")
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' The next set of procedures are for interacting with the user on edits to the current record

' ---------------------------------
' SUB:          cmbPark_code_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' ---------------------------------
Private Sub cmbPark_code_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    If Not IsNull(Me.cmbSite_ID) And Me.Parent!Park_code <> Me.cmbPark_code Then
        MsgBox "The park does not match the park in the site record", vbOKOnly, _
        "Validation error"
        DoCmd.CancelEvent
        Me.ActiveControl.Undo
    ElseIf Not Me.NewRecord Then
        If MsgBox("Are you sure you want to change the park for this location?", _
            vbYesNo + vbDefaultButton2, "Confirm change to critical info") = vbNo Then
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbSite_ID_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' ---------------------------------
Private Sub cmbSite_ID_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    If Not Me.NewRecord Then
        If MsgBox("Are you sure you want to change the site for this location?", _
            vbYesNo + vbDefaultButton2, "Confirm change to critical info") = vbNo Then
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
        ElseIf Me.cmbSite_ID = "" Then
            MsgBox "Site ID cannot be set to null", vbOKOnly
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          txtLocation_code_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' ---------------------------------
Private Sub txtLocation_code_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    If Not Me.NewRecord Then
        If MsgBox("Are you sure you want to change the location code?", _
            vbYesNo + vbDefaultButton2, "Confirm change to critical info") = vbNo Then
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbLocation_type_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' ---------------------------------
Private Sub cmbLocation_type_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    If Not Me.NewRecord Then
        If MsgBox("Are you sure you want to change the location type?", _
            vbYesNo + vbDefaultButton2, "Confirm change to critical info") = vbNo Then
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbLocation_status_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' ---------------------------------
Private Sub cmbLocation_status_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    If Not Me.NewRecord Then
        If MsgBox("Are you sure you want to change the location status?", _
            vbYesNo + vbDefaultButton2, "Confirm change to critical info") = vbNo Then
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
