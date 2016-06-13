Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =19440
    DatasheetFontHeight =9
    ItemSuffix =30
    Left =1365
    Top =6510
    Right =12555
    Bottom =11100
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf634d72ab6c4e440
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT tbl_Events.* FROM tbl_Events ORDER BY tbl_Events.Start_date DESC; "
    Caption ="fsub_Events"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
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
        Begin FormHeader
            Height =300
            BackColor =13025979
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =105
                    Top =60
                    Width =1155
                    Height =240
                    Name ="labStart_date"
                    Caption ="Sample date"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2880
                    Top =60
                    Width =1680
                    Height =240
                    Name ="labEvent_notes"
                    Caption ="Sampling event notes"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =18060
                    Top =60
                    Width =900
                    Height =240
                    Name ="labEntered_by"
                    Caption ="Entered by"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =16500
                    Top =60
                    Width =672
                    Height =240
                    Name ="labEntered_date"
                    Caption ="Entered"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =15060
                    Top =60
                    Width =984
                    Height =240
                    Name ="labUpdated_by"
                    Caption ="Updated by"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =13560
                    Top =60
                    Width =708
                    Height =240
                    Name ="labUpdated_date"
                    Caption ="Updated"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1500
                    Top =60
                    Width =720
                    Height =240
                    Name ="labEnd_date"
                    Caption ="End date"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =9000
                    Top =60
                    Width =900
                    Height =240
                    Name ="labVerified_by"
                    Caption ="Verified by"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7560
                    Top =60
                    Width =660
                    Height =240
                    Name ="labVerified_date"
                    Caption ="Verified"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10560
                    Top =60
                    Width =660
                    Height =240
                    Name ="labCertified_date"
                    Caption ="Certified"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =12000
                    Top =60
                    Width =945
                    Height =240
                    Name ="labCertified_by"
                    Caption ="Certified by"
                    FontName ="Arial"
                End
            End
        End
        Begin Section
            Height =780
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =1200
                    Height =252
                    ColumnWidth =1896
                    TabIndex =1
                    Name ="txtStart_date"
                    ControlSource ="Start_date"
                    Format ="yyyy mmm dd"
                    StatusBarText ="Start date of the sampling event"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16500
                    Top =60
                    Height =252
                    ColumnWidth =1896
                    TabIndex =11
                    Name ="txtEntered_date"
                    ControlSource ="Entered_date"
                    Format ="mm/dd/yyyy hh:nn"
                    StatusBarText ="Date on which data entry occurred"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13500
                    Top =60
                    Height =252
                    ColumnWidth =1896
                    TabIndex =9
                    Name ="txtUpdated_date"
                    ControlSource ="Updated_date"
                    Format ="mm/dd/yyyy hh:nn"
                    StatusBarText ="Date of the most recent edits"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15000
                    Top =60
                    Height =252
                    ColumnWidth =2568
                    TabIndex =10
                    Name ="txtUpdated_by"
                    ControlSource ="Updated_by"
                    StatusBarText ="Person who made the most recent updates"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =18000
                    Top =60
                    Height =252
                    ColumnWidth =2568
                    TabIndex =12
                    Name ="txtEntered_by"
                    ControlSource ="Entered_by"
                    StatusBarText ="Person who entered the data for this event"
                    FontName ="Arial"

                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16560
                    Top =420
                    Width =839
                    Height =252
                    TabIndex =15
                    Name ="txtProject_code"
                    ControlSource ="Project_code"
                    StatusBarText ="Project code, for linking information with other data sets and applications"
                    FontName ="Arial"

                    LayoutCachedLeft =16560
                    LayoutCachedTop =420
                    LayoutCachedWidth =17399
                    LayoutCachedHeight =672
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15540
                            Top =420
                            Width =1005
                            Height =240
                            Name ="labProject_code"
                            Caption ="Project code:"
                            FontName ="Arial"
                            LayoutCachedLeft =15540
                            LayoutCachedTop =420
                            LayoutCachedWidth =16545
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =60
                    Width =1260
                    Height =252
                    TabIndex =2
                    Name ="txtEnd_date"
                    ControlSource ="End_date"
                    Format ="yyyy mmm dd"
                    StatusBarText ="End date of the sampling event (optional)"
                    FontName ="Arial"

                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =18480
                    Top =420
                    Width =900
                    Height =252
                    TabIndex =16
                    Name ="txtDeclination"
                    ControlSource ="Declination"
                    StatusBarText ="Declination correction factor for measurement of compass bearings"
                    FontName ="Arial"

                    LayoutCachedLeft =18480
                    LayoutCachedTop =420
                    LayoutCachedWidth =19380
                    LayoutCachedHeight =672
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =17520
                            Top =420
                            Width =915
                            Height =240
                            Name ="labDeclination"
                            Caption ="Declination:"
                            FontName ="Arial"
                            LayoutCachedLeft =17520
                            LayoutCachedTop =420
                            LayoutCachedWidth =18435
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =60
                    Width =4560
                    Height =663
                    TabIndex =3
                    Name ="txtEvent_notes"
                    ControlSource ="Event_notes"
                    StatusBarText ="Comments about the sampling event"
                    FontName ="Arial"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =723
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11460
                    Top =420
                    Width =4020
                    Height =252
                    TabIndex =14
                    Name ="txtQA_notes"
                    ControlSource ="QA_notes"
                    StatusBarText ="Quality assurance comments for the selected sampling event"
                    FontName ="Arial"

                    LayoutCachedLeft =11460
                    LayoutCachedTop =420
                    LayoutCachedWidth =15480
                    LayoutCachedHeight =672
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10620
                            Top =420
                            Width =825
                            Height =240
                            Name ="labQA_notes"
                            Caption ="QA notes:"
                            FontName ="Arial"
                            LayoutCachedLeft =10620
                            LayoutCachedTop =420
                            LayoutCachedWidth =11445
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7500
                    Top =60
                    Height =252
                    TabIndex =5
                    Name ="txtVerified_date"
                    ControlSource ="Verified_date"
                    Format ="mm/dd/yyyy hh:nn"
                    StatusBarText ="Date on which data were verified"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10500
                    Top =60
                    Height =252
                    TabIndex =7
                    Name ="txtCertified_date"
                    ControlSource ="Certified_date"
                    Format ="mm/dd/yyyy hh:nn"
                    StatusBarText ="Date on which data were certified"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12000
                    Top =60
                    Height =252
                    TabIndex =8
                    Name ="txtCertified_by"
                    ControlSource ="Certified_by"
                    StatusBarText ="Person who certified data for accuracy and completeness"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9000
                    Top =60
                    Height =252
                    TabIndex =6
                    Name ="txtVerified_by"
                    ControlSource ="Verified_by"
                    StatusBarText ="Person who verified accurate data transcription"
                    FontName ="Arial"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =420
                    Width =720
                    Height =300
                    TabIndex =4
                    Name ="cmdDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Delete this event record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =600
                    Top =420
                    Width =720
                    Height =300
                    Name ="cmdEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Open the form to edit this event record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9840
                    Top =420
                    Width =720
                    TabIndex =13
                    ConditionalFormat = Begin
                        0x010000006c000000010000000000000002000000000000000500000001010000 ,
                        0xff000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x540072007500650000000000
                    End
                    Name ="cmbIs_excluded"
                    ControlSource ="Is_excluded"
                    RowSourceType ="Value List"
                    RowSource ="-1;Yes;0;No"
                    ColumnWidths ="0;576"
                    StatusBarText ="Flag to exclude the sampling event from data summary output"
                    FontName ="Arial"
                    ControlTipText ="Flag to exclude the sampling event from data summary output"

                    LayoutCachedLeft =9840
                    LayoutCachedTop =420
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =660
                    ConditionalFormat14 = Begin
                        0x010001000000000000000200000001010000ff000000ffffff00040000005400 ,
                        0x720075006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =7500
                            Top =420
                            Width =2280
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="labIs_excluded"
                            Caption ="Exclude from summary output:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =7500
                            LayoutCachedTop =420
                            LayoutCachedWidth =9780
                            LayoutCachedHeight =660
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
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
' FORM NAME:    fsub_Events_Browser
' Description:  Standard data browser subform for viewing and editing sampling event records
' Data source:  In-line record source query based on tbl_Events
' Data access:  edit only (for "admin" or "power user" app modes)
' Pages:        none
' Functions:    none
' References:   none
' Source/date:  John R. Boetsch, October 2008
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, 11/12/2008 - revised cmdEdit to depend on Loc type (as w/ Data Gateway code)
'               JRB, 2/18/2009 - simplified open statement for data entry form, and edited
'                   to open either the data entry or the recon form depending on location type
'               JRB, 12/31/2009 - added cmbIs_excluded
'               --------------------------------------------------------------------------------------
'               BLC, 6/3/2014 - Adapted for NCPN WQ Utilities tool
'               BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' =================================

' ---------------------------------
' SUB:
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
'               BLC, 8/5/2014 - changed to use setUserAccess for initializing control settings based on app mode
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    'set button caption depending on app mode
    setUserAccess Me

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

    ' Check if the current event record is certified
    If IsNull(Me.Certified_date) = False And (IsNull(Me.Updated_date) _
        Or Me.Certified_date >= Me.Updated_date) Then

        Select Case TempVars.item("UserAccessLevel")
          Case "admin", "power user"
            ' Request confirmation before allowing edits
            If MsgBox("This record is certified ... are you certain you want to edit it?", _
                vbYesNo + vbExclamation + vbDefaultButton2, _
                "Confirm certified data edit") = vbNo Then
                DoCmd.CancelEvent
                GoTo Exit_Procedure
            End If
            MsgBox "Please log the edits to certified data ..."
            DoCmd.OpenForm "frm_Edit_Log", , , , , , "Update tbl_Events"

          Case "data entry"
            ' Warn the user and disallow edits
            MsgBox "Edits to certified event data are not allowed in data entry mode", _
                vbOKOnly + vbCritical, "This event record has been certified"
            DoCmd.CancelEvent
            GoTo Exit_Procedure

          Case Else
            ' Read only
            DoCmd.CancelEvent
            GoTo Exit_Procedure
        End Select

    ElseIf TempVars.item("UserAccessLevel") = "read only" Then
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
' SUB:          cmbIs_excluded_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' ---------------------------------
Private Sub cmbIs_excluded_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    Select Case TempVars.item("UserAccessLevel")
      Case "admin", "power user"    ' Change is permitted
        Dim strMsgConfirm As String
    
        ' Warn the user about changing the value
        If Me.ActiveControl = True Then
            strMsgConfirm = "Checking this box means that data for this sampling" & _
                vbCrLf & "event won't appear in data summary output ... are you sure?"
        Else
            strMsgConfirm = "Unchecking this box means that data for this sampling" & _
                vbCrLf & "event will now appear in data summary output ... are you sure?"
        End If
        If MsgBox(strMsgConfirm, vbYesNo + vbExclamation + vbDefaultButton2, _
            "Confirm change") = vbNo Then
            Me.ActiveControl.Undo
            DoCmd.CancelEvent
        End If

      Case Else
        MsgBox "Only those with power user privileges can make this change"
        Me.ActiveControl.Undo
        DoCmd.CancelEvent
    End Select

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim varLoc As Variant
    Dim varEvent As Variant

    ' Bail out if no event
    If IsNull(Me.Event_ID) Then GoTo Exit_Procedure

    Select Case Me.Parent!Location_type
      Case "Origin"
        ' Open the main data entry form
        DoCmd.OpenForm "frm_Data_Entry", , , "[Event_ID]=""" & Me.Event_ID & _
            """", , , Me.Location_ID

      Case "Incidental"
        ' Open the rare bird observation data entry form
        DoCmd.OpenForm "frm_Rare_Bird_Obs", , , "[Event_ID]=""" & Me.Event_ID & _
            """", , , Me.Parent!Park_code

      Case Else
        ' A new or existing sampling point .. must find the related origin record and
        '   set the open form filter to that record
        Set db = CurrentDb
        Set rst = db.OpenRecordset("SELECT tbl_Locations.Location_ID, tbl_Events.Event_ID " & _
            "FROM tbl_Locations " & _
            "INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID " & _
            "WHERE (((tbl_Locations.Site_ID)=""" & Me.Parent!Site_ID & _
            """) AND ((tbl_Locations.Location_type)='Origin')" & _
            " AND ((tbl_Events.Start_date)=#" & Me.txtStart_date & "#));", dbOpenSnapshot)
        ' If no records ...
        If rst.EOF Then
            If MsgBox("No transect events match this sample date." & vbCrLf & _
                vbCrLf & "Would you like to open the point sampling form instead?", _
                vbYesNo + vbExclamation, "No matching transect visit") = vbYes Then _
                DoCmd.OpenForm "frm_Point_Establishment", , , , , , "Event=" & Me.Event_ID
            GoTo Exit_Procedure
        End If
        varLoc = rst!Location_ID
        varEvent = rst!Event_ID

        ' Filter by location and event
        DoCmd.OpenForm "frm_Data_Entry", , , "[Event_ID]=""" & varEvent & _
            """", , , varLoc
    End Select

    ' This code must come after opening the form as the subform bookmark is lost when requerying
    ' Requery the referring form first (to show any recent changes before resetting)
    gvarRefForm.Requery
    gvarRefCtl.Requery

    ' Set the global reference variables to the current form
    Set gvarRefForm = Me.Form
    Set gvarRefCtl = Me.txtStart_date

Exit_Procedure:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 91     ' Object variable not set - resume next statement
        Resume Next
      Case 2467   ' Object does not exist - resume next statement
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdDelete_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' ---------------------------------
Private Sub cmdDelete_Click()
    On Error GoTo Err_Handler

    Dim strSQL As String
    Dim strMsg As String
    Dim varResponse As VbMsgBoxResult

    ' Bail out of delete if in data entry or or read only mode
    If TempVars.item("UserAccessLevel") = "data entry" Or _
        TempVars.item("UserAccessLevel") = "read only" Then
        MsgBox "This form may only be used to delete records in power user mode." _
            , , "Cannot delete the event record"
        GoTo Exit_Procedure
    End If

    If IsNull(Me.Event_ID) = False Then
    ' If there is a record ...

    ' Confirm record deletion
        strMsg = "You are about to delete all sampling event data for the " & _
            vbCrLf & "following sample location and date:" & _
            vbCrLf & vbCrLf & Me.Parent!Park_code & "." & Me.Parent!Location_code & _
            " sampled on " & Me.Start_date
        varResponse = MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbQuestion, _
            "Delete the sampling event?")
        Select Case varResponse
          Case vbYes
            ' Extra confirmation in case of certified event data
            If IsNull(Me.Certified_date) = False And (IsNull(Me.Updated_date) _
                Or Me.Certified_date >= Me.Updated_date) Then
                If MsgBox("This is a certified record ..." & vbCrLf & vbCrLf & _
                    "Please log the deletion in the Edit Log." & vbCrLf & vbCrLf _
                    & Me.Parent!Park_code & "." & Me.Parent!Location_code & _
                    " sampled on " & Me.Start_date & vbCrLf & vbCrLf _
                    & "(Also, this is your last chance to cancel the delete!)", vbCritical + _
                    vbOKCancel, "Confirm Delete Certified Record") = vbCancel Then
                    GoTo Exit_Procedure
                Else
                    DoCmd.OpenForm "frm_Edit_Log", , , , , , "Delete tbl_Events"
                End If
            End If
            ' Build the statement to delete the sampling event (and all down-stream records)
            strSQL = "DELETE * FROM tbl_Events WHERE ((tbl_Events.Event_ID) = """ _
                & Me.Event_ID & """)"
          Case vbNo
            GoTo Exit_Procedure

          Case Else
            GoTo Exit_Procedure
        End Select

        DoCmd.SetWarnings False
        DoCmd.RunSQL strSQL
        DoCmd.SetWarnings True
        Me.Requery

    End If

Exit_Procedure:
    On Error Resume Next
    DoCmd.SetWarnings True
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2501   ' Canceled RunSQL command
        MsgBox "The record was not deleted"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub
