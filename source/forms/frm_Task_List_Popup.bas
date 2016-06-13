Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =8580
    DatasheetFontHeight =9
    ItemSuffix =20
    Left =5250
    Top =2085
    Right =14130
    Bottom =10590
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xae250b0a8d01e340
    End
    RecordSource ="tbl_Task_List"
    Caption =" Sample Location Task Item"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
        Begin Section
            Height =8520
            BackColor =11050649
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =4320
                    Left =1200
                    Top =420
                    Width =1980
                    Height =252
                    ColumnWidth =2568
                    FontSize =9
                    TabIndex =1
                    Name ="cmbLocation_ID"
                    ControlSource ="Location_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Location_ID, IIf(IsNull([Site_code]),[tbl_Locations].[Park_"
                        "code],[Site_code]) & '.' & [Location_code] AS Loc_code, IIf([Location_type]='Ori"
                        "gin','Transect origin',IIf([Location_type]='New' Or [Location_type]='Survey' Or "
                        "[Location_type]='Additional','Sample point',[Location_type])) AS Loc_type, tbl_L"
                        "ocations.Location_status FROM tbl_Sites RIGHT JOIN tbl_Locations ON tbl_Sites.Si"
                        "te_ID = tbl_Locations.Site_ID WHERE (((tbl_Locations.Park_code)=[Forms]![frm_Tas"
                        "k_List_Popup]![cmbPark])) ORDER BY IIf(IsNull([Site_code]),[tbl_Locations].[Park"
                        "_code],[Site_code]) & '.' & [Location_code], tbl_Locations.Location_status; "
                    ColumnWidths ="0;1440;1440;1152"
                    StatusBarText ="Sampling location"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    OnGotFocus ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1200
                            Top =120
                            Width =1680
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="labLocation_ID"
                            Caption ="Sample location"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3300
                    Top =420
                    Width =2340
                    Height =252
                    ColumnWidth =1140
                    FontSize =9
                    TabIndex =2
                    Name ="txtRequest_date"
                    ControlSource ="Request_date"
                    StatusBarText ="Date of the task request"
                    DefaultValue ="=Date()"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3300
                            Top =120
                            Width =1380
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labRequest_date"
                            Caption ="Request date"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =1068
                    Width =6780
                    Height =252
                    ColumnWidth =6312
                    FontSize =9
                    TabIndex =4
                    Name ="txtTask_desc"
                    ControlSource ="Task_desc"
                    StatusBarText ="Brief description of the task"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =780
                            Width =1800
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labTask_desc"
                            Caption ="Brief description"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =4920
                    Height =252
                    ColumnWidth =1140
                    FontSize =9
                    TabIndex =7
                    Name ="txtDate_completed"
                    ControlSource ="Date_completed"
                    StatusBarText ="Date the task was completed"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =4620
                            Width =1440
                            Height =255
                            FontSize =9
                            Name ="labDate_completed"
                            Caption ="Date completed"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =1728
                    Width =8340
                    Height =2772
                    FontSize =9
                    TabIndex =6
                    Name ="txtTask_notes"
                    ControlSource ="Task_notes"
                    StatusBarText ="Notes about the task"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1440
                            Width =1380
                            Height =252
                            FontSize =9
                            Name ="labTask_notes"
                            Caption ="Task notes"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =420
                    Width =960
                    Height =264
                    FontSize =9
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    Name ="cmbPark"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.Park_code FROM tlu_Parks ORDER BY tlu_Parks.Park_code; "
                    StatusBarText ="Park code"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =600
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labPark_code"
                            Caption ="Park"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3600
                    Left =5760
                    Top =420
                    Width =2640
                    Height =264
                    FontSize =9
                    TabIndex =3
                    Name ="cmbRequested_by"
                    ControlSource ="Requested_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active],'Active','') AS Is_a"
                        "ctive FROM tlu_Project_Crew ORDER BY IIf([Contact_is_active],'Active','') DESC ,"
                        " tlu_Project_Crew.Contact_ID; "
                    ColumnWidths ="2592;1008"
                    StatusBarText ="Name of the person making the initial request"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5760
                            Top =120
                            Width =1440
                            Height =255
                            FontSize =9
                            Name ="labRequested_by"
                            Caption ="Requested by"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3600
                    Left =1680
                    Top =4920
                    Width =2640
                    Height =264
                    FontSize =9
                    TabIndex =8
                    Name ="cmbFollowup_by"
                    ControlSource ="Followup_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active],'Active','') AS Is_a"
                        "ctive FROM tlu_Project_Crew ORDER BY IIf([Contact_is_active],'Active','') DESC ,"
                        " tlu_Project_Crew.Contact_ID; "
                    ColumnWidths ="2592;1008"
                    StatusBarText ="Name of the person following up on or completing the task"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1680
                            Top =4620
                            Width =1320
                            Height =252
                            FontSize =9
                            Name ="labFollowup_by"
                            Caption ="Follow-up by"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =5628
                    Width =8340
                    Height =2772
                    FontSize =9
                    TabIndex =9
                    Name ="txtFollowup_notes"
                    ControlSource ="Followup_notes"
                    StatusBarText ="Comments regarding what was done to follow-up on or complete this task"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =5340
                            Width =1500
                            Height =255
                            FontSize =9
                            Name ="labFollowup_notes"
                            Caption ="Follow-up notes"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =1080
                    FontSize =9
                    TabIndex =5
                    Name ="cmbTask_status"
                    ControlSource ="Task_status"
                    RowSourceType ="Value List"
                    RowSource ="Active;Complete;Inactive"
                    StatusBarText ="Status of the task"
                    DefaultValue ="\"Active\""
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7020
                            Top =780
                            Width =1140
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labTask_status"
                            Caption ="Task status"
                            FontName ="Arial"
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
' FORM NAME:    frm_Task_List_Popup
' Description:  Standard form for viewing and editing tasks associated with sample locations
' Data source:  tbl_Task_List
' Data access:  edit, add, no delete
' Pages:        none
' Functions:    none
' References:   fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, July 2006
' Revisions:    JRB, 5/21/2007 - changed txtTask_status to cmbTask_status; locked cmbPark
'                   on existing records
'               JRB, 1/22/2008 - added validation code, revised layout
'               JRB, 5/20/2008 - updated description
'               JRB, 10/14/2008 - updated Form_Open to include read only mode; updated
'                   gvarRefForm/Ctl error code in Form_Close
'               --------------------------------------------------------------------------------------
'               BLC, 6/3/2014 - Adapted for NCPN WQ Utilities tool
'               BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
' =================================

' ---------------------------------
' SUB:          Form_Open
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
'               BLC, 7/29/2014 - updated to use TempVars.Item("Park") vs. cPark
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    If fxnSwitchboardIsOpen Then
        Select Case TempVars.item("UserAccessLevel")
          Case "read only"
            ' Disable edit controls for "read only" application status
            Me.AllowAdditions = False
            Me.AllowEdits = False
        End Select
    End If

    If Me.OpenArgs <> "" Then
        Me.cmbPark = Me.OpenArgs
    ElseIf fxnSwitchboardIsOpen Then
        Me.cmbPark = TempVars.item("Park") '[Forms]![frm_Switchboard]![cPark]
    Else:
        Me.cmbPark.SetFocus
    End If
    If Me.DataEntry = False Then Me.cmbPark.Locked = True

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbLocation_ID_GotFocus()
    On Error GoTo Err_Handler

    Me.ActiveControl.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbLocation_ID_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    Me.ActiveControl.Undo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbLocation_ID_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim strCondition As String

    If Not Me.NewRecord Then
        ' Confirm changes to the location if not a new record
        If MsgBox("Are you sure you want to change the location associated with this task?", _
            vbYesNo + vbDefaultButton2, "Confirm point change") = vbNo Then
            DoCmd.CancelEvent
            Me.cmbLocation_ID.Undo
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Validate the record before updating
    If IsNull(Me.Location_ID) Then
        MsgBox "Please enter the sample location associated with task" & _
            vbCrLf & "  or hit ESC to undo changes to the record", vbOKOnly, "Validation error"
        Me.cmbLocation_ID.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.Request_date) Then
        MsgBox "Please enter the request date for the task", vbOKOnly, "Validation error"
        Me.txtRequest_date.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.Task_desc) Then
        MsgBox "Please enter a brief task description", vbOKOnly, "Validation error"
        Me.txtTask_desc.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.Task_status) Then
        MsgBox "Please enter the task status", vbOKOnly, "Validation error"
        Me.cmbTask_status.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf Me.Task_status = "Complete" And IsNull(Me.txtDate_completed) Then
        MsgBox "Please enter the completion date", vbOKOnly, "Validation error"
        Me.txtDate_completed.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf Me.Task_status <> "Complete" And Not IsNull(Me.txtDate_completed) Then
        MsgBox "Either the task status should be 'Complete' or" & vbCrLf & _
            "the completion date should be blank.", vbOKOnly, "Validation error"
        Me.cmbTask_status.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub Form_Close()
    On Error GoTo Err_Handler

    gvarRefForm.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 91     ' Do nothing - object variable not set
      Case 2467   ' Do nothing - object does not exist
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub
