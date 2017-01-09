Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =14160
    ItemSuffix =20
    Left =2430
    Top =2460
    Right =13680
    Bottom =4320
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x21cadc2ab6c4e440
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT tbl_Task_List.* FROM tbl_Task_List ORDER BY tbl_Task_List.Task_status, tb"
        "l_Task_List.Request_date DESC; "
    Caption =" Sample Location Task Item"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnDblClick ="[Event Procedure]"
    AllowDatasheetView =0
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
            Height =1020
            BackColor =13025979
            Name ="Detail"
            Begin
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =360
                    Width =900
                    Height =252
                    ColumnWidth =1932
                    TabIndex =1
                    Name ="txtRequest_date"
                    ControlSource ="Request_date"
                    Format ="Short Date"
                    StatusBarText ="Date of the task request"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =60
                            Width =1380
                            Height =255
                            Name ="labRequest_date"
                            Caption ="Request date / by"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =348
                    Width =2460
                    Height =612
                    ColumnWidth =5832
                    TabIndex =3
                    Name ="txtTask_desc"
                    ControlSource ="Task_desc"
                    StatusBarText ="Brief description of the task"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1860
                            Top =60
                            Width =1440
                            Height =252
                            Name ="labTask_desc"
                            Caption ="Brief description"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9000
                    Top =720
                    Height =252
                    ColumnWidth =1236
                    TabIndex =7
                    Name ="txtDate_completed"
                    ControlSource ="Date_completed"
                    StatusBarText ="Date the task was completed"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7620
                            Top =720
                            Width =1332
                            Height =252
                            Name ="labDate_completed"
                            Caption ="Date completed"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4380
                    Top =348
                    Width =3168
                    Height =612
                    TabIndex =4
                    Name ="txtTask_notes"
                    ControlSource ="Task_notes"
                    StatusBarText ="Notes about the task"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4380
                            Top =60
                            Width =1080
                            Height =252
                            Name ="labTask_notes"
                            Caption ="Task notes"
                            FontName ="Tahoma"
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
                    Left =60
                    Top =720
                    Width =1728
                    Height =252
                    ColumnWidth =1380
                    TabIndex =2
                    Name ="cmbRequested_by"
                    ControlSource ="Requested_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active],'Active','') AS Is_a"
                        "ctive FROM tlu_Project_Crew ORDER BY IIf([Contact_is_active],'Active','') DESC ,"
                        " tlu_Project_Crew.Contact_ID; "
                    ColumnWidths ="2592;1008"
                    StatusBarText ="Name of the person making the initial request"
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3600
                    Left =9000
                    Top =360
                    Width =1848
                    Height =252
                    ColumnWidth =1620
                    TabIndex =6
                    Name ="cmbFollowup_by"
                    ControlSource ="Followup_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active],'Active','') AS Is_a"
                        "ctive FROM tlu_Project_Crew ORDER BY IIf([Contact_is_active],'Active','') DESC ,"
                        " tlu_Project_Crew.Contact_ID; "
                    ColumnWidths ="2592;1008"
                    StatusBarText ="Name of the person following up on or completing the task"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9060
                            Top =60
                            Width =1140
                            Height =252
                            Name ="labFollowup_by"
                            Caption ="Follow-up by"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10920
                    Top =348
                    Width =3168
                    Height =612
                    TabIndex =8
                    Name ="txtFollowup_notes"
                    ControlSource ="Followup_notes"
                    StatusBarText ="Comments regarding what was done to follow-up on or complete this task"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10920
                            Top =60
                            Width =1260
                            Height =252
                            Name ="labFollowup_notes"
                            Caption ="Follow-up notes"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7620
                    Top =360
                    Width =1260
                    ColumnWidth =960
                    TabIndex =5
                    Name ="cmbTask_status"
                    ControlSource ="Task_status"
                    RowSourceType ="Value List"
                    RowSource ="Active;Complete;Inactive"
                    StatusBarText ="Status of the task"
                    DefaultValue ="\"Active\""
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7620
                            Top =60
                            Width =1140
                            Height =252
                            Name ="labTask_status"
                            Caption ="Task status"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =360
                    Width =780
                    Height =309
                    Name ="cmdCloseup"
                    Caption ="Closeup"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Open this task record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
' FORM NAME:    fsub_Task_List_Browser
' Description:  Standard data browser subform for viewing and editing task records
' Data source:  In-line SQL statement based on tbl_Task_List
' Data access:  edit, add and delete
' Pages:        none
' Functions:    none
' References:   none
' Source/date:  John R. Boetsch, July 2007
' Adapted:      Bonnie Campbell, June 2014
' Revisions:    JRB, 1/22/2008 - added validation code, revised subform layout
'               JRB, 7/31/2008 - documentation and standardization
'               JRB, 11/12/2008 - updates to setting/requerying global variables; added
'                   Dirty and Delete events
'               JRB, 12/28/2009 - updated Form_Open to only enable txtRequest_date when the
'                   parent form is the data browser
'               BLC, 6/12/2014 - Revised to use TempVars.Item("UserAccessLevel") vs. cAppMode
' =================================

Dim ctlPark As Control  ' the park control in the parent form

' ---------------------------------
' SUB:     Form_Open
' Description: Opens sub form & sets controls based on UserAccessLevel
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Created John R. Boetsch, September 2008
'               Adapted 06/12/2014 Bonnie Campbell, June 2014
'               Revised 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - Revised to use TempVars.Item("UserAccessLevel") vs. cAppMode
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Set the subform color depending on which form it belongs to
    Select Case Me.Parent.Name
      Case "frm_Data_Entry"
        Me.Detail.BackColor = 13692912  ' light straw (main data entry form)
        Set ctlPark = Me.Parent!cmbPark
        Me.txtRequest_date.Enabled = False
      Case "frm_Data_Browser"
        Me.Detail.BackColor = 13027014  ' steel blue
        Set ctlPark = Me.Parent!cmbPark_code
        Me.txtRequest_date.Enabled = True
        Me.txtRequest_date.DefaultValue = "=Date()"
      Case Else
        Me.Detail.BackColor = 13027014  ' steel blue
        Set ctlPark = Me.Parent.Parent!cmbPark_code
        Me.txtRequest_date.Enabled = False
    End Select

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_Open
' Description: Opens sub form & sets controls based on UserAccessLevel
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Created John R. Boetsch, September 2008
'               Adapted 06/12/2014 Bonnie Campbell, June 2014
'               Revised 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - Revised to use TempVars.Item("UserAccessLevel") vs. cAppMode
' ---------------------------------
Private Sub Form_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Note: this event is ignored on inserting a new record if BeforeInsert code exists

    If TempVars.item("UserAccessLevel") = "read only" Then
        ' Edits not allowed
        Cancel = True
        GoTo Exit_Procedure
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_Delete
' Description: Deletes task record
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Created John R. Boetsch, September 2008
'               Adapted 06/12/2014 Bonnie Campbell, June 2014
'               Revised 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - Revised to use TempVars.Item("UserAccessLevel") vs. cAppMode
' ---------------------------------
Private Sub Form_Delete(Cancel As Integer)
    On Error GoTo Err_Handler

    If TempVars.item("UserAccessLevel") <> "admin" And _
        TempVars.item("UserAccessLevel") <> "power user" Then
        ' Edits not allowed
        Cancel = True
        GoTo Exit_Procedure
    Else
        MsgBox "Instead of deleting the task record you may also" & vbCrLf & _
            "set the task status to 'Complete' or 'Inactive'", vbOKOnly, _
            "Reminder - Task status"
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_BeforeUpdate
' Description: Task list form validation
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Created John R. Boetsch, September 2008
'               Adapted 06/12/2014 Bonnie Campbell, June 2014
'               Revised 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - Documentation
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Validate the record before updating
    If IsNull(Me.Request_date) Then
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

' ---------------------------------
' SUB:     Form_DblClick
' Description: Opens Task List popup form
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Created John R. Boetsch, September 2008
'               Adapted 06/12/2014 Bonnie Campbell, June 2014
'               Revised 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - Documentation
' ---------------------------------
Private Sub Form_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Bail out if Location ID is missing ...
    If IsNull(Me.Location_ID) Then GoTo Exit_Procedure

    ' Save the record if it is new or there are changes
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    
    DoCmd.OpenForm "frm_Task_List_Popup", , , "[Location_ID]=""" & Me.Location_ID & _
        """ AND [Request_date] = #" & Me.Request_date & "# AND [Task_desc] = """ & _
        Me.txtTask_desc & """", acFormEdit, , ctlPark.value

    ' This code must come after opening the form as the subform bookmark is lost when requerying
    ' Requery the referring form first (to show any recent changes before resetting)
    gvarRefForm.Requery
    ' Set the global reference variables
    Set gvarRefForm = Me.Form

Exit_Procedure:
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
' SUB:     cmdCloseup_Click
' Description: Standard form close
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Created John R. Boetsch, September 2008
'               Revised 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - Documentation
' ---------------------------------
Private Sub cmdCloseup_Click()
    On Error GoTo Err_Handler

    ' Save the record if it is new or there are changes
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    If IsNull(Me.Location_ID) Then
    ' If there is no location, launch a new record
        DoCmd.OpenForm "frm_Task_List_Popup", , , , acFormAdd, , ctlPark.value
    Else:
        DoCmd.OpenForm "frm_Task_List_Popup", , , "[Location_ID]=""" & Me.Location_ID & _
            """ AND [Request_date] = #" & Me.Request_date & "# AND [Task_desc] = """ & _
            Me.txtTask_desc & """", acFormEdit, , ctlPark.value
    End If

    ' This code must come after opening the form as the subform bookmark is lost when requerying
    ' Requery the referring form first (to show any recent changes before resetting)
    gvarRefForm.Requery
    ' Set the global reference variables
    Set gvarRefForm = Me.Form

Exit_Procedure:
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
