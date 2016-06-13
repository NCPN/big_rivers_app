Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =9
    ItemSuffix =40
    Left =270
    Top =210
    Right =12135
    Bottom =6630
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x409bb98e0f9fe340
    End
    RecordSource ="qrpt_Task_list"
    Caption ="Plot Task List"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x3804000038040000380400003804000000000000602700002409000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =9
            FontWeight =700
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =18
            BorderLineStyle =0
            BackStyle =0
            FontName ="Times New Roman"
            AsianLineBreak =255
        End
        Begin ListBox
            TextFontFamily =18
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Times New Roman"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =18
            BorderLineStyle =0
            BackStyle =0
            FontName ="Times New Roman"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            KeepTogether =2
            ControlSource ="Park_code"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Loc_code"
        End
        Begin BreakLevel
            ControlSource ="Request_date"
        End
        Begin PageHeader
            Height =360
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =3
                    Left =7200
                    Width =2760
                    Height =360
                    FontSize =14
                    Name ="labTitle"
                    Caption ="Task List"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3240
                    Top =60
                    Height =285
                    FontSize =11
                    TabIndex =1
                    Name ="txtLocation_code"
                    ControlSource ="Loc_code"
                    StatusBarText ="Alphanumeric code for the sample location"

                    Begin
                        Begin Label
                            TextAlign =0
                            Left =1800
                            Top =60
                            Width =1320
                            Height =285
                            FontSize =11
                            Name ="labLocation_code"
                            Caption ="Plot number"
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =720
                    Top =60
                    Width =960
                    Height =285
                    ColumnWidth =1110
                    FontSize =11
                    Name ="txtPark_code"
                    ControlSource ="Park_code"
                    StatusBarText ="Park code"

                    Begin
                        Begin Label
                            TextAlign =0
                            Left =60
                            Top =60
                            Width =555
                            Height =285
                            FontSize =11
                            Name ="labPark_code"
                            Caption ="Park"
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader2"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            BreakLevel =1
            Name ="GroupHeader4"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =2340
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1860
                    Top =180
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="txtRequest_date"
                    ControlSource ="Request_date"
                    Format ="Short Date"
                    StatusBarText ="Date of the task request"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Left =4560
                    Top =180
                    Width =5460
                    Height =270
                    FontSize =10
                    TabIndex =2
                    Name ="txtTask_desc"
                    ControlSource ="Task_desc"
                    StatusBarText ="Brief description of the task"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1500
                    Top =2040
                    Width =1500
                    Height =270
                    FontSize =10
                    TabIndex =5
                    Name ="txtDate_completed"
                    ControlSource ="Date_completed"
                    Format ="Short Date"
                    StatusBarText ="Date the task was completed"

                    Begin
                        Begin Label
                            Left =60
                            Top =2040
                            Width =1560
                            Height =270
                            FontSize =10
                            Name ="labDate_completed"
                            Caption ="Date completed"
                            Tag ="DetachedLabel"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Left =3060
                    Top =540
                    Width =6960
                    Height =660
                    FontSize =10
                    TabIndex =7
                    Name ="txtTask_notes"
                    ControlSource ="=\"Task notes:  \" & [Task_notes]"
                    StatusBarText ="Notes about the task"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Left =3060
                    Top =1320
                    Width =6960
                    Height =1020
                    FontSize =10
                    TabIndex =8
                    Name ="txtFollowup_notes"
                    ControlSource ="=\"Follow up notes:  \" & [Followup_notes]"
                    StatusBarText ="Comments regarding what was done to follow-up on or complete this task"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =720
                    Top =900
                    Width =1260
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="txtTask_status"
                    ControlSource ="Task_status"
                    StatusBarText ="Status of the task"

                    Begin
                        Begin Label
                            Left =60
                            Top =900
                            Width =600
                            Height =270
                            FontSize =10
                            Name ="labTask_status"
                            Caption ="Status"
                            Tag ="DetachedLabel"
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1320
                    Top =540
                    Width =1680
                    Height =270
                    FontSize =10
                    TabIndex =3
                    Name ="txtRequested_by"
                    ControlSource ="Requested_by"
                    StatusBarText ="Name of the person making the initial request"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1200
                    Top =1320
                    Width =1740
                    Height =270
                    FontSize =10
                    TabIndex =6
                    Name ="txtFollowup_by"
                    ControlSource ="Followup_by"
                    StatusBarText ="Name of the person following up on or completing the task"

                End
                Begin Label
                    Left =60
                    Top =1320
                    Width =1200
                    Height =270
                    FontSize =10
                    Name ="labFollowup_by"
                    Caption ="Followup by"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =60
                    Top =1680
                    Width =1860
                    Height =270
                    FontSize =10
                    Name ="labCompleted"
                    Caption ="Completed?    Y    N"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =480
                    Top =180
                    Width =1365
                    Height =270
                    FontSize =10
                    Name ="labRequest_date"
                    Caption ="Date requested"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3420
                    Top =180
                    Width =1260
                    Height =270
                    FontSize =10
                    Name ="labTask_desc"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =60
                    Top =540
                    Width =1290
                    Height =270
                    FontSize =10
                    Name ="labRequested_by"
                    Caption ="Requested by"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =3060
                    Top =1260
                    Width =6900
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =60
                    Width =9960
                    Name ="Line34"
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    RunningSum =1
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =180
                    Width =360
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="txtCounter"
                    ControlSource ="=1"

                End
            End
        End
        Begin PageFooter
            Height =270
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Width =4560
                    Height =270
                    Name ="txtFooterTime"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5460
                    Width =4560
                    Height =270
                    TabIndex =1
                    Name ="txtFooterPage"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                End
                Begin Line
                    Width =9360
                    Name ="Line29"
                End
                Begin Line
                    Width =10020
                    Name ="Line30"
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
' REPORT NAME:  rpt_Task_List
' Description:  Data output report of active tasks associated with sampling locations
' Data source:  qrpt_Task_list
' Functions:    none
' References:   fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, July 2008
' Revisions:    JRB, 10/20/2008 - updated record source to show only active tasks
'               JRB, 5/20/2009 - minor revisions, data source
'               JRB, 6/8/2009 - moved filtering code to switchboard
' =================================

Private Sub Report_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Close the report if the switchboard is not open
    If fxnSwitchboardIsOpen = False Then
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this report to function properly.", , "Cannot open the report ..."
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

    Me.labTitle.Caption = Me.OpenArgs & " Task List"

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub Report_NoData(Cancel As Integer)
    On Error GoTo Err_Handler

    MsgBox "No task records match those criteria ..."
    DoCmd.CancelEvent
    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2585   ' no records returned, do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub
