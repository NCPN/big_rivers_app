﻿Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =48
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =7080
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =495
    Top =6585
    Right =8145
    Bottom =10815
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x52b32e30d4cee240
    End
    RecordSource ="tsys_Bug_Reports"
    Caption =" Bug Reports"
    BeforeUpdate ="[Event Procedure]"
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
            Height =5100
            BackColor =9677753
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1320
                    Top =120
                    Width =3000
                    Height =270
                    ColumnWidth =1440
                    Name ="txtBug_ID"
                    ControlSource ="Bug_ID"
                    StatusBarText ="Unique identifier for each bug record"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =585
                            Height =255
                            FontWeight =700
                            Name ="labBug_ID"
                            Caption ="Bug ID"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5580
                    Top =120
                    Width =1380
                    Height =270
                    ColumnWidth =1140
                    TabIndex =1
                    Name ="txtReport_date"
                    ControlSource ="Report_date"
                    Format ="Short Date"
                    StatusBarText ="Date the bug was reported"
                    DefaultValue ="=Now()"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4380
                            Top =120
                            Width =1035
                            Height =255
                            FontWeight =700
                            Name ="labReport_date"
                            Caption ="Report date"
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
                    Left =1308
                    Top =480
                    Width =2346
                    Height =270
                    ColumnWidth =2568
                    TabIndex =2
                    Name ="cmbFound_by"
                    ControlSource ="Found_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active],'Active','') AS Is_a"
                        "ctive FROM tlu_Project_Crew ORDER BY IIf([Contact_is_active],'Active','') DESC ,"
                        " tlu_Project_Crew.Contact_ID; "
                    ColumnWidths ="2592;1008"
                    StatusBarText ="Person who found the bug"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =480
                            Width =825
                            Height =255
                            FontWeight =700
                            Name ="labFound_by"
                            Caption ="Found by"
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
                    Left =4920
                    Top =480
                    Width =2046
                    Height =270
                    ColumnWidth =2568
                    TabIndex =3
                    Name ="cmbReported_by"
                    ControlSource ="Reported_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active],'Active','') AS Is_a"
                        "ctive FROM tlu_Project_Crew ORDER BY IIf([Contact_is_active],'Active','') DESC ,"
                        " tlu_Project_Crew.Contact_ID; "
                    ColumnWidths ="2592;1008"
                    StatusBarText ="Person who filled out this bug report"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3720
                            Top =480
                            Width =1080
                            Height =255
                            FontWeight =700
                            Name ="labReported_by"
                            Caption ="Reported by"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    AllowAutoCorrect = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =840
                    Width =5640
                    Height =2400
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="txtReport_details"
                    ControlSource ="Report_details"
                    StatusBarText ="Nature of the bug report"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =840
                            Width =675
                            Height =465
                            FontWeight =700
                            Name ="labReport_details"
                            Caption ="Report details"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =3360
                    Width =1800
                    Height =270
                    ColumnWidth =1140
                    TabIndex =5
                    Name ="txtFix_date"
                    ControlSource ="Fix_date"
                    Format ="Short Date"
                    StatusBarText ="Date the bug was fixed"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =3360
                            Width =705
                            Height =255
                            FontWeight =700
                            Name ="labFix_date"
                            Caption ="Fix date"
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
                    Left =4908
                    Top =3360
                    Width =2106
                    Height =270
                    ColumnWidth =2568
                    TabIndex =6
                    Name ="cmbFixed_by"
                    ControlSource ="Fixed_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active],'Active','') AS Is_a"
                        "ctive FROM tlu_Project_Crew ORDER BY IIf([Contact_is_active],'Active','') DESC ,"
                        " tlu_Project_Crew.Contact_ID; "
                    ColumnWidths ="2592;1008"
                    StatusBarText ="Person who fixed the bug"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4020
                            Top =3360
                            Width =750
                            Height =255
                            FontWeight =700
                            Name ="labFixed_by"
                            Caption ="Fixed by"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    AllowAutoCorrect = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =3720
                    Width =5640
                    Height =1320
                    ColumnWidth =3000
                    TabIndex =7
                    Name ="txtFix_details"
                    ControlSource ="Fix_details"
                    StatusBarText ="Notes on fix"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =3720
                            Width =900
                            Height =255
                            FontWeight =700
                            Name ="labFix_details"
                            Caption ="Fix details"
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
' FORM NAME:    fsub_Bug_Reports
' Description:  Standard subform for viewing and creating application bug reports
' Data source:  tsys_Bug_Reports
' Data access:  edit only
' Pages:        none
' Functions:    none
' References:   fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, May 5, 2006
' Revisions:    JRB, 3/13/2008 - put Found_by and Reported_by defaults in quotes
'               JRB, 9/5/2008 - fixed reference to switchboard control in BeforeUpdate
'               JRB, 10/6/2008 - updated to allow edits (default is locked as subform)
' =================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' If the form was opened from the switchboard, set the found/reported by
    '   controls to the current user indicated in the main switchboard control
    If Me.OpenArgs = 1 Then
        Me.cmbFound_by.DefaultValue = """" & Environ("Username") & """"
        Me.cmbReported_by.DefaultValue = """" & Environ("Username") & """"
        Me.cmbFound_by.SetFocus
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Upon saving the record, associate the bug with the release indicated
    '   in the main switchboard control
    If Me.OpenArgs = 1 And fxnSwitchboardIsOpen Then
        Me.Release_ID = [Forms]![frm_Switchboard]![cmbVersion]
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
