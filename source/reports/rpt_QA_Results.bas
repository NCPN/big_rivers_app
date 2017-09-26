Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13800
    DatasheetFontHeight =10
    ItemSuffix =24
    Left =2895
    Top =480
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x1aef8d7cadebe240
    End
    RecordSource ="tbl_QA_Results"
    Caption =" Quality Assurance Report"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xf003000080040000f00300008004000000000000e83500003801000001000000 ,
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
            SortOrder = NotDefault
            GroupHeader = NotDefault
            ControlSource ="Time_frame"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Data_scope"
        End
        Begin BreakLevel
            ControlSource ="Query_name"
        End
        Begin BreakLevel
            ControlSource ="QA_user"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =360
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    Width =4794
                    Height =306
                    FontSize =11
                    Name ="labTitle"
                    Caption ="Quality Assurance and Data Validation Results"
                End
            End
        End
        Begin PageHeader
            Height =630
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =11895
                    Width =1860
                    Height =306
                    ColumnWidth =2040
                    FontSize =11
                    TabIndex =1
                    Name ="txtReport_run_time"
                    ControlSource ="=Now()"
                    Format ="mm/dd/yy hh:nn"
                    StatusBarText ="Run time of the query results"

                    Begin
                        Begin Label
                            TextAlign =3
                            Left =10065
                            Width =1770
                            Height =300
                            FontSize =11
                            Name ="labReport_run_time"
                            Caption ="Report run time:"
                            Tag ="DetachedLabel"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    Width =3135
                    Height =300
                    FontSize =11
                    Name ="labTimeframe"
                    Caption ="Data timeframe (season/year):"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5535
                    Width =786
                    Height =306
                    FontSize =11
                    Name ="labScope"
                    Caption ="Scope:"
                    Tag ="DetachedLabel"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6375
                    Width =3534
                    Height =285
                    FontSize =11
                    TabIndex =2
                    Name ="cmbScope_text"
                    ControlSource ="Data_scope"
                    RowSourceType ="Value List"
                    RowSource ="0;Uncertifed events only;1;Both certified and uncertified events;2;Certified eve"
                        "nts only"
                    ColumnWidths ="0;1440"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3195
                    Width =2274
                    Height =306
                    FontSize =11
                    Name ="txtTime_frame"
                    ControlSource ="Time_frame"

                End
                Begin Label
                    FontUnderline = NotDefault
                    Left =420
                    Top =360
                    Width =1155
                    Height =270
                    FontSize =10
                    Name ="labQuery_name"
                    Caption ="Query name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    Left =4845
                    Top =360
                    Width =975
                    Height =270
                    FontSize =10
                    Name ="labQuery_result"
                    Caption ="N Records"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontUnderline = NotDefault
                    Left =5940
                    Top =360
                    Width =1635
                    Height =270
                    FontSize =10
                    Name ="labQuery_description"
                    Caption ="Query description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontUnderline = NotDefault
                    Left =9120
                    Top =360
                    Width =1395
                    Height =270
                    FontSize =10
                    Name ="labRemedy_desc"
                    Caption ="Remedy details"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =0
                    Left =4260
                    Top =360
                    Width =525
                    Height =270
                    FontSize =10
                    Name ="labQuery_type"
                    Caption ="Type"
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    Left =12650
                    Top =360
                    Width =645
                    Height =270
                    FontSize =10
                    Name ="labQA_user"
                    Caption ="QA by"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            Name ="GroupHeader1"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            BreakLevel =1
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =312
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =420
                    Top =60
                    Width =3600
                    Height =252
                    ColumnWidth =4932
                    FontSize =9
                    Name ="txtQuery_name"
                    ControlSource ="Query_name"
                    StatusBarText ="Name of the quality assurance query"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5100
                    Top =60
                    Width =600
                    Height =252
                    FontSize =9
                    TabIndex =1
                    Name ="txtQuery_result"
                    ControlSource ="Query_result"
                    StatusBarText ="Query result, typically the number of records returned"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Left =5760
                    Top =60
                    Width =3300
                    Height =252
                    ColumnWidth =9960
                    FontSize =9
                    TabIndex =2
                    Name ="txtQuery_description"
                    ControlSource ="Query_description"
                    StatusBarText ="Description of the query"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Left =9120
                    Top =60
                    Width =3120
                    Height =252
                    ColumnWidth =6648
                    FontSize =9
                    TabIndex =3
                    Name ="txtRemedy_desc"
                    ControlSource ="Remedy_desc"
                    StatusBarText ="Evaluation expression built into the query"

                End
                Begin TextBox
                    RunningSum =1
                    IMESentenceMode =3
                    Top =60
                    Width =360
                    FontSize =9
                    TabIndex =4
                    Name ="txtCount"
                    ControlSource ="=1"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4080
                    Top =60
                    Width =1020
                    FontSize =9
                    TabIndex =5
                    Name ="cmbQuery_type"
                    ControlSource ="Query_type"
                    RowSourceType ="Value List"
                    RowSource ="1;Critical;2;Warning;3;Information"
                    ColumnWidths ="0;2160"
                    StatusBarText ="Severity of data errors being trapped: 1=critical, 2=warning, 3=information"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12300
                    Top =60
                    Width =1500
                    FontSize =9
                    TabIndex =6
                    Name ="txtUser"
                    ControlSource ="QA_user"
                    StatusBarText ="Run time of the query results"

                End
            End
        End
        Begin PageFooter
            Height =300
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Top =60
                    Width =5040
                    FontSize =9
                    Name ="txtTimestamp"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Top =60
                    Width =5040
                    FontSize =9
                    TabIndex =1
                    Name ="txtPage"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                End
                Begin Line
                    Width =13380
                    Name ="Line14"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
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
' REPORT NAME:  rpt_QA_Results
' Description:  Standard report of data validation results and remedies
' Data source:  tbl_QA_Results
' Functions:    none
' References:   fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, July 2008
' Revisions:    JRB, 10/29/2008 - changed to span multiple-years/scopes
' =================================

Private Sub Report_NoData(Cancel As Integer)
    On Error GoTo Err_Handler

    MsgBox "No records match the filter criteria ..."
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
