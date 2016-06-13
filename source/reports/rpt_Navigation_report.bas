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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =61
    Left =1980
    Top =1170
    Right =14415
    Bottom =10065
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x68c7aaa7d368e340
    End
    RecordSource ="qrpt_Navigation_report"
    Caption =" Site Navigation Report"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xd0020000d0020000d0020000d002000000000000302a0000400b000001000000 ,
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
            GroupFooter = NotDefault
            KeepTogether =2
            ControlSource ="Park_code"
        End
        Begin BreakLevel
            ControlSource ="Site_code"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Site_ID"
        End
        Begin BreakLevel
            ControlSource ="Arm_sort"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Transect_arm"
        End
        Begin BreakLevel
            ControlSource ="Loc_sort"
        End
        Begin BreakLevel
            ControlSource ="Location_code"
        End
        Begin PageHeader
            Height =720
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =600
                    Top =120
                    Width =1140
                    Height =252
                    FontSize =9
                    Name ="txtPark_code"
                    ControlSource ="Park_code"
                    StatusBarText ="Park in which the site is located"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontFamily =34
                            Left =60
                            Top =120
                            Width =528
                            Height =252
                            Name ="labPark_code"
                            Caption ="Park: "
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Rectangle
                    Top =60
                    Width =10740
                    Height =360
                    Name ="Box37"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =3
                    TextFontFamily =34
                    Left =7980
                    Top =120
                    Width =2715
                    Height =255
                    Name ="labTitle"
                    Caption ="Navigation Report"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3420
                    Top =120
                    Width =1332
                    Height =252
                    FontSize =9
                    TabIndex =1
                    Name ="txtSite_code"
                    ControlSource ="Site_code"
                    StatusBarText ="Unique alphanumeric code for each site"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =2460
                            Top =120
                            Width =888
                            Height =252
                            Name ="labSite_code"
                            Caption ="Site code"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6060
                    Top =120
                    Width =540
                    FontSize =9
                    TabIndex =2
                    Name ="txtPanel_name"
                    ControlSource ="Panel_name"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =5160
                            Top =120
                            Width =855
                            Height =240
                            Name ="labPanel_name"
                            Caption ="Panel:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1500
                    Top =480
                    Width =1080
                    FontSize =9
                    TabIndex =3
                    Name ="txtTransect_arm"
                    ControlSource ="Transect_arm"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Top =480
                    Width =1500
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="txtlabTransect_arm"
                    ControlSource ="=IIf(IsNull([Transect_arm]),Null,\"Transect arm:\")"
                    FontName ="Tahoma"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            Name ="GroupHeader0"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            BreakLevel =2
            Name ="GroupHeader2"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanShrink = NotDefault
            ForceNewPage =1
            Height =0
            BreakLevel =4
            Name ="GroupHeader1"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2880
            Name ="Detail"
            Begin
                Begin TextBox
                    FontUnderline = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Top =60
                    Width =1260
                    Height =252
                    FontSize =9
                    FontWeight =700
                    Name ="txtLocation_code"
                    ControlSource ="=[txtSite_code] & \".\" & [Location_code]"
                    StatusBarText ="Alphanumeric code for the sample location (e.g., NN1, or TO for transect origin)"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4260
                    Top =60
                    Width =480
                    TabIndex =1
                    Name ="txtAzimuth_to_point"
                    ControlSource ="Azimuth_to_point"
                    StatusBarText ="Azimuth (degrees, declination corrected) to the sampling point from the previous"
                        " point, to facilitate relocating the position"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =3180
                            Top =60
                            Width =1080
                            Height =225
                            FontSize =8
                            Name ="labAzimuth_to_point"
                            Caption ="Az. to point:"
                            FontName ="Tahoma"
                            Tag ="DetachedLabel"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1080
                    Top =660
                    Width =9660
                    Height =228
                    TabIndex =2
                    Name ="txtTravel_notes"
                    ControlSource ="Travel_notes"
                    StatusBarText ="Comments about navigation to the point – kept up to date as conditions change"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9900
                    Top =60
                    Width =840
                    TabIndex =3
                    Name ="txtLoc_established"
                    ControlSource ="Loc_established"
                    Format ="General Date"
                    StatusBarText ="Date the sample location was established"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =8940
                            Top =60
                            Width =900
                            Height =240
                            FontSize =8
                            Name ="labLoc_established"
                            Caption ="Est. date:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Line
                    Width =10740
                    Name ="Line36"
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    Left =60
                    Top =2340
                    Width =10680
                    Height =240
                    TabIndex =4
                    Name ="subMarkers"
                    SourceObject ="Report.rsub_Nav_markers"
                    LinkChildFields ="Location_ID"
                    LinkMasterFields ="Location_ID"

                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    Left =60
                    Top =2040
                    Width =10680
                    Height =240
                    TabIndex =5
                    Name ="subFeatures"
                    SourceObject ="Report.rsub_Nav_features"
                    LinkChildFields ="Location_ID"
                    LinkMasterFields ="Location_ID"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1560
                    Top =1800
                    Width =9180
                    Height =228
                    TabIndex =6
                    Name ="txtEvent_notes"
                    ControlSource ="Event_notes"
                    StatusBarText ="Comments about navigation to the point – kept up to date as conditions change"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1080
                    Top =960
                    Width =9660
                    Height =228
                    TabIndex =7
                    Name ="txtLocation_desc"
                    ControlSource ="Location_desc"
                    StatusBarText ="Comments about navigation to the point – kept up to date as conditions change"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Top =660
                    Width =1020
                    FontWeight =700
                    TabIndex =8
                    Name ="txtlabTravel"
                    ControlSource ="=IIf(IsNull([Travel_notes]),Null,\"Travel desc:\")"
                    StatusBarText ="Dominant slope aspect, in degrees, corrected for declination"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Top =1800
                    Width =1500
                    FontWeight =700
                    TabIndex =9
                    Name ="txtlabEvent_notes"
                    ControlSource ="=IIf(IsNull([Event_notes]),Null,\"Notes (last visit):\")"
                    StatusBarText ="Dominant slope aspect, in degrees, corrected for declination"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Top =1500
                    Width =1020
                    FontWeight =700
                    TabIndex =10
                    Name ="txtlabLastVisit"
                    ControlSource ="=IIf(IsNull([Start_date]),Null,\"Last visit:\")"
                    StatusBarText ="Dominant slope aspect, in degrees, corrected for declination"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1080
                    Top =1500
                    Width =1020
                    Height =228
                    TabIndex =11
                    Name ="txtStart_date"
                    ControlSource ="Start_date"
                    Format ="General Date"
                    StatusBarText ="Type of feature"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2100
                    Top =1500
                    Width =600
                    Height =228
                    TabIndex =12
                    Name ="txtStart_time"
                    ControlSource ="Start_time"
                    Format ="Short Time"
                    StatusBarText ="Type of feature"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2340
                    Top =60
                    Width =720
                    Height =228
                    TabIndex =13
                    Name ="txtTrail_or_road"
                    ControlSource ="Trail_or_road"
                    StatusBarText ="Type of feature"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =34
                            Left =1350
                            Top =60
                            Width =975
                            Height =225
                            FontSize =8
                            Name ="labTrail_or_road"
                            Caption ="Trail/road:"
                            FontName ="Tahoma"
                            Tag ="DetachedLabel"
                        End
                    End
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    Left =60
                    Top =2640
                    Width =10680
                    Height =240
                    TabIndex =14
                    Name ="subTasks"
                    SourceObject ="Report.rsub_Task_list"
                    LinkChildFields ="Location_ID"
                    LinkMasterFields ="Location_ID"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1080
                    Top =1260
                    Width =9660
                    Height =228
                    TabIndex =15
                    Name ="txtLocation_notes"
                    ControlSource ="Location_notes"
                    StatusBarText ="Comments about navigation to the point – kept up to date as conditions change"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Top =1260
                    Width =1020
                    FontWeight =700
                    TabIndex =16
                    Name ="Text58"
                    ControlSource ="=IIf(IsNull([Location_notes]),Null,\"Loc notes:\")"
                    StatusBarText ="Dominant slope aspect, in degrees, corrected for declination"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Top =960
                    Width =1020
                    FontWeight =700
                    TabIndex =17
                    Name ="Text59"
                    ControlSource ="=IIf(IsNull([Location_desc]),Null,\"Loc desc:\")"
                    StatusBarText ="Dominant slope aspect, in degrees, corrected for declination"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =5715
                    Top =60
                    Width =465
                    TabIndex =18
                    Name ="Elevation_m"
                    ControlSource ="Elevation_m"
                    Format ="Fixed"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =34
                            Left =4800
                            Top =60
                            Width =870
                            Height =225
                            FontSize =8
                            Name ="labElevation_m"
                            Caption ="Elev. (m):"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =8565
                    Top =60
                    Width =375
                    TabIndex =19
                    Name ="txtSlope_deg"
                    ControlSource ="Slope_deg"
                    Format ="Fixed"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =34
                            Left =7440
                            Top =60
                            Width =1095
                            Height =225
                            FontSize =8
                            Name ="labSlope_deg"
                            Caption ="Slope (deg):"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =6975
                    Top =60
                    Width =405
                    TabIndex =20
                    Name ="txtAspect_deg"
                    ControlSource ="Aspect_deg"
                    Format ="Fixed"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =34
                            Left =6240
                            Top =60
                            Width =735
                            Height =225
                            FontSize =8
                            Name ="labAspect_deg"
                            Caption ="Aspect:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =1950
                    Top =360
                    Width =750
                    Height =228
                    TabIndex =21
                    Name ="txtUTM_east"
                    ControlSource ="UTME"
                    Format ="Fixed"
                    StatusBarText ="Final UTM easting (zone 10N, meters), including any offsets and corrections"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =34
                            Left =1320
                            Top =360
                            Width =570
                            Height =225
                            FontSize =8
                            Name ="labUTME"
                            Caption ="UTMs:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =2700
                    Top =360
                    Width =780
                    Height =228
                    TabIndex =22
                    Name ="txtUTM_north"
                    ControlSource ="UTMN"
                    Format ="Fixed"
                    StatusBarText ="Final UTM northing (zone 10N, meters), including any offsets and corrections"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =3480
                    Top =360
                    Width =690
                    Height =228
                    TabIndex =23
                    Name ="txtDatum"
                    ControlSource ="=\"(\" & [Nav_datum] & \")\""
                    StatusBarText ="Datum of UTM_east and UTM_north"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7050
                    Top =360
                    Width =2310
                    Height =228
                    TabIndex =24
                    Name ="txtCoord_type"
                    ControlSource ="Source"
                    StatusBarText ="Coordinate type stored in UTM_east and UTM_north: target, field, post-processed"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =34
                            Left =5820
                            Top =360
                            Width =1215
                            Height =225
                            FontSize =8
                            Name ="labCoord_type"
                            Caption ="Coord type:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =5130
                    Top =360
                    Width =540
                    TabIndex =25
                    Name ="txtEst_horiz_error"
                    ControlSource ="=IIf([Est_horiz_error]<>-99,[Est_horiz_error])"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =34
                            Left =4260
                            Top =360
                            Width =840
                            Height =225
                            FontSize =8
                            Name ="labEst_horiz_error"
                            Caption ="Error (m)"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Line
                    Left =1140
                    Top =600
                    Width =9420
                    Name ="Line60"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =15
            BreakLevel =2
            Name ="GroupFooter0"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =10800
                    Name ="Line42"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =300
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =4560
                    FontSize =9
                    Name ="txtTimestamp"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6180
                    Top =60
                    Width =4560
                    FontSize =9
                    TabIndex =1
                    Name ="txtPageNo"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Tahoma"

                End
                Begin Line
                    Width =10740
                    Name ="Line35"
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
' REPORT NAME:  rpt_Navigation_report_no_species
' Description:  Field season navigation report - without species observation sections
' Data source:  qrpt_Navigation_report
' Functions:    none
' References:   fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, April 2007
' Revisions:    <name, date, desc>
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

    Me.labTitle.Caption = Me.OpenArgs & " Navigation Report"
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub Report_NoData(Cancel As Integer)
    On Error GoTo Err_Handler

    MsgBox "No records match those criteria ..."
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
