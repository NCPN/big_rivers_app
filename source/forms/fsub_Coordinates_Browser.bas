Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AutoResize = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =22920
    DatasheetFontHeight =10
    ItemSuffix =42
    Left =405
    Top =2850
    Right =11610
    Bottom =6075
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8fccd52ab6c4e440
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT tbl_Events.Start_date AS Coord_date, tbl_Events.Location_ID, tbl_Events.U"
        "pdated_date, tbl_Events.Certified_date, tbl_Coordinates.* FROM tbl_Events INNER "
        "JOIN tbl_Coordinates ON tbl_Events.Event_ID = tbl_Coordinates.Event_ID ORDER BY "
        "Abs([Is_best]) DESC , tbl_Events.Start_date DESC; "
    Caption ="fsub_Coordinates"
    BeforeUpdate ="[Event Procedure]"
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
        Begin FormHeader
            Height =480
            BackColor =13025979
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =5040
                    Top =240
                    Width =534
                    Height =240
                    Name ="labIs_best"
                    Caption ="Best"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1380
                    Top =240
                    Width =852
                    Height =240
                    Name ="labUTM_east"
                    Caption ="UTME (X)"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2280
                    Top =240
                    Width =912
                    Height =240
                    Name ="labUTM_north"
                    Caption ="UTMN (Y)"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =4140
                    Width =900
                    Height =444
                    Name ="labEst_horiz_error"
                    Caption ="Est. horiz. error (m) "
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3360
                    Top =240
                    Width =588
                    Height =240
                    Name ="labDatum"
                    Caption ="Datum"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5760
                    Top =240
                    Width =1176
                    Height =240
                    Name ="labCoord_type"
                    Caption ="Coord type"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8640
                    Top =240
                    Width =1440
                    Height =240
                    Name ="labCoordinate_notes"
                    Caption ="Coordinate notes"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =16020
                    Top =240
                    Width =960
                    Height =240
                    Name ="labCoord_updated_by"
                    Caption ="Updated by"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =20880
                    Top =240
                    Width =1146
                    Height =240
                    Name ="labCoord_created_date"
                    Caption ="Date created"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =18720
                    Top =240
                    Width =1146
                    Height =240
                    Name ="labCoord_updated"
                    Caption ="Date updated"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =240
                    Width =945
                    Height =240
                    Name ="labCoord_date"
                    Caption ="Coord. date"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    Left =7380
                    Top =240
                    Width =960
                    Height =240
                    Name ="labCoord_label"
                    Caption ="Coord label"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =12780
                    Top =240
                    Width =960
                    Height =240
                    Name ="labElevation_m"
                    Caption ="Elevation_m"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =13800
                    Top =240
                    Width =1020
                    Height =240
                    Name ="labAspect_deg"
                    Caption ="Aspect_deg"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =14880
                    Top =240
                    Width =900
                    Height =240
                    Name ="labSlope_deg"
                    Caption ="Slope_deg"
                    FontName ="Arial"
                End
            End
        End
        Begin Section
            Height =720
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5040
                    Top =60
                    Width =660
                    Height =252
                    ColumnWidth =2568
                    TabIndex =5
                    ConditionalFormat = Begin
                        0x010000006c000000010000000000000002000000000000000500000001010000 ,
                        0x0000ff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x540072007500650000000000
                    End
                    Name ="cmbIs_best"
                    ControlSource ="Is_best"
                    RowSourceType ="Value List"
                    RowSource ="Yes;No"
                    StatusBarText ="Indicates whether this set of coordinates is the best available for this locatio"
                        "n"
                    FontName ="Arial"
                    Format ="Yes/No"

                    ConditionalFormat14 = Begin
                        0x0100010000000000000002000000010100000000ff00ffffff00040000005400 ,
                        0x720075006500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1380
                    Top =60
                    Width =840
                    Height =252
                    ColumnWidth =2568
                    TabIndex =1
                    Name ="txtUTM_east"
                    ControlSource ="UTM_east"
                    StatusBarText ="Final UTM easting (zone 10N, meters), including any offsets and corrections"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2280
                    Top =60
                    Width =957
                    Height =252
                    ColumnWidth =2568
                    TabIndex =2
                    Name ="txtUTM_north"
                    ControlSource ="UTM_north"
                    StatusBarText ="Final UTM northing (zone 10N, meters), including any offsets and corrections"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4320
                    Top =60
                    Width =648
                    Height =252
                    ColumnWidth =2568
                    TabIndex =4
                    Name ="txtEst_horiz_error"
                    ControlSource ="Est_horiz_error"
                    StatusBarText ="Estimated horizontal error (meters) of UTM_east and UTM_north"
                    FontName ="Arial"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3300
                    Top =60
                    Width =960
                    Height =252
                    ColumnWidth =2568
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"10\""
                    Name ="cmbDatum"
                    ControlSource ="Datum"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Datum"
                    StatusBarText ="Datum of UTM_east and UTM_north"
                    FontName ="Arial"

                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =16860
                    Top =420
                    Width =1680
                    Height =252
                    ColumnWidth =2568
                    TabIndex =20
                    Name ="txtGPS_file_name"
                    ControlSource ="GPS_file_name"
                    StatusBarText ="GPS rover file used for data downloads"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =15660
                            Top =420
                            Width =1194
                            Height =240
                            Name ="labGPS_file_name"
                            Caption ="GPS file name:"
                            FontName ="Arial"
                            Tag ="DetachedLabel"
                        End
                    End
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13620
                    Top =420
                    Width =1980
                    Height =252
                    ColumnWidth =2568
                    TabIndex =19
                    ColumnInfo ="\"\";\"\";\"10\";\"50\""
                    Name ="cmbGPS_model"
                    ControlSource ="GPS_model"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_GPS_Model.GPS_model FROM tlu_GPS_Model ORDER BY tlu_GPS_Model.Sort_or"
                        "der; "
                    StatusBarText ="Make and model of GPS unit used to collect field coordinates"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =12660
                            Top =420
                            Width =903
                            Height =240
                            Name ="labGPS_model"
                            Caption ="GPS model:"
                            FontName ="Arial"
                            Tag ="DetachedLabel"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8580
                    Top =60
                    Width =4200
                    Height =252
                    ColumnWidth =3000
                    TabIndex =8
                    Name ="txtCoordinate_notes"
                    ControlSource ="Coordinate_notes"
                    StatusBarText ="Notes about this set of coordinates"
                    FontName ="Arial"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5760
                    Top =60
                    Width =1500
                    Height =252
                    TabIndex =6
                    Name ="cmbCoord_type"
                    ControlSource ="Coord_type"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Coord_Type.Coord_type FROM tlu_Coord_Type ORDER BY tlu_Coord_Type.Sor"
                        "t_order; "
                    StatusBarText ="Coordinate type stored in UTM_east and UTM_north: target, field, post-processed"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1200
                    Height =252
                    Name ="txtCoord_date"
                    ControlSource ="Coord_date"
                    Format ="yyyy mmm dd"
                    StatusBarText ="Start date of the sampling event when coordinate data were collected"
                    FontName ="Arial"

                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1380
                    Top =420
                    Width =840
                    Height =252
                    TabIndex =12
                    Name ="txtField_UTME"
                    ControlSource ="Field_UTME"
                    StatusBarText ="Field UTM Easting (zone 10N), in meters"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =420
                            Width =1125
                            Height =255
                            Name ="labField_UTME"
                            Caption ="Field coords:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =420
                    Width =960
                    Height =252
                    TabIndex =13
                    Name ="txtField_UTMN"
                    ControlSource ="Field_UTMN"
                    StatusBarText ="Field UTM Northing (zone 10N), in meters"
                    FontName ="Arial"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3300
                    Top =420
                    Width =960
                    Height =252
                    TabIndex =14
                    ColumnInfo ="\"\";\"\";\"10\";\"10\""
                    Name ="cmbField_datum"
                    ControlSource ="Field_datum"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Datum"
                    StatusBarText ="Field datum"
                    DefaultValue ="=[Forms]![frm_Switchboard]![cDatum]"
                    FontName ="Arial"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5040
                    Left =11400
                    Top =420
                    Width =1200
                    Height =252
                    TabIndex =18
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"24\""
                    Name ="cmbField_coord_source"
                    ControlSource ="Field_coord_source"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Coord_Source.Coord_source, tlu_Coord_Source.Coord_source_desc FROM tl"
                        "u_Coord_Source ORDER BY tlu_Coord_Source.Sort_order; "
                    ColumnWidths ="1440;3600"
                    StatusBarText ="Source of coordinate data"
                    DefaultValue ="\"GPS\""
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9780
                            Top =420
                            Width =1557
                            Height =255
                            Name ="labField_coord_source"
                            Caption ="Field coord. source:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4320
                    Top =420
                    Width =660
                    Height =252
                    TabIndex =15
                    Name ="txtField_horiz_error"
                    ControlSource ="Field_horiz_error"
                    StatusBarText ="GPS estimated horizontal accuracy, in same units as coordinate system"
                    FontName ="Arial"

                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9000
                    Top =420
                    Width =660
                    Height =252
                    TabIndex =17
                    Name ="txtField_offset_m"
                    ControlSource ="Field_offset_m"
                    StatusBarText ="Distance in meters from the coordinates to the target"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7440
                            Top =420
                            Width =1530
                            Height =255
                            Name ="labField_offset_m"
                            Caption ="Offset distance (m):"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =6720
                    Top =420
                    Width =600
                    Height =252
                    TabIndex =16
                    Name ="txtField_offset_azimuth"
                    ControlSource ="Field_offset_azimuth"
                    StatusBarText ="Azimuth (degrees, declination corrected) from the coordinates to the target"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5100
                            Top =420
                            Width =1620
                            Height =255
                            Name ="labField_offset_azimuth"
                            Caption ="Offset azimuth (deg):"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7320
                    Top =60
                    Width =1200
                    Height =252
                    TabIndex =7
                    Name ="txtCoord_label"
                    ControlSource ="Coord_label"
                    StatusBarText ="Name of the coordinate feature (e.g., plot center, NW corner)"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12840
                    Top =60
                    Width =900
                    Height =255
                    TabIndex =9
                    Name ="txtElevation_m"
                    ControlSource ="Elevation_m"
                    StatusBarText ="Elevation in meters, derived from GIS using final UTMs"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14880
                    Top =60
                    Width =900
                    Height =255
                    TabIndex =11
                    Name ="txtSlope_deg"
                    ControlSource ="Slope_deg"
                    StatusBarText ="Slope steepness in degrees, derived from GIS using final UTMs"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13860
                    Top =60
                    Width =900
                    Height =255
                    TabIndex =10
                    Name ="txtAspect_deg"
                    ControlSource ="Aspect_deg"
                    StatusBarText ="Slope aspect in degrees, derived from GIS using final UTMs"
                    FontName ="Arial"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3600
                    Left =15900
                    Top =60
                    Width =2639
                    Height =252
                    ColumnWidth =2568
                    TabIndex =21
                    Name ="cmbCoord_updated_by"
                    ControlSource ="Coord_updated_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active],'Active','') AS Is_a"
                        "ctive FROM tlu_Project_Crew ORDER BY IIf([Contact_is_active],'Active','') DESC ,"
                        " tlu_Project_Crew.Contact_ID; "
                    ColumnWidths ="2592;1008"
                    StatusBarText ="Person who made the most recent edits"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20820
                    Top =60
                    Width =2037
                    Height =252
                    ColumnWidth =1896
                    TabIndex =23
                    Name ="txtCoord_created_date"
                    ControlSource ="Coord_created_date"
                    Format ="yyyy mmm dd hh:nn"
                    StatusBarText ="Time stamp for record creation"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =18660
                    Top =60
                    Width =2037
                    Height =252
                    TabIndex =22
                    Name ="txtCoord_updated"
                    ControlSource ="Coord_updated"
                    Format ="yyyy mmm dd hh:nn"
                    StatusBarText ="Date of the last update to this record"
                    FontName ="Arial"

                End
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19500
                    Top =420
                    Width =3360
                    Height =252
                    ColumnWidth =1440
                    TabIndex =24
                    Name ="txtCoord_ID"
                    ControlSource ="Coord_ID"
                    StatusBarText ="Unique identifier for each coordinate record"
                    FontName ="Arial"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =18660
                    Top =420
                    Width =804
                    Height =240
                    Name ="labCoord_ID"
                    Caption ="Coord_ID:"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
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
' FORM NAME:    fsub_Coordinates_Browser
' Description:  Standard data browser subform for viewing and editing event coordinate records
' Data source:  In-line SQL statement based on tbl_Coordinates and tbl_Events
' Data access:  edit only
' Pages:        none
' Functions:    none
' References:   fxnGUIDGen, fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, July 2008
' Revisions:    JRB, 10/17/2008 - updated bounding box text
' =================================

Private Sub Form_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Note: this event is ignored on inserting a new record if BeforeInsert code exists

    ' Check if the current event record in the parent form is certified
    If IsNull(Me.Certified_date) = False And (IsNull(Me.Updated_date) _
        Or Me.Certified_date >= Me.Updated_date) Then

        Select Case Forms!frm_Switchboard!cAppMode
          Case "admin", "power user"
            ' Request confirmation before allowing edits
            If MsgBox("This record is certified ... are you certain you want to edit it?", _
                vbYesNo + vbExclamation + vbDefaultButton2, _
                "Confirm certified data edit") = vbNo Then
                DoCmd.CancelEvent
                GoTo Exit_Procedure
            End If
            MsgBox "Please log the edits to certified data ..."
            DoCmd.OpenForm "frm_Edit_Log", , , , , , "Update tbl_Coordinates"

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

    ElseIf Forms!frm_Switchboard!cAppMode <> "admin" And _
        Forms!frm_Switchboard!cAppMode <> "power user" Then
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
' SUB:          Form_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Validate the record before updating

    ' Hard validation rules:
    If IsNull(Me.txtUTM_east) <> IsNull(Me.txtUTM_north) Then
        MsgBox "Please enter a complete set of UTM coordinates", vbOKOnly, "Validation error"
        Me.txtUTM_east.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtUTM_east) = False And IsNull(Me.cmbDatum) Then
        MsgBox "Please enter the datum", vbOKOnly, "Validation error"
        Me.cmbDatum.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtUTM_east) = False And IsNull(Me.cmbCoord_type) Then
        MsgBox "Please enter the coordinate type", vbOKOnly, "Validation error"
        Me.cmbCoord_type.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtUTM_east) And IsNull(Me.cmbCoord_type) = False Then
        MsgBox "Either the UTM coordinates should be entered" & vbCrLf & _
            "or coordinate type should be blank", vbOKOnly, "Validation error"
        Me.cmbCoord_type.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtUTM_east) And IsNull(Me.txtEst_horiz_error) = False Then
        MsgBox "Either the UTM coordinates should be entered" & vbCrLf & _
            "or the estimated horizontal error should be blank", vbOKOnly, "Validation error"
        Me.txtEst_horiz_error.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtField_UTME) <> IsNull(Me.txtField_UTMN) Then
        MsgBox "Please enter a complete set of field coordinates", vbOKOnly, "Validation error"
        Me.txtField_UTME.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtField_horiz_error) = False And Me.cmbField_coord_source <> "GPS" Then
        MsgBox "Either the field source should be GPS" & vbCrLf & _
            "or the GPS error should be blank.", vbOKOnly, "Validation error"
        Me.cmbField_coord_source.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf (IsNull(Me.txtGPS_file_name) = False Or IsNull(Me.cmbGPS_model) = False) And _
        Me.cmbField_coord_source <> "GPS" Then
        MsgBox "Either the field source should be GPS or" & vbCrLf & _
            "the GPS file name and model should be blank.", vbOKOnly, "Validation error"
        Me.cmbField_coord_source.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtField_offset_azimuth) <> IsNull(Me.txtField_offset_m) Then
        MsgBox "Please enter a complete offset: distance and azimuth", vbOKOnly, _
            "Validation error"
        Me.txtField_offset_azimuth.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtField_UTME) = False And IsNull(Me.txtField_UTMN) = False Then
        If IsNull(Me.cmbField_coord_source) Then
            MsgBox "Please enter the field coordinate source", vbOKOnly, "Validation error"
            Me.cmbField_coord_source.SetFocus
            DoCmd.CancelEvent
            GoTo Exit_Procedure
        ElseIf IsNull(Me.cmbField_datum) Then
            MsgBox "Please enter the field datum", vbOKOnly, "Validation error"
            Me.cmbField_datum.SetFocus
            DoCmd.CancelEvent
            GoTo Exit_Procedure
        End If
  
        ' Test coordinates for bounding box
        Dim strPark As String
        Dim varUTME As Variant
        Dim varUTMN As Variant
        strPark = Me.Parent!Park_code
        varUTME = Me.txtField_UTME
        varUTMN = Me.txtField_UTMN
        Select Case strPark
          Case "OLYM"
            If varUTME < 370000 Or varUTME > 492000 Or _
                varUTMN < 5256000 Or varUTMN > 5349000 Then
                MsgBox "Coordinates are outside the bounding box for " _
                    & strPark, vbOKOnly, "Validation error"
                Me.txtField_UTME.SetFocus
                DoCmd.CancelEvent
                GoTo Exit_Procedure
            End If
          Case "NOCA"
            If varUTME < 599000 Or varUTME > 683000 Or _
                varUTMN < 5345000 Or varUTMN > 5430000 Then
                MsgBox "Coordinates are outside the bounding box for " _
                    & strPark, vbOKOnly, "Validation error"
                Me.txtField_UTME.SetFocus
                DoCmd.CancelEvent
                GoTo Exit_Procedure
            End If
          Case "MORA"
            If varUTME < 566000 Or varUTME > 620000 Or _
               varUTMN < 5167000 Or varUTMN > 5208000 Then
                MsgBox "Coordinates are outside the bounding box for " _
                    & strPark, vbOKOnly, "Validation error"
                Me.txtField_UTME.SetFocus
                DoCmd.CancelEvent
                GoTo Exit_Procedure
            End If
          Case "LEWI"
            If varUTME < 416000 Or varUTME > 435000 Or _
                varUTMN < 5084000 Or varUTMN > 5130000 Then
                MsgBox "Coordinates are outside the bounding box for " _
                    & strPark, vbOKOnly, "Validation error"
                Me.txtField_UTME.SetFocus
                DoCmd.CancelEvent
                GoTo Exit_Procedure
            End If
          Case "SAJH"
            If varUTME < 488000 Or varUTME > 502500 Or _
                varUTMN < 5366200 Or varUTMN > 5383000 Then
                MsgBox "Coordinates are outside the bounding box for " _
                    & strPark, vbOKOnly, "Validation error"
                Me.txtField_UTME.SetFocus
                DoCmd.CancelEvent
                GoTo Exit_Procedure
            End If
          Case "EBLA"
            If varUTME < 516000 Or varUTME > 529000 Or _
                varUTMN < 5332500 Or varUTMN > 5346000 Then
                MsgBox "Coordinates are outside the bounding box for " _
                    & strPark, vbOKOnly, "Validation error"
                Me.txtField_UTME.SetFocus
                DoCmd.CancelEvent
                GoTo Exit_Procedure
            End If
          Case "FOVA"
            If varUTME < 525100 Or varUTME > 527240 Or _
                varUTMN < 5051300 Or varUTMN > 5053000 Then
                MsgBox "Coordinates are outside the bounding box for " _
                    & strPark, vbOKOnly, "Validation error"
                Me.txtField_UTME.SetFocus
                DoCmd.CancelEvent
                GoTo Exit_Procedure
            End If
          Case Else
           ' Nothing
            MsgBox "No bounding box available for this park:" & _
                vbCrLf & strPark & vbCrLf & vbCrLf & _
                "Unable to verify that coordinates are in or near the park."
        End Select
    Else
        If Not IsNull(Me.Field_horiz_error) Or Not IsNull(Me.Field_offset_azimuth) Or _
            Not IsNull(Me.Field_offset_m) Or Not IsNull(Me.GPS_file_name) Then
            If MsgBox("Field coordinate information has been entered but field coordinates" & _
                vbCrLf & "are missing ... OK to continue, CANCEL to enter coordinates.", _
                vbOKCancel, "Missing field coordinates") = vbCancel Then
                Me.txtField_UTME.SetFocus
                DoCmd.CancelEvent
                GoTo Exit_Procedure
            End If
        End If
    End If

    ' Prior to saving, include a timestamp for edits
    '   NOTE: not updating the events table to enable switching "best" coordinates and other
    '       information without triggering updates
    If Me.NewRecord = False Then Me!Coord_updated = Now()
    If fxnSwitchboardIsOpen Then
        If IsNull(TempVars.item("UserAccessLevel")) = False Then
            Me!Coord_updated_by = TempVars.item("UserAccessLevel")
        Else
            Me!Coord_updated_by = Environ("Username")
        End If
    Else
        Me!Coord_updated_by = Environ("Username")
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
