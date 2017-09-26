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
    Width =9360
    DatasheetFontHeight =10
    ItemSuffix =11
    Left =900
    Top =450
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe1cb1c939781e340
    End
    RecordSource ="qrpt_Navigation_target_coordinates_all"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xd0020000d0020000d0020000d00200000000000090240000f000000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =240
            Name ="ReportHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =3720
                    Width =1260
                    Height =240
                    FontWeight =700
                    Name ="labNav_source"
                    Caption ="Coord source"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =240
                    Width =516
                    Height =228
                    FontWeight =700
                    Name ="labUTME"
                    Caption ="UTME"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =1260
                    Width =540
                    Height =228
                    FontWeight =700
                    Name ="labUTMN"
                    Caption ="UTMN"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =2040
                    Width =720
                    Height =228
                    FontWeight =700
                    Name ="labDatum"
                    Caption ="Datum"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =5160
                    Width =1056
                    Height =228
                    FontWeight =700
                    Name ="labCoord_type"
                    Caption ="Coord_type"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =6720
                    Width =840
                    Height =225
                    FontWeight =700
                    Name ="labElevation_m"
                    Caption ="Elev. (m)"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =7560
                    Width =1020
                    Height =225
                    FontWeight =700
                    Name ="labSlope_deg"
                    Caption ="Slope (deg)"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =8580
                    Width =720
                    Height =225
                    FontWeight =700
                    Name ="labAspect_deg"
                    Caption ="Aspect"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =2820
                    Width =840
                    Height =225
                    FontWeight =700
                    Name ="labEst_horiz_error"
                    Caption ="Error (m)"
                    FontName ="Tahoma"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanShrink = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Width =900
                    Height =228
                    TabIndex =1
                    Name ="txtUTM_east"
                    ControlSource ="UTME"
                    Format ="Fixed"
                    StatusBarText ="Final UTM easting (zone 10N, meters), including any offsets and corrections"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1020
                    Width =960
                    Height =228
                    TabIndex =2
                    Name ="txtUTM_north"
                    ControlSource ="UTMN"
                    Format ="Fixed"
                    StatusBarText ="Final UTM northing (zone 10N, meters), including any offsets and corrections"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3900
                    Width =780
                    Height =228
                    Name ="txtNav_source"
                    ControlSource ="Nav_source"
                    Format ="Yes/No"
                    StatusBarText ="Indicates whether this set of coordinates is the best available for this locatio"
                        "n"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2040
                    Width =720
                    Height =228
                    TabIndex =4
                    Name ="txtDatum"
                    ControlSource ="Nav_datum"
                    StatusBarText ="Datum of UTM_east and UTM_north"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4860
                    Width =1800
                    Height =228
                    TabIndex =3
                    Name ="txtCoord_type"
                    ControlSource ="Coord_type"
                    StatusBarText ="Coordinate type stored in UTM_east and UTM_north: target, field, post-processed"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6720
                    Width =840
                    TabIndex =5
                    Name ="Elevation_m"
                    ControlSource ="Elevation_m"
                    Format ="Fixed"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7680
                    Width =720
                    TabIndex =6
                    Name ="txtSlope_deg"
                    ControlSource ="Slope_deg"
                    Format ="Fixed"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8580
                    Width =720
                    TabIndex =7
                    Name ="txtAspect_deg"
                    ControlSource ="Aspect_deg"
                    Format ="Fixed"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2940
                    Width =540
                    TabIndex =8
                    Name ="txtEst_horiz_error"
                    ControlSource ="=IIf([Est_horiz_error]<>-99,[Est_horiz_error])"
                    FontName ="Tahoma"

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
