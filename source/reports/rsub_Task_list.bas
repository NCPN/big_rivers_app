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
    Width =11520
    DatasheetFontHeight =10
    ItemSuffix =33
    Left =735
    Top =255
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf00acddc5624e340
    End
    RecordSource ="qrsub_Task_list"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000020010000680100002001000000000000002d0000f000000001000000 ,
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
            ControlSource ="Request_date"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =240
            Name ="ReportHeader"
            Begin
                Begin Label
                    TextFontFamily =34
                    Left =60
                    Width =1620
                    Height =225
                    FontSize =8
                    Name ="labRequest_date"
                    Caption ="Task request date"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontFamily =34
                    Left =8580
                    Width =1500
                    Height =228
                    FontSize =8
                    Name ="labRequested_by"
                    Caption ="Request by"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontFamily =34
                    Left =10140
                    Width =660
                    Height =228
                    FontSize =8
                    Name ="labTask_status"
                    Caption ="Status"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontFamily =34
                    Left =1860
                    Width =1680
                    Height =228
                    FontSize =8
                    Name ="labTask_desc"
                    Caption ="Task description"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontFamily =34
                    Left =11280
                    Width =174
                    Height =228
                    FontSize =8
                    Name ="labLocation_ID"
                    Caption ="Location_ID"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontFamily =34
                    Left =4920
                    Width =2160
                    Height =228
                    FontSize =8
                    Name ="labTask_notes"
                    Caption ="Notes"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1320
                    Width =3540
                    TabIndex =2
                    Name ="txtTask_desc"
                    ControlSource ="Task_desc"
                    StatusBarText ="Brief description of the feature"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8580
                    Width =1500
                    TabIndex =3
                    Name ="txtRequested_by"
                    ControlSource ="Requested_by"
                    StatusBarText ="Distance in meters, measured from the previous point for travel features"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Width =1020
                    ColumnWidth =1740
                    TabIndex =1
                    Name ="txtRequest_date"
                    ControlSource ="Request_date"
                    Format ="Short Date"
                    StatusBarText ="Type of feature"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10140
                    Width =1020
                    TabIndex =4
                    Name ="txtTask_status"
                    ControlSource ="Task_status"
                    StatusBarText ="Current status of the feature"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11220
                    Width =240
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="Sample location"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4920
                    Width =3600
                    TabIndex =5
                    Name ="txtTask_notes"
                    ControlSource ="Task_notes"
                    StatusBarText ="Brief description of the feature"
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
