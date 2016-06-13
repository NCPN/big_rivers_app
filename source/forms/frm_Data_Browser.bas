Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
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
    Width =11655
    DatasheetFontHeight =10
    ItemSuffix =88
    Left =6060
    Top =2250
    Right =17715
    Bottom =12630
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8e5cc12ab6c4e440
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT tbl_Sites.Park_code AS Site_park, tbl_Sites.Site_code, tbl_Sites.Park_reg"
        "ion, tbl_Sites.Stratum_ID, tbl_Sites.Panel_type, tbl_Sites.Panel_name, tbl_Sites"
        ".Site_status, tbl_Locations.* FROM tbl_Sites RIGHT JOIN tbl_Locations ON tbl_Sit"
        "es.Site_ID = tbl_Locations.Site_ID ORDER BY tbl_Locations.Park_code, tbl_Sites.S"
        "ite_code, tbl_Locations.Location_code; "
    Caption =" Project Location Data Browser"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnUndo ="[Event Procedure]"
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
            Height =1320
            BackColor =13692912
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =10860
                    Top =60
                    Width =720
                    Height =294
                    FontWeight =700
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9900
                    Top =60
                    Width =691
                    Height =294
                    FontWeight =700
                    TabIndex =22
                    Name ="cmdNew"
                    Caption ="New"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Add a new sampling location record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9120
                    Top =60
                    Width =691
                    Height =294
                    FontWeight =700
                    TabIndex =21
                    Name ="cmdDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =8340
                    Top =60
                    Width =691
                    Height =294
                    FontWeight =700
                    TabIndex =20
                    Name ="cmdSave"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =9780
                    Top =564
                    Width =1080
                    TabIndex =9
                    Name ="cmbStatusFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Site_Status.Site_status, tlu_Site_Status.Site_status_desc FROM tlu_Si"
                        "te_Status ORDER BY tlu_Site_Status.Sort_order; "
                    ColumnWidths ="1008;3312"
                    StatusBarText ="Filter by location status"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by location status"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9000
                            Top =564
                            Width =720
                            Height =225
                            Name ="labStatusFilter"
                            Caption ="Status:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =660
                    Top =564
                    Width =840
                    TabIndex =1
                    Name ="cmbParkFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Park_code FROM tbl_Locations GROUP BY tbl_Locations.Park_co"
                        "de ORDER BY tbl_Locations.Park_code; "
                    StatusBarText ="Filter by park"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    OnGotFocus ="[Event Procedure]"
                    ControlTipText ="Filter by park"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =564
                            Width =540
                            Height =228
                            FontWeight =700
                            Name ="labParkFilter"
                            Caption ="Park:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =1560
                    Top =540
                    Width =480
                    Height =300
                    TabIndex =2
                    Name ="togFilterByPark"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the park filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7260
                    Top =564
                    Width =1080
                    TabIndex =7
                    Name ="cmbTypeFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Location_Type.Location_type FROM tlu_Location_Type ORDER BY tlu_Locat"
                        "ion_Type.Sort_order; "
                    StatusBarText ="Filter by location type"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by location type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6300
                            Top =564
                            Width =900
                            Height =228
                            Name ="labTypeFilter"
                            Caption ="Loc type:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =8400
                    Top =540
                    Width =480
                    Height =300
                    TabIndex =8
                    Name ="togFilterByType"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the location type filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =20
                    ListWidth =2304
                    Left =2700
                    Top =564
                    Width =936
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cmbSiteFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Sites.Site_ID, tbl_Sites.Site_code, tbl_Sites.Site_status FROM tbl_Si"
                        "tes WHERE (((tbl_Sites.Park_code) Like IIf(Abs([Forms]![frm_Data_Browser]![togFi"
                        "lterByPark])=1,Nz([Forms]![frm_Data_Browser]![cmbParkFilter],\"*\"),\"*\")))  UN"
                        "ION SELECT 'NA' AS Site_ID, '(blank)' AS Site_code, Null As Site_status FROM tbl"
                        "_Sites ORDER BY tbl_Sites.Site_status, tbl_Sites.Site_code;"
                    ColumnWidths ="0;720;1584"
                    StatusBarText ="Filter by site"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    OnNotInList ="[Event Procedure]"
                    ControlTipText ="Filter by site"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2100
                            Top =564
                            Width =540
                            Height =228
                            Name ="labSiteFilter"
                            Caption ="Site:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =3660
                    Top =540
                    Width =480
                    Height =300
                    TabIndex =4
                    Name ="togFilterBySite"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the site filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =10920
                    Top =540
                    Width =480
                    Height =300
                    TabIndex =10
                    Name ="togFilterByStatus"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the status filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListRows =20
                    Left =3900
                    Top =960
                    Width =1560
                    Height =255
                    TabIndex =13
                    Name ="cmbRegionFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Sites.Park_region FROM tbl_Sites WHERE (((tbl_Sites.Park_code) Like N"
                        "z([Forms]![frm_Data_Browser]![cmbParkFilter],\"*\"))) GROUP BY tbl_Sites.Park_re"
                        "gion ORDER BY tbl_Sites.Park_region; "
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by park region"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3120
                            Top =960
                            Width =720
                            Height =255
                            Name ="labRegionFilter"
                            Caption ="Region:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =20
                    ListWidth =2880
                    Left =900
                    Top =961
                    Width =1620
                    Height =255
                    TabIndex =11
                    ForeColor =0
                    Name ="cmbStratumFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Strata.Stratum_ID, [Park_code] & \" - \" & [Stratum_name] AS Stratum,"
                        " tbl_Strata.Stratification_date FROM tbl_Strata WHERE (((tbl_Strata.Park_code) L"
                        "ike Nz([Forms]![frm_Data_Browser]![cmbParkFilter],\"*\"))) ORDER BY tbl_Strata.P"
                        "ark_code, tbl_Strata.Stratum_name; "
                    ColumnWidths ="0;1728;1152"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by stratum"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =961
                            Width =780
                            Height =255
                            Name ="labStratumFilter"
                            Caption ="Stratum:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =2580
                    Top =937
                    Width =480
                    Height =300
                    TabIndex =12
                    Name ="togFilterByStratum"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Filter by the selected stratum"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =5520
                    Top =937
                    Width =480
                    Height =300
                    TabIndex =14
                    Name ="togFilterByRegion"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Filter by the selected park region"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =9360
                    Left =7080
                    Top =960
                    Width =1200
                    Height =255
                    TabIndex =15
                    Name ="cmbPanelTypeFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Panel_Type.Panel_type, tlu_Panel_Type.Panel_type_desc FROM tlu_Panel_"
                        "Type ORDER BY tlu_Panel_Type.Sort_order; "
                    ColumnWidths ="1152;8208"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by sampling panel type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6060
                            Top =960
                            Width =960
                            Height =255
                            Name ="labPanelTypeFilter"
                            Caption ="Panel type:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =8340
                    Top =937
                    Width =480
                    Height =300
                    TabIndex =16
                    Name ="togFilterByPanelType"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Filter by the selected panel type"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =10980
                    Top =937
                    Width =480
                    Height =300
                    TabIndex =18
                    Name ="togFilterByPanelName"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Filter by the selected panel name"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10020
                    Top =960
                    Width =900
                    Height =270
                    TabIndex =17
                    Name ="cmbPanelNameFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Sites.Panel_name FROM tbl_Sites GROUP BY tbl_Sites.Panel_name ORDER B"
                        "Y tbl_Sites.Panel_name; "
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by panel name"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8880
                            Top =960
                            Width =1080
                            Height =255
                            Name ="labPanelNameFilter"
                            Caption ="Panel name:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6840
                    Top =60
                    Width =1201
                    Height =294
                    FontWeight =700
                    TabIndex =19
                    Name ="cmdFiltersOff"
                    Caption ="Filters Off"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Turn off all form filters"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListRows =20
                    ListWidth =4032
                    Left =4800
                    Top =564
                    Width =936
                    TabIndex =5
                    Name ="cmbLocFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Location_ID, tbl_Locations.Location_code, [tbl_Locations].["
                        "Park_code] & '.' & [Site_code] AS Site, tbl_Locations.Location_type, tbl_Locatio"
                        "ns.Location_status FROM tbl_Sites RIGHT JOIN tbl_Locations ON tbl_Sites.Site_ID "
                        "= tbl_Locations.Site_ID WHERE (((tbl_Locations.Park_code) Like Nz([Forms]![frm_D"
                        "ata_Browser]![cmbParkFilter],\"*\")) AND ((Nz([tbl_Locations].[Site_ID],\"*\")) "
                        "Like Nz([Forms]![frm_Data_Browser]![cmbSiteFilter],\"*\"))) ORDER BY tbl_Locatio"
                        "ns.Location_status, tbl_Sites.Site_code, tbl_Locations.Location_code; "
                    ColumnWidths ="0;720;1152;1008;1152"
                    StatusBarText ="Filter by sample location"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    OnNotInList ="[Event Procedure]"
                    ControlTipText ="Filter by sample location"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4200
                            Top =564
                            Width =540
                            Height =228
                            Name ="labLocFilter"
                            Caption ="Loc:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =5760
                    Top =540
                    Width =480
                    Height =300
                    TabIndex =6
                    Name ="togFilterByLoc"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the sample location filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            SpecialEffect =2
            Height =8655
            BackColor =13025979
            Name ="Detail"
            Begin
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5580
                    Top =120
                    Width =972
                    Height =252
                    ColumnWidth =972
                    TabIndex =2
                    Name ="txtLocation_code"
                    ControlSource ="Location_code"
                    StatusBarText ="Alphanumeric code used to identify the sample location"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4200
                            Top =120
                            Width =1320
                            Height =255
                            FontWeight =700
                            Name ="labLocation_code"
                            Caption ="Location code"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5880
                    Top =480
                    Width =4200
                    Height =252
                    ColumnWidth =2568
                    TabIndex =7
                    Name ="txtLocation_name"
                    ControlSource ="Location_name"
                    StatusBarText ="Brief colloquial name of the sample location (optional)"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4620
                            Top =480
                            Width =1206
                            Height =252
                            Name ="labLocation_name"
                            Caption ="Location name"
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
                    Left =1080
                    Top =120
                    Width =960
                    Height =252
                    ColumnWidth =2568
                    Name ="cmbPark_code"
                    ControlSource ="Park_code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.Park_code FROM tlu_Parks ORDER BY tlu_Parks.Park_code; "
                    StatusBarText ="Park code"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =120
                            Width =957
                            Height =252
                            FontWeight =700
                            Name ="labPark_code"
                            Caption ="Park code"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =6480
                    Left =7980
                    Top =120
                    Width =1560
                    Height =252
                    ColumnWidth =2568
                    TabIndex =3
                    Name ="cmbLocation_type"
                    ControlSource ="Location_type"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Location_Type.Location_type, tlu_Location_Type.Loc_type_desc FROM tlu"
                        "_Location_Type ORDER BY tlu_Location_Type.Sort_order; "
                    ColumnWidths ="1008;5472"
                    StatusBarText ="Indicates the type of sample location"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6720
                            Top =120
                            Width =1200
                            Height =252
                            FontWeight =700
                            Name ="labLocation_type"
                            Caption ="Location type"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =7200
                    Left =10440
                    Top =120
                    Width =1074
                    Height =252
                    ColumnWidth =2568
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"20\""
                    Name ="cmbLocation_status"
                    ControlSource ="Location_status"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Site_Status.Site_status, tlu_Site_Status.Site_status_desc FROM tlu_Si"
                        "te_Status ORDER BY tlu_Site_Status.Sort_order; "
                    ColumnWidths ="1008;6192"
                    StatusBarText ="Status of the sample location"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9720
                            Top =120
                            Width =660
                            Height =252
                            FontWeight =700
                            Name ="labLocation_status"
                            Caption ="Status"
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
                    Left =3240
                    Top =1200
                    Width =8280
                    Height =720
                    ColumnWidth =3000
                    TabIndex =13
                    Name ="txtLocation_notes"
                    ControlSource ="Location_notes"
                    StatusBarText ="Notes about the sample location"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =1200
                            Width =474
                            Height =252
                            Name ="labLocation_notes"
                            Caption ="Notes"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =1260
                    Width =1224
                    Height =252
                    ColumnWidth =1896
                    TabIndex =11
                    Name ="txtLoc_established"
                    ControlSource ="Loc_established"
                    Format ="yyyy mmm dd"
                    StatusBarText ="Date the sample location was established"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =1260
                            Width =1026
                            Height =252
                            Name ="labLoc_established"
                            Caption ="Established"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =1620
                    Width =1224
                    Height =252
                    ColumnWidth =1896
                    TabIndex =12
                    Name ="txtLoc_discontinued"
                    ControlSource ="Loc_discontinued"
                    Format ="yyyy mmm dd"
                    StatusBarText ="Date the sample location was discontinued"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =1620
                            Width =1023
                            Height =252
                            Name ="labLoc_discontinued"
                            Caption ="Discontinued"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =85
                    Top =3180
                    Width =11655
                    Height =5475
                    FontWeight =700
                    TabIndex =16
                    Name ="pgLocChildren"
                    FontName ="Arial"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =59
                            Top =3600
                            Width =11461
                            Height =4920
                            Name ="pgSite"
                            ControlTipText ="Information about the parent site record"
                            Caption =" Site Information"
                            LayoutCachedLeft =59
                            LayoutCachedTop =3600
                            LayoutCachedWidth =11520
                            LayoutCachedHeight =8520
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =59
                                    Top =3656
                                    Width =11460
                                    Height =4860
                                    Name ="subSite"
                                    SourceObject ="Form.fsub_Sites_Browser"
                                    LinkChildFields ="Site_ID"
                                    LinkMasterFields ="Site_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =59
                            Top =3600
                            Width =11461
                            Height =4920
                            Name ="pgSchedule"
                            Caption =" Schedule"
                            LayoutCachedLeft =59
                            LayoutCachedTop =3600
                            LayoutCachedWidth =11520
                            LayoutCachedHeight =8520
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =0
                                    Left =59
                                    Top =4016
                                    Width =11460
                                    Height =4500
                                    Name ="subSchedule"
                                    SourceObject ="Form.fsub_Schedule_Browser"
                                    LinkChildFields ="Site_ID"
                                    LinkMasterFields ="Site_ID"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =120
                                            Top =3720
                                            Width =1920
                                            Height =255
                                            Name ="labSchedule"
                                            Caption ="Scheduled years"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =6900
                                    Top =3600
                                    Width =2280
                                    Height =300
                                    TabIndex =1
                                    Name ="cmdScheduleForm"
                                    Caption ="Open the schedule browser"
                                    StatusBarText ="Open the schedule browser form"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open the schedule browser form"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =60
                            Top =3600
                            Width =11460
                            Height =4920
                            Name ="pgCoordinates"
                            Caption =" Coordinates"
                            LayoutCachedLeft =60
                            LayoutCachedTop =3600
                            LayoutCachedWidth =11520
                            LayoutCachedHeight =8520
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =3
                                    Left =60
                                    Top =3660
                                    Width =7920
                                    Height =1080
                                    Name ="subTarget_coords"
                                    SourceObject ="Form.fsub_Target_Coords_Browser"
                                    LinkChildFields ="Location_ID"
                                    LinkMasterFields ="Location_ID"

                                End
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =0
                                    Left =60
                                    Top =5040
                                    Width =11460
                                    Height =3480
                                    TabIndex =1
                                    Name ="subCoordinates"
                                    SourceObject ="Form.fsub_Coordinates_Browser"
                                    LinkChildFields ="Location_ID"
                                    LinkMasterFields ="Location_ID"

                                    Begin
                                        Begin Label
                                            FontUnderline = NotDefault
                                            OverlapFlags =223
                                            Left =120
                                            Top =4800
                                            Width =10650
                                            Height =240
                                            Name ="labField_coords"
                                            Caption ="Event coordinates - field coordinates collected during sampling events, and fina"
                                                "l coordinates derived from GPS, field coordinates, or target coords:"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =3
                                    OverlapFlags =223
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =9360
                                    Top =4020
                                    Width =900
                                    Height =255
                                    TabIndex =2
                                    Name ="txtUTME_public"
                                    ControlSource ="UTME_public"
                                    StatusBarText ="UTM easting (zone 10N, meters).  Note: in addition to any measurement error, the"
                                        "se coordinates may have been offset up to 2 km from their actual position."
                                    FontName ="Arial"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =8160
                                            Top =3660
                                            Width =3300
                                            Height =255
                                            Name ="labPublic_coords"
                                            Caption ="Public coords for reports, presentations, etc."
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =3
                                    OverlapFlags =223
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =10320
                                    Top =4020
                                    Width =1140
                                    Height =255
                                    TabIndex =3
                                    Name ="txtUTMN_public"
                                    ControlSource ="UTMN_public"
                                    StatusBarText ="UTM northing (zone 10N, meters).  Note: in addition to any measurement error, th"
                                        "ese coordinates may have been offset up to 2 km from their actual position."
                                    FontName ="Arial"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =8160
                                            Top =4020
                                            Width =1140
                                            Height =255
                                            Name ="labUTMN_public"
                                            Caption ="Public UTM E/N"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =3
                                    OverlapFlags =223
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =9360
                                    Top =4380
                                    Width =2100
                                    Height =255
                                    TabIndex =4
                                    Name ="txtPublic_offset"
                                    ControlSource ="Public_offset"
                                    StatusBarText ="Type of processing performed to make coordinates publishable"
                                    FontName ="Arial"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =3
                                            Left =8220
                                            Top =4380
                                            Width =1080
                                            Height =255
                                            Name ="labPublic_offset"
                                            Caption ="Public offset"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    BorderWidth =3
                                    OverlapFlags =247
                                    Left =8100
                                    Top =3660
                                    Width =3420
                                    Height =1080
                                    BackColor =13025979
                                    Name ="Box61"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =3600
                            Width =11460
                            Height =4920
                            Name ="pgTasks"
                            Caption =" Tasks"
                            LayoutCachedLeft =60
                            LayoutCachedTop =3600
                            LayoutCachedWidth =11520
                            LayoutCachedHeight =8520
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =0
                                    Left =60
                                    Top =3660
                                    Width =11460
                                    Height =4860
                                    Name ="subTasks"
                                    SourceObject ="Form.fsub_Task_List_Browser"
                                    LinkChildFields ="Location_ID"
                                    LinkMasterFields ="Location_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =59
                            Top =3600
                            Width =11461
                            Height =4920
                            Name ="pgEvents"
                            Caption =" Events"
                            LayoutCachedLeft =59
                            LayoutCachedTop =3600
                            LayoutCachedWidth =11520
                            LayoutCachedHeight =8520
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =0
                                    Left =59
                                    Top =3656
                                    Width =11460
                                    Height =4860
                                    Name ="subEvents"
                                    SourceObject ="Form.fsub_Events_Browser"
                                    LinkChildFields ="Location_ID"
                                    LinkMasterFields ="Location_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =3600
                            Width =11400
                            Height =4920
                            Name ="pgImages"
                            Caption =" Images"
                            LayoutCachedLeft =120
                            LayoutCachedTop =3600
                            LayoutCachedWidth =11520
                            LayoutCachedHeight =8520
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =120
                                    Top =3660
                                    Width =11400
                                    Height =4860
                                    Name ="subImages"
                                    SourceObject ="Form.fsub_Images_Browser"
                                    LinkChildFields ="Location_ID"
                                    LinkMasterFields ="Location_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =59
                            Top =3600
                            Width =11461
                            Height =4920
                            Name ="pgMarkers"
                            Caption =" Markers"
                            LayoutCachedLeft =59
                            LayoutCachedTop =3600
                            LayoutCachedWidth =11520
                            LayoutCachedHeight =8520
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =0
                                    Left =59
                                    Top =3656
                                    Width =11460
                                    Height =4860
                                    Name ="subMarkers"
                                    SourceObject ="Form.fsub_Markers"
                                    LinkChildFields ="Location_ID"
                                    LinkMasterFields ="Location_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =59
                            Top =3600
                            Width =11461
                            Height =4920
                            Name ="pgFeatures"
                            Caption =" Features"
                            LayoutCachedLeft =59
                            LayoutCachedTop =3600
                            LayoutCachedWidth =11520
                            LayoutCachedHeight =8520
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =0
                                    Left =59
                                    Top =3656
                                    Width =11460
                                    Height =4860
                                    Name ="subFeatures"
                                    SourceObject ="Form.fsub_Features_Browser"
                                    LinkChildFields ="Location_ID"
                                    LinkMasterFields ="Location_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =3600
                            Width =11460
                            Height =4920
                            Name ="pgAnalysis"
                            Caption =" Analysis Info"
                            LayoutCachedLeft =60
                            LayoutCachedTop =3600
                            LayoutCachedWidth =11520
                            LayoutCachedHeight =8520
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =0
                                    Left =60
                                    Top =3840
                                    Width =11460
                                    Height =2160
                                    Name ="subAnalysisNotes"
                                    SourceObject ="Form.fsub_Analysis_Notes"
                                    LinkChildFields ="Location_ID"
                                    LinkMasterFields ="Location_ID"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =60
                                            Top =3600
                                            Width =2760
                                            Height =255
                                            Name ="labAnalysisNotes"
                                            Caption ="Analysis notes by sample point"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =0
                                    Left =60
                                    Top =6420
                                    Width =11460
                                    Height =2100
                                    TabIndex =1
                                    Name ="subVarianceGroups"
                                    SourceObject ="Form.fsub_Variance_Groups"
                                    LinkChildFields ="Site_ID"
                                    LinkMasterFields ="Site_ID"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =60
                                            Top =6180
                                            Width =2760
                                            Height =255
                                            Name ="labVarianceGroups"
                                            Caption ="Variance groups by sample site"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    AllowAutoCorrect = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =2040
                    Width =4977
                    Height =1020
                    TabIndex =14
                    Name ="txtTravel_notes"
                    ControlSource ="Travel_notes"
                    StatusBarText ="Directions for relocating the sample location"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =2040
                            Width =660
                            Height =540
                            Name ="labTravel_notes"
                            Caption ="Travel notes"
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
                    Left =6780
                    Top =2040
                    Width =4752
                    Height =1020
                    TabIndex =15
                    Name ="txtLocation_desc"
                    ControlSource ="Location_desc"
                    StatusBarText ="Environmental description of the sampling location"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6000
                            Top =2040
                            Width =720
                            Height =600
                            Name ="labLocation_desc"
                            Caption ="Location desc"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =3720
                    Top =480
                    Width =480
                    Height =300
                    TabIndex =6
                    Name ="togFilterByTrailOrRoad"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Filter by the selected trail or road indicator"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2520
                    Top =480
                    Width =1140
                    TabIndex =5
                    Name ="cmbTrail_or_road"
                    ControlSource ="Trail_or_road"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Trail_Or_Road.Trail_code FROM tlu_Trail_Or_Road ORDER BY tlu_Trail_Or"
                        "_Road.Sort_order; "
                    StatusBarText ="Indicates whether or not the sample location is along a road or trail"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =1380
                            Top =480
                            Width =1050
                            Height =240
                            Name ="labTrail_or_road"
                            Caption ="Trail or road"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListRows =20
                    ListWidth =4608
                    Left =2940
                    Top =120
                    Width =1080
                    TabIndex =1
                    Name ="cmbSite_ID"
                    ControlSource ="Site_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Sites.Site_ID, tbl_Sites.Site_code, tbl_Sites.Park_code, tbl_Sites.Si"
                        "te_status, tbl_Sites.Site_name FROM tbl_Sites ORDER BY tbl_Sites.Site_code, tbl_"
                        "Sites.Site_status; "
                    ColumnWidths ="0;720;1008;1152;1728"
                    StatusBarText ="Site membership of the sample location (transect)"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2220
                            Top =120
                            Width =645
                            Height =240
                            FontWeight =700
                            Name ="labSite_ID"
                            Caption ="Site ID"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2520
                    Top =840
                    Width =660
                    TabIndex =8
                    Name ="txtAzimuth_to_point"
                    ControlSource ="Azimuth_to_point"
                    StatusBarText ="Azimuth (degrees, declination corrected) to the sampling point from the previous"
                        " point, to facilitate relocating the position; 999 signifies points along the tr"
                        "ail"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1140
                            Top =840
                            Width =1320
                            Height =240
                            Name ="labAzimuth_to_point"
                            Caption ="Azimuth to point"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4800
                    Top =840
                    Width =840
                    TabIndex =9
                    Name ="cmbDirection_changed"
                    ControlSource ="Direction_changed"
                    RowSourceType ="Value List"
                    RowSource ="Yes;No"
                    StatusBarText ="Indicates whether the azimuth to the point was changed to accommodate navigation"
                    FontName ="Arial"
                    Format ="Yes/No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3300
                            Top =840
                            Width =1440
                            Height =240
                            Name ="labDirection_changed"
                            Caption ="Direction changed"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7320
                    Top =840
                    Width =4200
                    TabIndex =10
                    Name ="txtReason_for_change"
                    ControlSource ="Reason_for_change"
                    StatusBarText ="Brief comments about why the azimuth was changed"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5760
                            Top =840
                            Width =1515
                            Height =240
                            Name ="labReason_for_change"
                            Caption ="Reason for change"
                            FontName ="Arial"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =420
            BackColor =11651021
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1020
                    Top =120
                    Width =3840
                    Height =252
                    ColumnWidth =1440
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="Unique identifier for each sample location"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =120
                            Width =900
                            Height =270
                            Name ="labLocation_ID"
                            Caption ="Location ID:"
                            FontName ="Arial"
                        End
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
                    Left =10140
                    Top =120
                    Height =252
                    TabIndex =3
                    Name ="txtLoc_created_date"
                    ControlSource ="Loc_created_date"
                    Format ="yyyy mmm dd hh:nn"
                    StatusBarText ="Time stamp when the record was created"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9120
                            Top =120
                            Width =957
                            Height =240
                            Name ="labLoc_created_date"
                            Caption ="Loc created"
                            FontName ="Arial"
                        End
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
                    Left =6060
                    Top =120
                    Height =252
                    TabIndex =1
                    Name ="txtLoc_updated"
                    ControlSource ="Loc_updated"
                    Format ="yyyy mmm dd hh:nn"
                    StatusBarText ="Date of the last update to this record"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4920
                            Top =120
                            Width =1065
                            Height =240
                            Name ="labLoc_updated"
                            Caption ="Updated / by:"
                        End
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
                    Left =7560
                    Top =120
                    Height =252
                    TabIndex =2
                    Name ="txtLoc_updated_by"
                    ControlSource ="Loc_updated_by"
                    StatusBarText ="Person who made the most recent edits"

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
' FORM NAME:    frm_Data_Browser
' Description:  Standard data browser form for viewing and editing project data by
'                   sample location
' Data source:  In-line SQL statement based on tbl_Sites and tbl_Locations
' Data access:  edit; add only by cmdNew; delete only by cmdDelete
' Pages:        pgSite, pgSchedule, pgCoordinates, pgTasks, pgEvents, pgImages, pgMarkers,
'                   pgFeatures, pgAnalysis
' Functions:    fxnFilterRecords
' References:   fxnGUIDGen, fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, July 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, 11/12/2008 - standardization; added location filter; updated
'                   fxnFilterRecords
'               JRB, 2/6/2009 - minor updates and fixes to filters; moved edit log call on
'                    delete from after delete confirm to cmdDelete
'               JRB, 5/1/2009 - updated Form_Open to allow opening to a blank record from
'                   QA Tool
'               JRB, 9/20/2010 - added pgImages and subImages
'               BLC, 6/13/2014 - updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
' =================================

' ---------------------------------
' SUB:     Form_Open
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 7/29/2014 - updated to use TempVars.Item("Park") vs. cPark
'               BLC, 8/25/2014 - shifted UI control/subform intialization to mod_User (setUserAccess)
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Close the form if the switchboard is not open
    If fxnSwitchboardIsOpen = False Then
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this form to function properly.", , "Cannot open the form ..."
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

    ' initialize controls/subforms based on user access level
    setUserAccess Me

    If Me.FilterOn Then
    ' If the form is filtered, set the site filter according to the filtered record
        Me.cmbParkFilter = Me.Park_code
        Me.togFilterByPark = True
        If Not IsNull(Me.Site_ID) Then
            Me.cmbSiteFilter = Me.Site_ID
            Me.togFilterBySite = True
        Else
            ' If no site, fill in so that the query for cmbLocFilter will show "rare"
            Me.cmbSiteFilter = "NA"
        End If
        If Me.OpenArgs = "Location_ID" Then
            Me.cmbLocFilter = Me.Location_ID
            Me.togFilterByLoc = True
        End If
        fxnFilterRecords (True)
    ElseIf Me.OpenArgs = "off" Then
        ' Default when opening from QA Tool - to distinguish from opening directly to a record
        ' Turn filter toggles off
        Me.togFilterByPark = False
        Me.togFilterBySite = False
        Me.togFilterByLoc = False
        Me.togFilterByStatus = False
        Me.togFilterByType = False
        Me.togFilterByRegion = False
        Me.togFilterByStratum = False
        Me.togFilterByPanelType = False
        Me.togFilterByPanelName = False
        ' ... including those embedded in the record
        Me.togFilterByTrailOrRoad = False
        fxnFilterRecords
        Me.cmdFiltersOff.Enabled = False
        Me.cmbParkFilter.SetFocus
    Else
    ' Set the default form filter
        If fxnSwitchboardIsOpen Then Me.cmbParkFilter = TempVars.item("Park")
        Me.togFilterByPark = True
        Me.cmbTypeFilter = "Origin"
        Me.togFilterByType = True
        Me.cmbStatusFilter = "Active"
        Me.togFilterByStatus = True
        fxnFilterRecords
        Me.cmdClose.SetFocus
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2455 ' Invalid reference
        MsgBox "Unable to open the form - no records meet filter criteria", , "No record found"
        Cancel = True
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_Current
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
' ---------------------------------
Private Sub Form_Current()
    On Error GoTo Err_Handler

    If Me.NewRecord Then
        If Me.AllowAdditions = False Then Me.AllowAdditions = True
        Me.cmdDelete.Caption = "Undo"
        Me.cmdDelete.Enabled = False
        ' Enable the 1-1 subforms (otherwise they appear blank, no way to add child records)
        Me.subSite.Form.AllowAdditions = True
        Me.subTarget_coords.Form.AllowAdditions = True
    Else
        ' When moving to a different record because a user has selected a different record
        '   on a referring form, update the filter ctls
        If Me.cmbParkFilter <> Me.Park_code And Me.togFilterByPark Then
            Me.cmbParkFilter = Me.Park_code
            Me.cmbParkFilter.Requery
        End If
        If Me.cmbSiteFilter <> Me.Site_ID And Me.togFilterBySite Then
            Me.cmbSiteFilter = Me.Site_ID
            Me.cmbSiteFilter.Requery
        End If
        If Me.cmbLocFilter <> Me.Location_ID And Me.togFilterByLoc Then
            Me.cmbLocFilter = Me.Location_ID
            Me.cmbLocFilter.Requery
        End If

        Me.AllowAdditions = False
        Me.cmdDelete.Caption = "Delete"

        ' Limit the following to edit mode only
        If TempVars.item("UserAccessLevel") = "admin" Or _
            TempVars.item("UserAccessLevel") = "power user" Then
            ' Enable the 1-1 subforms if no child records exist
            '   (otherwise blank form, no way for the user to add)
            If DCount("*", "tbl_Target_Coords", _
                "[Location_ID]=""" & Me.Location_ID & """") = 0 Then _
                Me.subTarget_coords.Form.AllowAdditions = True
            Me.cmdDelete.Enabled = True
        Else
            Me.cmdDelete.Enabled = False
        End If
        ' Save button only enabled upon dirty/insert
        Me.cmdSave.Enabled = False
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_Dirty
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
' ---------------------------------
Private Sub Form_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    Select Case TempVars.item("UserAccessLevel")
      Case "admin", "power user"
        ' allow edits to the main form
      Case Else
        ' do not allow edits to the main form
        Cancel = True
        GoTo Exit_Procedure
    End Select

    Me.cmdDelete.Caption = "Undo"
    Me.cmdDelete.Enabled = True
    Me.cmdSave.Enabled = True

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_BeforeInsert
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Create the GUID primary key value
    Me.Location_ID = fxnGUIDGen
    Me.Location_status = "Active"
    Me.cmdDelete.Enabled = True
    Me.cmdSave.Enabled = True

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_Undo
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub Form_Undo(Cancel As Integer)
    On Error GoTo Err_Handler

    Me.cmdSave.Enabled = False
    Me.cmdDelete.Caption = "Delete"

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - updated to use TempVars.Item("User") vs. cUser
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Validate the record and cancel updates if not valid
    If IsNull(Me.cmbPark_code) Then
        MsgBox "Please enter the park", vbOKOnly, "Validation error"
        Me.cmbPark_code.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtLocation_code) Then
        MsgBox "Please enter the location code", vbOKOnly, "Validation error"
        Me.txtLocation_code.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.cmbLocation_type) Then
        MsgBox "Please indicate the location type", vbOKOnly, "Validation error"
        Me.cmbLocation_type.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.cmbLocation_status) Then
        MsgBox "Please fill in the location status", vbOKOnly, "Validation error"
        Me.cmbLocation_status.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf Me.cmbLocation_type <> "Incidental" And IsNull(Me.cmbSite_ID) Then
        ' Site ID required for all except incidental/rare bird obs locations
        MsgBox "Please enter the site", vbOKOnly, "Validation error"
        Me.cmbSite_ID.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ' Check that the park matches the park in the sites table
    ElseIf Not IsNull(Me.cmbSite_ID) And Me.Site_park <> Me.cmbPark_code Then
        MsgBox "The park does not match the park in the site record", vbOKOnly, _
            "Validation error"
        Me.cmbPark_code.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ' Make sure that the record is not a duplicate prior to saving
    ' ... Site ID and location code are unique for all except incidental/rare bird obs locations
    ElseIf Me.cmbLocation_type <> "Incidental" And DCount("*", "tbl_Locations", _
        "[Site_ID]=""" & Me.cmbSite_ID & """ AND [Location_code]=""" & Me.txtLocation_code & _
        """ AND [Location_ID] <> """ & Me.Location_ID & """") > 0 Then
        MsgBox "A record with the same site and location code already exists.", _
            vbOKOnly, "Duplicate record found"
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf Me.txtLoc_discontinued < Me.Loc_established Then
        MsgBox "The discontinued date cannot be before the establisment date", _
            vbOKOnly, "Validation error"
        Me.txtLoc_discontinued.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtLoc_discontinued) = False Then
        If IsNull(Me.txtLoc_established) = False And _
            (Me.txtLoc_established > Me.txtLoc_discontinued) Then
            MsgBox "The discontinued date must be after the establishment date", _
                vbOKOnly, "Validation error"
            Me.txtLoc_discontinued.SetFocus
            DoCmd.CancelEvent
            GoTo Exit_Procedure
        End If
        If Me.cmbLocation_status = "Active" Then
            MsgBox "This location has a discontinued date. If the location" & vbCrLf & _
                "is discontinued, please change the status", vbOKOnly, "Validation error"
            Me.cmbLocation_status.SetFocus
            DoCmd.CancelEvent
            GoTo Exit_Procedure
        End If
    End If
    If IsNull(Me.txtLoc_established) = False And Me.cmbLocation_status = "Proposed" Then
        MsgBox "This location has an establishment date," & vbCrLf & _
            "but its status still indicates proposed", vbOKOnly, "Validation error"
        Me.cmbLocation_status.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

    ' Prior to saving, include a timestamp for edits
    If Me.NewRecord = False Then Me.txtLoc_updated = Now()
    ' Add the current user name to updated by
    If fxnSwitchboardIsOpen Then
        If IsNull(TempVars.item("User")) = False Then
            Me.txtLoc_updated_by = TempVars.item("User")
        Else
            Me.txtLoc_updated_by = Environ("Username")
        End If
    Else
        Me.txtLoc_updated_by = Environ("Username")
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub Form_AfterUpdate()
    On Error GoTo Err_Handler

    Me.subSite.Form!subLocations.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' The next set of procedures are for user form management and record navigation

' ---------------------------------
' SUB:     cmdFiltersOff_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdFiltersOff_Click()
    On Error GoTo Err_Handler

    ' Turn off the filters
    Me.cmdClose.SetFocus
    ' Undo the filter toggles
    Me.togFilterByPark = False
    Me.togFilterBySite = False
    Me.togFilterByLoc = False
    Me.togFilterByStatus = False
    Me.togFilterByType = False
    Me.togFilterByRegion = False
    Me.togFilterByStratum = False
    Me.togFilterByPanelType = False
    Me.togFilterByPanelName = False
    ' ... including those embedded in the record
    Me.togFilterByTrailOrRoad = False
    fxnFilterRecords
    Me.cmdFiltersOff.Enabled = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmdSave_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdSave_Click()
    On Error GoTo Err_Handler

    ' Save the record if appropriate
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    Me.cmdClose.SetFocus
    Me.cmdSave.Enabled = False
    Me.cmdDelete.Caption = "Delete"
    Me.subSite.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2001   ' Run time canceled event (validation error)
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmdNew_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdNew_Click()
    On Error GoTo Err_Handler

    ' When user wants a new record, move to a blank record and reset the form
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    If Me.AllowAdditions = False Then Me.AllowAdditions = True
    Me.DataEntry = True

    ' The filter is automatically turned off at this point, so undo the filter toggles
    Me.togFilterByPark = False
    Me.togFilterBySite = False
    Me.togFilterByLoc = False
    Me.togFilterByStatus = False
    Me.togFilterByType = False
    Me.togFilterByRegion = False
    Me.togFilterByStratum = False
    Me.togFilterByPanelType = False
    Me.togFilterByPanelName = False
    ' ... including those embedded in the record
    Me.togFilterByTrailOrRoad = False

    Me.cmbPark_code.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2001   ' Run time canceled event (validation error) - exit without
                    '   moving to a new record
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmdDelete_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdDelete_Click()
    On Error GoTo Err_Handler

    Dim blnHasChildren As Boolean   ' True if the location record has locked child records
    Dim strChildTables As String    ' Name(s) of table(s) containing child records
    Dim strTable As String          ' Current table name
    Dim intNumEvents As Integer     ' Number of related event records
    Dim strSQL As String

    ' Undo changes or delete the record according to whether it is new or being edited
    If Me.NewRecord Then
        Me.Undo
        Me.cmdClose.SetFocus
    ElseIf Me.Dirty Then
        Me.Undo
        Me.cmdDelete.Caption = "Delete"
        Me.cmdClose.SetFocus
    Else
        ' Bail out of delete if in data entry or or read only mode
        'If Forms!frm_Switchboard!cAppMode = "data entry" Or _
        '    Forms!frm_Switchboard!cAppMode = "read only" Then _
        If Me!fsub_DbAdmin.frm_Switchboard!cAppMode = "data entry" Or _
            Me!fsub_DbAdmin.frm_Switchboard!cAppMode = "read only" Then _
            GoTo Exit_Procedure

        ' Check to see if there are any records in related tables that would prevent
        '   successful deletion (i.e. because cascade deletes is off)
        blnHasChildren = False
        strChildTables = ""

        strTable = "tbl_Analysis_Notes"
        If DCount("*", strTable, "[Location_ID]=""" & Me.Location_ID & """") > 0 Then
            If blnHasChildren = True Then
                strChildTables = strChildTables & ", "
            Else: blnHasChildren = True
            End If
            strChildTables = strChildTables & strTable
        End If

        strTable = "tbl_Markers"
        If DCount("*", strTable, "[Location_ID]=""" & Me.Location_ID & """") > 0 Then
            If blnHasChildren = True Then
                strChildTables = strChildTables & ", "
            Else: blnHasChildren = True
            End If
            strChildTables = strChildTables & strTable
        End If

        strTable = "tbl_Target_Coords"
        If DCount("*", strTable, "[Location_ID]=""" & Me.Location_ID & """") > 0 Then
            If blnHasChildren = True Then
                strChildTables = strChildTables & ", "
            Else: blnHasChildren = True
            End If
            strChildTables = strChildTables & strTable
        End If

        strTable = "tbl_Task_List"
        If DCount("*", strTable, "[Location_ID]=""" & Me.Location_ID & """") > 0 Then
            If blnHasChildren = True Then
                strChildTables = strChildTables & ", "
            Else: blnHasChildren = True
            End If
            strChildTables = strChildTables & strTable
        End If

        If blnHasChildren Then
            ' Notify that deleting cannot be done by this form
            MsgBox "This record has related records in other tables that are" & vbCrLf & _
                "preventing it from being deleted in this form. These child" & vbCrLf & _
                "records must be deleted individually before the location" & vbCrLf & _
                "itself can be deleted." & vbCrLf & vbCrLf & strChildTables, _
                vbOKOnly + vbExclamation, "Unable to delete record"
            GoTo Exit_Procedure

        Else    ' If no child records remain ...
            ' Confirm with the user depending on how many related event records exist
            intNumEvents = DCount("*", "tbl_Events", "[Location_ID]=""" & Me.Location_ID & """")
            Select Case intNumEvents
              Case 0
                If MsgBox("Are you sure you want to delete this sampling location?", _
                    vbCritical + vbYesNo + vbDefaultButton2, _
                    "Confirm delete of sampling location") = vbYes Then
                    GoTo Delete_Record
                Else: GoTo Exit_Procedure
                End If
              Case 1
                If MsgBox("Are you certain you want to delete this location??" & _
                    vbCrLf & vbCrLf & _
                    "Park: " & Me.Park_code & "   Site: " & Me.Site_code & _
                    "   Loc code: " & Me.Location_code & vbCrLf & vbCrLf & _
                    "WARNING: this sample location has a related sampling event!", _
                    vbCritical + vbYesNo + vbDefaultButton2, _
                    "Confirm delete of location and sample event data ...") = vbYes Then
                    GoTo Delete_Record
                Else: GoTo Exit_Procedure
                End If
              Case Else
                If MsgBox("Are you certain you want to delete this location??" & _
                    vbCrLf & vbCrLf & _
                    "Park: " & Me.Park_code & "   Site: " & Me.Site_code & _
                    "   Loc code: " & Me.Location_code & vbCrLf & vbCrLf & _
                    "WARNING: this sample location has " & intNumEvents & _
                    " related sampling events!", _
                    vbCritical + vbYesNo + vbDefaultButton2, _
                    "Confirm delete of location and multiple sample events ...") = vbYes Then
                    GoTo Delete_Record
                Else: GoTo Exit_Procedure
                End If
            End Select

Delete_Record:
            ' Build the statement to delete the record
            strSQL = "DELETE * FROM tbl_Locations WHERE ((tbl_Locations.Location_ID) = """ _
                & Me.Location_ID & """)"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
            ' Go to a new record (record deleted)
            cmdNew_Click    ' Ask the user to log edits
            MsgBox "Please provide documentation of the deleted record ..."
            DoCmd.OpenForm "frm_Edit_Log", , , , , , "Delete tbl_Locations"
        End If
    End If

Exit_Procedure:
    On Error Resume Next
    DoCmd.SetWarnings True
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2501   ' Run time canceled event - exit without moving to a new record
        MsgBox "The record was not deleted"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmdClose_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

    ' Initiate save record to trigger validation
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2001   ' Run time canceled event (validation error) - exit without closing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' =================================
' The next set of procedures filters the recordset depending on user input

' ---------------------------------
' SUB:     cmbParkFilter_GotFocus
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbParkFilter_GotFocus()

    Me.ActiveControl.Requery

End Sub

' ---------------------------------
' SUB:     cmdParkFitler_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbParkFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByPark = Not IsNull(Me.cmbParkFilter)
    ' Turn off any filters that depend on park
    Me.cmbSiteFilter = Null
    Me.togFilterBySite = False
    Me.cmbLocFilter = Null
    Me.togFilterByLoc = False
    Me.cmbStratumFilter = Null
    Me.togFilterByStratum = False
    Me.cmbRegionFilter = Null
    Me.togFilterByRegion = False
    'Me.cmbWatershedFilter = Null
    'Me.togFilterByWatershed = False
    fxnFilterRecords
    Me.togFilterByPark.SetFocus
    Me.cmbSiteFilter.Requery
    Me.cmbLocFilter.Requery
    Me.cmbStratumFilter.Requery
    Me.cmbRegionFilter.Requery
    'Me.cmbWatershedFilter.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     togFilterByPark_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByPark_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbParkFilter) = False Then fxnFilterRecords _
        Else Me.togFilterByPark = False
    Me.cmbSiteFilter.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbSiteFilter_NotInList
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbSiteFilter_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    MsgBox "The site is not in the list. Make" & vbCrLf & _
        "sure the park filter is correct.", , "Not in list"
    Me.ActiveControl.Undo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbSiteFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbSiteFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterBySite = Not IsNull(Me.cmbSiteFilter)
    Me.cmbLocFilter = Null
    Me.togFilterByLoc = False
    fxnFilterRecords
    Me.cmbLocFilter.Requery
    Me.togFilterBySite.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     togFilterBySite_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterBySite_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbSiteFilter) = False Then fxnFilterRecords _
        Else Me.togFilterBySite = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbLocFilter_NotInList
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbLocFilter_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    MsgBox "The location code is not in the list. Make" & vbCrLf & _
        "sure the park/site filters are correct.", _
        , "Not in list"
    Me.ActiveControl.Undo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbLocFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbLocFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByLoc = Not IsNull(Me.cmbLocFilter)
    fxnFilterRecords
    Me.togFilterByLoc.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     togFilterByLoc_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByLoc_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbLocFilter) = False Then fxnFilterRecords _
        Else Me.togFilterByLoc = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbTypeFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbTypeFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByType = Not IsNull(Me.cmbTypeFilter)
    fxnFilterRecords
    Me.togFilterByType.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     togFilterByType_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByType_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbTypeFilter) = False Then fxnFilterRecords _
        Else Me.togFilterByType = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbStatusFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbStatusFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByStatus = Not IsNull(Me.cmbStatusFilter)
    fxnFilterRecords
    Me.togFilterByStatus.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     togFilterByStatus_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByStatus_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbStatusFilter) = False Then fxnFilterRecords _
        Else Me.togFilterByStatus = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbStratumFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbStratumFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByStratum = Not IsNull(Me.cmbStratumFilter)
    fxnFilterRecords
    Me.togFilterByStratum.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     togFilterByStratum_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByStratum_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbStratumFilter) = False Then fxnFilterRecords _
        Else Me.togFilterByStratum = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbRegionFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbRegionFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByRegion = Not IsNull(Me.cmbRegionFilter)
    fxnFilterRecords
    Me.togFilterByRegion.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     togFilterByRegion_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByRegion_AfterUpdate()
    On Error GoTo Err_Handler

    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbPanelTypeFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbPanelTypeFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByPanelType = Not IsNull(Me.cmbPanelTypeFilter)
    fxnFilterRecords
    Me.togFilterByPanelType.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     togFilterByPanelType_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByPanelType_AfterUpdate()
    On Error GoTo Err_Handler

    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbPanelNameFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbPanelNameFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByPanelName = Not IsNull(Me.cmbPanelNameFilter)
    fxnFilterRecords
    Me.togFilterByPanelName.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     togFilterByPanelName_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByPanelName_AfterUpdate()
    On Error GoTo Err_Handler

    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     togFilterByTrailOrRoad_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByTrailOrRoad_AfterUpdate()
    On Error GoTo Err_Handler

    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' The next set of procedures are for interacting with the user on edits to the current record

' ---------------------------------
' SUB:     cmbPark_code_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbPark_code_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    If Not IsNull(Me.cmbSite_ID) And Me.Site_park <> Me.cmbPark_code Then
        MsgBox "The park does not match the park in the site record", vbOKOnly, _
            "Validation error"
        DoCmd.CancelEvent
        Me.ActiveControl.Undo
        If Not Me.Dirty Then Me.cmdDelete.Caption = "Delete"
    ElseIf Not Me.NewRecord Then
        If MsgBox("Are you sure you want to change the park for this location?", _
            vbYesNo + vbDefaultButton2, "Confirm change to critical info") = vbNo Then
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
            If Not Me.Dirty Then Me.cmdDelete.Caption = "Delete"
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbSite_ID_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbSite_ID_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    If Not Me.NewRecord Then
        If MsgBox("Are you sure you want to change the site for this location?", _
            vbYesNo + vbDefaultButton2, "Confirm change to critical info") = vbNo Then
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
            If Not Me.Dirty Then Me.cmdDelete.Caption = "Delete"
        ElseIf Me.cmbSite_ID = "" Then
            MsgBox "Site ID cannot be set to null", vbOKOnly
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
            If Not Me.Dirty Then Me.cmdDelete.Caption = "Delete"
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     txtLocation_code_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub txtLocation_code_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    If Not Me.NewRecord Then
        If MsgBox("Are you sure you want to change the location code?", _
            vbYesNo + vbDefaultButton2, "Confirm change to critical info") = vbNo Then
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
            If Not Me.Dirty Then Me.cmdDelete.Caption = "Delete"
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbLocation_type_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbLocation_type_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    If Not Me.NewRecord Then
        If MsgBox("Are you sure you want to change the location type?", _
            vbYesNo + vbDefaultButton2, "Confirm change to critical info") = vbNo Then
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
            If Not Me.Dirty Then Me.cmdDelete.Caption = "Delete"
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbLocation_status_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbLocation_status_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    If Not Me.NewRecord Then
        If MsgBox("Are you sure you want to change the location status?", _
            vbYesNo + vbDefaultButton2, "Confirm change to critical info") = vbNo Then
            DoCmd.CancelEvent
            Me.ActiveControl.Undo
            If Not Me.Dirty Then Me.cmdDelete.Caption = "Delete"
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' PAGE NAME:    Site Information (pgSite)
' Description:  Information about the parent site record
' Unbound ctls: none
' Subforms:     subSite
' =================================

' =================================
' PAGE NAME:    Schedule (pgSchedule)
' Description:  Sampling schedule years for the current record
' Unbound ctls: none
' Subforms:     subSchedule
' =================================

' ---------------------------------
' SUB:     cmdScheduleForm_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdScheduleForm_Click()
    On Error GoTo Err_Handler

    DoCmd.OpenForm "frm_Schedule"

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' PAGE NAME:    Coordinates (pgCoordinates)
' Description:  Target, public, and event coordinates for the current record
' Unbound ctls: none
' Subforms:     subTarget_coords, subCoordinates
' =================================

' =================================
' PAGE NAME:    Tasks (pgTasks)
' Description:  Task items for the current record
' Unbound ctls: none
' Subforms:     subTasks
' =================================

' =================================
' PAGE NAME:    Events (pgEvents)
' Description:  Sampling event records for the current record
' Unbound ctls: none
' Subforms:     subEvents
' =================================

' =================================
' PAGE NAME:    Images (pgImages)
' Description:  Image records associated with the current location
' Unbound ctls: none
' Subforms:     subImages
' =================================

' =================================
' PAGE NAME:    Markers (pgMarkers)
' Description:  Marker records associated with the current location
' Unbound ctls: none
' Subforms:     subMarkers
' =================================

' =================================
' PAGE NAME:    Features (pgFeatures)
' Description:  Feature records associated with the current location
' Unbound ctls: none
' Subforms:     subFeatures
' =================================

' =================================
' PAGE NAME:    Analysis Info (pgAnalysis)
' Description:  Information about the selected location and site in support of data analysis
' Unbound ctls: none
' Subforms:     subAnalysisNotes, subVarianceGroups
' =================================

' ---------------------------------
' FUNCTION:     fxnFilterRecords
' Description:  Filter the records by the indicated field
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 5, 2006
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, May 2008 - made code more robust and error-proof
'               JRB, 11/7/2008 - updated to add site subform filter controls; added
'                   bOpenFilterOn to skip filter building and go directly to reformatting
' ---------------------------------
Private Function fxnFilterRecords(Optional ByVal bOpenFilterOn As Boolean)
    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim bFilterOn As Boolean

    bFilterOn = bOpenFilterOn   ' default is false
    If bOpenFilterOn Then GoTo Reformat_controls
    
    strFilter = ""

    ' Build the filter string depending on which fields are being filtered on
    If Me.togFilterByPark Then
        bFilterOn = True
        strFilter = "[Park_code] = """ & Me.cmbParkFilter & """"
    End If
    If Me.togFilterBySite Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        If cmbSiteFilter = "NA" Then
            strFilter = strFilter & "[Site_ID] Is Null"
        Else
            strFilter = strFilter & "[Site_ID] = """ & Me.cmbSiteFilter & """"
        End If
    End If
    If Me.togFilterByLoc Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Location_ID] = """ & Me.cmbLocFilter & """"
    End If
    If Me.togFilterByType Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Location_type] = """ & Me.cmbTypeFilter & """"
    End If
    If Me.togFilterByStatus Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Location_status] = """ & Me.cmbStatusFilter & """"
    End If
    If Me.togFilterByStratum Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Stratum_ID] = """ & Me.cmbStratumFilter & """"
    End If
    ' And for controls that allow null values ...
    If Me.togFilterByRegion Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        If IsNull(Me.cmbRegionFilter) Then
            strFilter = strFilter & "[Park_region] Is Null"
        Else
            strFilter = strFilter & "[Park_region] = """ & Me.cmbRegionFilter & """"
        End If
    End If
    If Me.togFilterByPanelType Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        If IsNull(Me.cmbPanelTypeFilter) Then
            strFilter = strFilter & "[Panel_type] Is Null"
        Else
            strFilter = strFilter & "[Panel_type] = """ & Me.cmbPanelTypeFilter & """"
        End If
    End If
    If Me.togFilterByPanelName Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        If IsNull(Me.cmbPanelNameFilter) Then
            strFilter = strFilter & "[Panel_name] Is Null"
        Else
            strFilter = strFilter & "[Panel_name] = """ & Me.cmbPanelNameFilter & """"
        End If
    End If

    ' And for controls that are embedded in the record ...
    If Me.togFilterByTrailOrRoad Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        If IsNull(Me.cmbTrail_or_road) Or Me.cmbTrail_or_road = "" Then
            strFilter = strFilter & "[Trail_or_road] Is Null"
        Else
            strFilter = strFilter & "[Trail_or_road] = """ & Me.cmbTrail_or_road & """"
        End If
    End If

    ' Set and activate the filter (or deactivate if none of the filter fields are on)
    Me.Filter = strFilter
    Me.FilterOn = bFilterOn

Reformat_controls:
    ' Enable/disable the command button accordingly
    Me.cmdFiltersOff.Enabled = bFilterOn

    ' Make the labels bold or not depending on filter settings
    ' For controls in the form header ...
    Me.labParkFilter.fontBold = Me.togFilterByPark
    Me.labSiteFilter.fontBold = Me.togFilterBySite
    Me.labLocFilter.fontBold = Me.togFilterByLoc
    Me.labTypeFilter.fontBold = Me.togFilterByType
    Me.labStatusFilter.fontBold = Me.togFilterByStatus
    Me.labStratumFilter.fontBold = Me.togFilterByStratum
    Me.labRegionFilter.fontBold = Me.togFilterByRegion
    Me.labPanelTypeFilter.fontBold = Me.togFilterByPanelType
    Me.labPanelNameFilter.fontBold = Me.togFilterByPanelName
    ' Update the font colors if filtering on that field
    If Me.togFilterByPark Then Me.cmbParkFilter.forecolor = 16711680 _
        Else Me.cmbParkFilter.forecolor = 0
    If Me.togFilterBySite Then Me.cmbSiteFilter.forecolor = 16711680 _
        Else Me.cmbSiteFilter.forecolor = 0
    If Me.togFilterByLoc Then Me.cmbLocFilter.forecolor = 16711680 _
        Else Me.cmbLocFilter.forecolor = 0
    If Me.togFilterByType Then Me.cmbTypeFilter.forecolor = 16711680 _
        Else Me.cmbTypeFilter.forecolor = 0
    If Me.togFilterByStatus Then Me.cmbStatusFilter.forecolor = 16711680 _
        Else Me.cmbStatusFilter.forecolor = 0
    If Me.togFilterByStratum Then Me.cmbStratumFilter.forecolor = 16711680 _
        Else Me.cmbStratumFilter.forecolor = 0
    If Me.togFilterByRegion Then Me.cmbRegionFilter.forecolor = 16711680 _
        Else Me.cmbRegionFilter.forecolor = 0
    If Me.togFilterByPanelType Then Me.cmbPanelTypeFilter.forecolor = 16711680 _
        Else Me.cmbPanelTypeFilter.forecolor = 0
    If Me.togFilterByPanelName Then Me.cmbPanelNameFilter.forecolor = 16711680 _
        Else Me.cmbPanelNameFilter.forecolor = 0
    ' Do the same for controls that are embedded in the record ...
    Me.labPark_code.FontItalic = Me.togFilterByPark
    If Me.togFilterByPark Then Me.labPark_code.forecolor = 16711680 _
        Else Me.labPark_code.forecolor = 0
    Me.labSite_ID.FontItalic = Me.togFilterBySite
    If Me.togFilterBySite Then Me.labSite_ID.forecolor = 16711680 _
        Else Me.labSite_ID.forecolor = 0
    Me.labLocation_code.FontItalic = Me.togFilterByLoc
    If Me.togFilterByLoc Then Me.labLocation_code.forecolor = 16711680 _
        Else Me.labLocation_code.forecolor = 0
    Me.labLocation_type.FontItalic = Me.togFilterByType
    If Me.togFilterByType Then Me.labLocation_type.forecolor = 16711680 _
        Else Me.labLocation_type.forecolor = 0
    Me.labLocation_status.FontItalic = Me.togFilterByStatus
    If Me.togFilterByStatus Then Me.labLocation_status.forecolor = 16711680 _
        Else Me.labLocation_status.forecolor = 0
    Me.labTrail_or_road.FontItalic = Me.togFilterByTrailOrRoad
    If Me.togFilterByTrailOrRoad Then Me.labTrail_or_road.forecolor = 16711680 _
        Else Me.labTrail_or_road.forecolor = 0
    ' ... And for controls that are embedded in the subform ...
    ' ... first make sure that the subform has a record
    If Me.Recordset.EOF = True Then GoTo Exit_Procedure
    If Me.subSite.Form.Recordset.EOF = True Then GoTo Exit_Procedure
    Me.subSite.Form!labStratum_ID.FontItalic = Me.togFilterByStratum
    Me.subSite.Form!labPark_region.FontItalic = Me.togFilterByRegion
    Me.subSite.Form!labPanel_type.FontItalic = Me.togFilterByPanelType
    Me.subSite.Form!labPanel_name.FontItalic = Me.togFilterByPanelName
    If Me.togFilterByStratum Then Me.subSite.Form!labStratum_ID.forecolor = 16711680 _
        Else Me.subSite.Form!labStratum_ID.forecolor = 0
    If Me.togFilterByRegion Then Me.subSite.Form!labPark_region.forecolor = 16711680 _
        Else: Me.subSite.Form!labPark_region.forecolor = 0
    If Me.togFilterByPanelType Then Me.subSite.Form!labPanel_type.forecolor = 16711680 _
        Else: Me.subSite.Form!labPanel_type.forecolor = 0
    If Me.togFilterByPanelName Then Me.subSite.Form!labPanel_name.forecolor = 16711680 _
        Else: Me.subSite.Form!labPanel_name.forecolor = 0

Exit_Procedure:
    Exit Function

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
        "Error encountered (#" & Err.Number & " - fxnFilterRecords)"
    Resume Exit_Procedure

End Function
