Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =48
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =15840
    DatasheetFontHeight =9
    ItemSuffix =30
    Left =1605
    Top =555
    Right =15885
    Bottom =9900
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2680758ff389e340
    End
    Caption =" Data Summary Tool"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =11520
            BackColor =15129564
            Name ="Detail"
            Begin
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =24
                    ListWidth =7200
                    Left =5520
                    Top =120
                    Width =8280
                    Height =300
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cmbQuery"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MSysObjects.Name, IIf(Mid([Name],7,1)='0','Quality check',IIf(Mid([Name],"
                        "7,1)='1','Annual report',IIf(Mid([Name],7,1)='2','Special analysis',IIf(Mid([Nam"
                        "e],7,1)='3','Action query',IIf(Mid([Name],7,1)='4','Export query',IIf(Left([Name"
                        "],5)='qsub_','Subquery','Other')))))) AS QType FROM MSysObjects WHERE (((MSysObj"
                        "ects.Name) Like \"qs_*\") AND ((MSysObjects.Type)=5) AND ((Mid([Name],7,1)) Like"
                        " [Forms]![frm_Summary_Tool]![cmbQTypeFilter])) OR (((MSysObjects.Name) Like \"qs"
                        "_*\") AND ((IIf(Mid([Name],7,1)='0','Quality check',IIf(Mid([Name],7,1)='1','Ann"
                        "ual report',IIf(Mid([Name],7,1)='2','Special analysis',IIf(Mid([Name],7,1)='3','"
                        "Action query',IIf(Mid([Name],7,1)='4','Export query',IIf(Left([Name],5)='qsub_',"
                        "'Subquery','Other')))))))=[Forms]![frm_Summary_Tool]![cmbQTypeFilter]) AND ((MSy"
                        "sObjects.Type)=5)) OR (((MSysObjects.Name) Like \"qsub*\") AND ((MSysObjects.Typ"
                        "e)=5) AND ((\"sub\")=[Forms]![frm_Summary_Tool]![cmbQTypeFilter])) ORDER BY MSys"
                        "Objects.Name; "
                    ColumnWidths ="5760;1440"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =120
                    LayoutCachedWidth =13800
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3900
                            Top =120
                            Width =1560
                            Height =240
                            FontSize =9
                            Name ="labQuery"
                            Caption ="Select the query:"
                            LayoutCachedLeft =3900
                            LayoutCachedTop =120
                            LayoutCachedWidth =5460
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin Subform
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =3480
                    Top =1260
                    Width =12360
                    Height =10260
                    TabIndex =8
                    Name ="subResults"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =1260
                    LayoutCachedWidth =15840
                    LayoutCachedHeight =11520
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15060
                    Top =660
                    Width =426
                    Height =426
                    TabIndex =7
                    Name ="cmdDesign"
                    Caption ="Design view"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000000d088888888888880a ,
                        0x080808080808080d000000000000000aa0eeeeeeee0dadadd0e0000ee0dadada ,
                        0xa0e0a0ee00adadadd0e00ee0d00adadaa0e0ee0da000adadd0eee0dad0b70ada ,
                        0xa0ee0dada0b80dadd0e0dadada0b70daa00dadadad0b00add0dadadadad0110a ,
                        0xadadadadada000ad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="View the selected query in design view"

                    LayoutCachedLeft =15060
                    LayoutCachedTop =660
                    LayoutCachedWidth =15486
                    LayoutCachedHeight =1086
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13980
                    Top =120
                    Width =426
                    Height =426
                    TabIndex =2
                    Name ="cmdChart"
                    Caption ="Chart view"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada00000000000000000ad0c010d0c010da ,
                        0x0da0c010a0c010ad0ad0c010d0c010da0da0c010a0c010ad0ad0c010d0c010da ,
                        0x0da0c000a0c010ad0ad0c0dad0c010da0da0c0ada00010ad0ad0c0dadad010da ,
                        0x0da000adada010ad0adadadadad010da0dadadadada010ad0adadadadad000da ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="View the selected query in chart view"

                    LayoutCachedLeft =13980
                    LayoutCachedTop =120
                    LayoutCachedWidth =14406
                    LayoutCachedHeight =546
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14520
                    Top =120
                    Width =426
                    Height =426
                    TabIndex =3
                    Name ="cmdPivotTable"
                    Caption ="Table view"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadadd00000000000000a ,
                        0xa0880fffffffff0dd0440f0f0f0f0f0aa0880fffffffff0dd0440f0f0f0f0f0a ,
                        0xa0880fffffffff0dd0440f0f0f0f0f0aa0880fffffffff0dd04400000000000a ,
                        0xa04448484848480dd04448484848480aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="View the selected query in pivot table view"

                    LayoutCachedLeft =14520
                    LayoutCachedTop =120
                    LayoutCachedWidth =14946
                    LayoutCachedHeight =546
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15060
                    Top =120
                    Width =426
                    Height =426
                    TabIndex =4
                    Name ="cmdCloseup"
                    Caption ="Zoom"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada00adadadadadadad000adadadadadada ,
                        0xa000adadadadadadda000a700007dadaada0000888800daddada07ee888870da ,
                        0xada708e88888807ddad08e888888880aada088888888880ddad088888888e80a ,
                        0xada088888888e80ddad70888888ee07aadad07888eee70addadad00888800ada ,
                        0xadadad700007adad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Open the selected query in a new window"

                    LayoutCachedLeft =15060
                    LayoutCachedTop =120
                    LayoutCachedWidth =15486
                    LayoutCachedHeight =546
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13980
                    Top =660
                    Width =426
                    Height =426
                    TabIndex =5
                    Name ="cmdExportExcel"
                    Caption ="Zoom"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada0000000dadadadadd00000dadadadada ,
                        0xad000dadadadadaddad0dadadadadadaadadadadad72727ddada2727272f272a ,
                        0xadad727272f272addada27272f2727daadada272f27272addadada2f2727dada ,
                        0xadada2f272727daddada2f27272727daadad72727d7272addada2727dad727da ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Export the selected query to Excel"

                    LayoutCachedLeft =13980
                    LayoutCachedTop =660
                    LayoutCachedWidth =14406
                    LayoutCachedHeight =1086
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14520
                    Top =660
                    Width =426
                    Height =426
                    TabIndex =6
                    Name ="cmdExportText"
                    Caption ="Zoom"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada0000000dadadadadd00000dadadadada ,
                        0xad000dad777777addad0dad00000077aadadad0ffffff07ddad000000888807a ,
                        0xad0e8e8e80fff07dda08e8e8e088807aad0e8e8e8e0ff07ddad0e0000808807a ,
                        0xada08e8e8e80f07ddada080000e0807aadad0e8e8e8e007ddadad0f0f0f000da ,
                        0xadadad0d0d0d0dad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Export the selected query to a text file"

                    LayoutCachedLeft =14520
                    LayoutCachedTop =660
                    LayoutCachedWidth =14946
                    LayoutCachedHeight =1086
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5520
                    Top =480
                    Width =8280
                    Height =660
                    TabIndex =1
                    Name ="txtDesc"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =480
                    LayoutCachedWidth =13800
                    LayoutCachedHeight =1140
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =300
                    Top =7800
                    Width =2880
                    Height =955
                    TabIndex =36
                    Name ="optgScope"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="Include data that hasn't yet passed the quality review?"

                    LayoutCachedLeft =300
                    LayoutCachedTop =7800
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =8755
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =360
                            Top =7860
                            Width =1890
                            Height =255
                            BackColor =13025979
                            Name ="labScope"
                            Caption ="Include uncertified data?"
                            FontName ="Arial"
                            LayoutCachedLeft =360
                            LayoutCachedTop =7860
                            LayoutCachedWidth =2250
                            LayoutCachedHeight =8115
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =480
                            Top =8190
                            OptionValue =0
                            Name ="optCertOnly"

                            LayoutCachedLeft =480
                            LayoutCachedTop =8190
                            LayoutCachedWidth =740
                            LayoutCachedHeight =8430
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =720
                                    Top =8160
                                    Width =2280
                                    Height =270
                                    FontWeight =700
                                    Name ="labCertOnly"
                                    Caption ="No (use only certified data)"
                                    FontName ="Arial"
                                    ControlTipText ="Run queries only on certified event data"
                                    LayoutCachedLeft =720
                                    LayoutCachedTop =8160
                                    LayoutCachedWidth =3000
                                    LayoutCachedHeight =8430
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =480
                            Top =8490
                            OptionValue =1
                            Name ="optBoth"

                            LayoutCachedLeft =480
                            LayoutCachedTop =8490
                            LayoutCachedWidth =740
                            LayoutCachedHeight =8730
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =720
                                    Top =8460
                                    Width =2400
                                    Height =270
                                    Name ="labBoth"
                                    Caption ="Yes (results are provisional)"
                                    FontName ="Arial"
                                    LayoutCachedLeft =720
                                    LayoutCachedTop =8460
                                    LayoutCachedWidth =3120
                                    LayoutCachedHeight =8730
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =840
                    Width =1980
                    Height =317
                    FontWeight =700
                    TabIndex =11
                    ForeColor =0
                    Name ="cmdOpenBrowser"
                    Caption ="Open data browser"
                    StatusBarText ="Open the project data browser"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Open the project data browser"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3540
                    Top =840
                    Width =1020
                    Height =317
                    FontWeight =700
                    TabIndex =12
                    ForeColor =0
                    Name ="cmdRequery"
                    Caption ="Requery"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Requery the results set for the selected query"

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
                    Left =1140
                    Top =1800
                    Width =1620
                    Height =270
                    TabIndex =13
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cmbParkFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.Park_code FROM tlu_Parks; "
                    StatusBarText ="Filter by park"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by park"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =1800
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =2070
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =480
                            Top =1800
                            Width =540
                            Height =228
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labParkFilter"
                            Caption ="Park:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =2880
                    Top =1800
                    Width =480
                    Height =300
                    TabIndex =14
                    ForeColor =0
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

                    LayoutCachedLeft =2880
                    LayoutCachedTop =1800
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =2100
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
                    ListWidth =6480
                    Left =1140
                    Top =2220
                    Width =1620
                    Height =270
                    TabIndex =15
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cmbTypeFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Location_Type.Location_type, tlu_Location_Type.Loc_type_desc FROM tlu"
                        "_Location_Type ORDER BY tlu_Location_Type.Sort_order; "
                    ColumnWidths ="1008;5472"
                    StatusBarText ="Filter by location type"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by location type"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =2220
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =2490
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =2220
                            Width =840
                            Height =228
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labTypeFilter"
                            Caption ="Loc type:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =2880
                    Top =2220
                    Width =480
                    Height =300
                    TabIndex =16
                    ForeColor =0
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

                    LayoutCachedLeft =2880
                    LayoutCachedTop =2220
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =2520
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =20
                    Left =1200
                    Top =6300
                    Width =1260
                    Height =270
                    TabIndex =31
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cmbYearFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qfrm_Data_Gateway.Sample_year FROM qfrm_Data_Gateway WHERE (((qfrm_Data_G"
                        "ateway.Sample_year) Is Not Null)) GROUP BY qfrm_Data_Gateway.Sample_year ORDER B"
                        "Y qfrm_Data_Gateway.Sample_year DESC; "
                    StatusBarText ="Filter by event year"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by event year"

                    LayoutCachedLeft =1200
                    LayoutCachedTop =6300
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =6570
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =540
                            Top =6300
                            Width =540
                            Height =255
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labYearFilter"
                            Caption ="Year:"
                            FontName ="Arial"
                            LayoutCachedLeft =540
                            LayoutCachedTop =6300
                            LayoutCachedWidth =1080
                            LayoutCachedHeight =6555
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =2580
                    Top =6300
                    Width =480
                    Height =300
                    TabIndex =32
                    ForeColor =0
                    Name ="togFilterByYear"
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
                    ControlTipText ="Turn the year filter on or off"

                    LayoutCachedLeft =2580
                    LayoutCachedTop =6300
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =6600
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =840
                    Width =1201
                    Height =294
                    FontWeight =700
                    TabIndex =10
                    ForeColor =0
                    Name ="cmdFiltersOff"
                    Caption ="Filters off"
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
                    ColumnCount =2
                    ListWidth =7200
                    Left =1140
                    Top =2640
                    Width =1620
                    Height =270
                    TabIndex =17
                    BackColor =-2147483643
                    Name ="cmbStatusFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Site_Status.Site_status, tlu_Site_Status.Site_status_desc FROM tlu_Si"
                        "te_Status ORDER BY tlu_Site_Status.Sort_order; "
                    ColumnWidths ="1008;6192"
                    StatusBarText ="Filter by location status"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by location status"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =2640
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =2910
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =300
                            Top =2640
                            Width =720
                            Height =225
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labStatusFilter"
                            Caption ="Status:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =2880
                    Top =2640
                    Width =480
                    Height =300
                    TabIndex =18
                    ForeColor =0
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

                    LayoutCachedLeft =2880
                    LayoutCachedTop =2640
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =2940
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
                    Left =1140
                    Top =4763
                    Width =1620
                    Height =270
                    TabIndex =27
                    BackColor =-2147483643
                    Name ="cmbPanelTypeFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Panel_Type.Panel_type, tlu_Panel_Type.Panel_type_desc FROM tlu_Panel_"
                        "Type ORDER BY tlu_Panel_Type.Sort_order; "
                    ColumnWidths ="1152;8208"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by sampling panel type"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =4763
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =5033
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =4763
                            Width =960
                            Height =255
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labPanelTypeFilter"
                            Caption ="Panel type:"
                            FontName ="Arial"
                            LayoutCachedLeft =120
                            LayoutCachedTop =4763
                            LayoutCachedWidth =1080
                            LayoutCachedHeight =5018
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =2880
                    Top =4740
                    Width =480
                    Height =300
                    TabIndex =28
                    ForeColor =0
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

                    LayoutCachedLeft =2880
                    LayoutCachedTop =4740
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =5040
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =2880
                    Top =5160
                    Width =480
                    Height =300
                    TabIndex =30
                    ForeColor =0
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

                    LayoutCachedLeft =2880
                    LayoutCachedTop =5160
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =5460
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =5183
                    Width =1620
                    Height =270
                    TabIndex =29
                    BackColor =-2147483643
                    Name ="cmbPanelNameFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Sites.Panel_name FROM tbl_Sites GROUP BY tbl_Sites.Panel_name ORDER B"
                        "Y tbl_Sites.Panel_name; "
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by panel name"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =5183
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =5453
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =5183
                            Width =1080
                            Height =255
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labPanelNameFilter"
                            Caption ="Panel name:"
                            FontName ="Arial"
                            LayoutCachedTop =5183
                            LayoutCachedWidth =1080
                            LayoutCachedHeight =5438
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1200
                    Top =6960
                    Width =1224
                    Height =270
                    TabIndex =33
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtStartDateFilter"
                    Format ="yyyy mmm dd"
                    StatusBarText ="Start date for filters"
                    FontName ="Arial"

                    LayoutCachedLeft =1200
                    LayoutCachedTop =6960
                    LayoutCachedWidth =2424
                    LayoutCachedHeight =7230
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =120
                            Top =6960
                            Width =966
                            Height =252
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labStartDateFilter"
                            Caption ="From date:"
                            FontName ="Arial"
                            LayoutCachedLeft =120
                            LayoutCachedTop =6960
                            LayoutCachedWidth =1086
                            LayoutCachedHeight =7212
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1200
                    Top =7320
                    Width =1224
                    Height =270
                    TabIndex =34
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtEndDateFilter"
                    Format ="yyyy mmm dd"
                    StatusBarText ="End date for filters"
                    FontName ="Arial"

                    LayoutCachedLeft =1200
                    LayoutCachedTop =7320
                    LayoutCachedWidth =2424
                    LayoutCachedHeight =7590
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =180
                            Top =7320
                            Width =903
                            Height =252
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labEndDateFilter"
                            Caption ="To date:"
                            FontName ="Arial"
                            LayoutCachedLeft =180
                            LayoutCachedTop =7320
                            LayoutCachedWidth =1083
                            LayoutCachedHeight =7572
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
                    ListWidth =4032
                    Left =1140
                    Top =3480
                    Width =1620
                    Height =270
                    TabIndex =21
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cmbLocFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Location_ID, tbl_Locations.Location_code, [tbl_Locations].["
                        "Park_code] & '.' & [Site_code] AS Site, tbl_Locations.Location_type, tbl_Locatio"
                        "ns.Location_status FROM tbl_Sites RIGHT JOIN tbl_Locations ON tbl_Sites.Site_ID "
                        "= tbl_Locations.Site_ID WHERE (((tbl_Locations.Park_code) Like Nz([Forms]![frm_S"
                        "ummary_Tool]![cmbParkFilter],\"*\")) AND ((Nz([tbl_Locations].[Site_ID],\"*\")) "
                        "Like Nz([Forms]![frm_Summary_Tool]![cmbSiteFilter],\"*\"))) ORDER BY tbl_Locatio"
                        "ns.Location_status, tbl_Sites.Site_code, tbl_Locations.Location_code; "
                    ColumnWidths ="0;720;1152;1008;1152"
                    StatusBarText ="Filter by sample location"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by sample location"
                    LayoutCachedLeft =1140
                    LayoutCachedTop =3480
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =3750
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =3480
                            Width =840
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labLocFilter"
                            Caption ="Location:"
                            FontName ="Arial"
                            LayoutCachedLeft =180
                            LayoutCachedTop =3480
                            LayoutCachedWidth =1020
                            LayoutCachedHeight =3720
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =2880
                    Top =3480
                    Width =480
                    Height =300
                    TabIndex =22
                    ForeColor =0
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

                    LayoutCachedLeft =2880
                    LayoutCachedTop =3480
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =3780
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =840
                    Top =1380
                    Width =1560
                    Height =255
                    Name ="labLocFilters"
                    Caption ="Location filters:"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =900
                    Top =5940
                    Width =1560
                    Height =255
                    Name ="labEventFilters"
                    Caption ="Event filters:"
                    LayoutCachedLeft =900
                    LayoutCachedTop =5940
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =6195
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =2580
                    Top =7320
                    Width =480
                    Height =300
                    TabIndex =35
                    ForeColor =0
                    Name ="togFilterByRange"
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
                    ControlTipText ="Turn the date range filter on or off"

                    LayoutCachedLeft =2580
                    LayoutCachedTop =7320
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =7620
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =720
                    Top =6600
                    Width =1575
                    Height =255
                    Name ="labOr"
                    Caption ="Or"
                    LayoutCachedLeft =720
                    LayoutCachedTop =6600
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =6855
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =20
                    ListWidth =2880
                    Left =1140
                    Top =3924
                    Width =1620
                    Height =270
                    TabIndex =23
                    BackColor =-2147483643
                    Name ="cmbStratumFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Strata.Stratum_ID, [Park_code] & \" - \" & [Stratum_name] AS Stratum,"
                        " tbl_Strata.Stratification_date FROM tbl_Strata WHERE (((tbl_Strata.Park_code) L"
                        "ike Nz([Forms]![frm_Summary_Tool]![cmbParkFilter],\"*\"))) ORDER BY tbl_Strata.P"
                        "ark_code, tbl_Strata.Stratum_name; "
                    ColumnWidths ="0;1728;1152"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by stratum"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =3924
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =4194
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =300
                            Top =3924
                            Width =780
                            Height =255
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labStratumFilter"
                            Caption ="Stratum:"
                            FontName ="Arial"
                            LayoutCachedLeft =300
                            LayoutCachedTop =3924
                            LayoutCachedWidth =1080
                            LayoutCachedHeight =4179
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =2880
                    Top =3900
                    Width =480
                    Height =300
                    TabIndex =24
                    ForeColor =0
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

                    LayoutCachedLeft =2880
                    LayoutCachedTop =3900
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =4200
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
                    Left =1140
                    Top =4343
                    Width =1620
                    Height =255
                    TabIndex =25
                    BackColor =-2147483643
                    Name ="cmbRegionFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Sites.Park_region FROM tbl_Sites WHERE (((tbl_Sites.Park_code) Like N"
                        "z([Forms]![frm_Summary_Tool]![cmbParkFilter],\"*\"))) GROUP BY tbl_Sites.Park_re"
                        "gion ORDER BY tbl_Sites.Park_region; "
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by park region"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =4343
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =4598
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =360
                            Top =4343
                            Width =720
                            Height =255
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labRegionFilter"
                            Caption ="Region:"
                            FontName ="Arial"
                            LayoutCachedLeft =360
                            LayoutCachedTop =4343
                            LayoutCachedWidth =1080
                            LayoutCachedHeight =4598
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =2880
                    Top =4320
                    Width =480
                    Height =300
                    TabIndex =26
                    ForeColor =0
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

                    LayoutCachedLeft =2880
                    LayoutCachedTop =4320
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =4620
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =5880
                    Width =3300
                    Height =3960
                    Name ="Box22"
                    LayoutCachedLeft =120
                    LayoutCachedTop =5880
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =9840
                End
                Begin OptionGroup
                    OverlapFlags =255
                    Left =300
                    Top =8820
                    Width =2880
                    Height =955
                    TabIndex =37
                    Name ="optgExcluded"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="Include events flagged to be excluded from summary data output?"

                    LayoutCachedLeft =300
                    LayoutCachedTop =8820
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =9775
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =360
                            Top =8880
                            Width =2490
                            Height =255
                            BackColor =13025979
                            Name ="labExcluded"
                            Caption ="Include excluded event records?"
                            FontName ="Arial"
                            LayoutCachedLeft =360
                            LayoutCachedTop =8880
                            LayoutCachedWidth =2850
                            LayoutCachedHeight =9135
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =480
                            Top =9210
                            OptionValue =0
                            Name ="optExclude"

                            LayoutCachedLeft =480
                            LayoutCachedTop =9210
                            LayoutCachedWidth =740
                            LayoutCachedHeight =9450
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =720
                                    Top =9180
                                    Width =2280
                                    Height =270
                                    FontWeight =700
                                    Name ="labExclude"
                                    Caption ="No (recommended)"
                                    FontName ="Arial"
                                    ControlTipText ="Exclude flagged events from query results"
                                    LayoutCachedLeft =720
                                    LayoutCachedTop =9180
                                    LayoutCachedWidth =3000
                                    LayoutCachedHeight =9450
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =480
                            Top =9510
                            OptionValue =1
                            Name ="optInclude"

                            LayoutCachedLeft =480
                            LayoutCachedTop =9510
                            LayoutCachedWidth =740
                            LayoutCachedHeight =9750
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =720
                                    Top =9480
                                    Width =540
                                    Height =270
                                    Name ="labInclude"
                                    Caption ="Yes"
                                    FontName ="Arial"
                                    ControlTipText ="Include all events in query results, even those flagged for exclusion"
                                    LayoutCachedLeft =720
                                    LayoutCachedTop =9480
                                    LayoutCachedWidth =1260
                                    LayoutCachedHeight =9750
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =1560
                    Top =9480
                    Width =1501
                    Height =294
                    FontWeight =700
                    TabIndex =38
                    ForeColor =0
                    Name ="cmdViewExcluded"
                    Caption ="View flagged"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="View flagged event records in a query"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =9480
                    LayoutCachedWidth =3061
                    LayoutCachedHeight =9774
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2820
                    Top =480
                    Width =606
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =9
                    BackColor =8454143
                    Name ="txtUnfilteredFlag"
                    ControlTipText ="Indicates whether results for the selected query can be filtered"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =480
                    LayoutCachedWidth =3426
                    LayoutCachedHeight =735
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =240
                            Top =480
                            Width =2520
                            Height =255
                            Name ="labUnfilteredFlag"
                            Caption ="Query returns filtered results?"
                            FontName ="Arial"
                            ControlTipText ="Indicates whether results for the selected query can be filtered"
                            LayoutCachedLeft =240
                            LayoutCachedTop =480
                            LayoutCachedWidth =2760
                            LayoutCachedHeight =735
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
                    ListWidth =5040
                    Left =1140
                    Top =3060
                    Width =1620
                    Height =270
                    TabIndex =19
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cmbSiteFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Sites.Site_ID, tbl_Sites.Site_code, tbl_Sites.Park_code, tbl_Sites.Si"
                        "te_status, tbl_Sites.Site_name FROM tbl_Sites WHERE (((tbl_Sites.Park_code) Like"
                        " Nz([Forms]![frm_Summary_Tool]![cmbParkFilter],\"*\"))) ORDER BY tbl_Sites.Site_"
                        "status, tbl_Sites.Site_code; "
                    ColumnWidths ="0;720;1152;1152;1728"
                    StatusBarText ="Filter by site"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by site"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =3060
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =3330
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =300
                            Top =3060
                            Width =720
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labSiteFilter"
                            Caption ="Site:"
                            FontName ="Arial"
                            LayoutCachedLeft =300
                            LayoutCachedTop =3060
                            LayoutCachedWidth =1020
                            LayoutCachedHeight =3300
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =2880
                    Top =3060
                    Width =480
                    Height =300
                    TabIndex =20
                    ForeColor =0
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

                    LayoutCachedLeft =2880
                    LayoutCachedTop =3060
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =3360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1650
                    Top =127
                    Width =1530
                    TabIndex =39
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cmbQTypeFilter"
                    RowSourceType ="Value List"
                    RowSource ="'*';'Show all queries';'1';'Annual report';'2';'Special analysis';'3';'Action qu"
                        "eries';'4';'Export queries';'0';'Quality checks';'sub';'Subqueries';'Other';'Oth"
                        "er'"
                    ColumnWidths ="0;2160"
                    StatusBarText ="Filter by query type"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"*\""
                    FontName ="Arial"
                    ControlTipText ="Filter by query type"

                    LayoutCachedLeft =1650
                    LayoutCachedTop =127
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =367
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =299
                            Top =120
                            Width =1260
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="labQTypeFilter"
                            Caption ="Query type:"
                            FontName ="Arial"
                            LayoutCachedLeft =299
                            LayoutCachedTop =120
                            LayoutCachedWidth =1559
                            LayoutCachedHeight =360
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
' FORM NAME:    frm_Summary_Tool
' Description:  Standard form for summarizing and exploring project data
' Data source:  unbound
' Data access:  edit only, no additions or deletions
' Pages:        none
' Functions:    fxnFilterRecords
' References:   fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, Jan 2010
' Revisions:    JRB, 7/10/2013 - updated by adding cmbSiteFilter and togFilterBySite, and
'                   by resizing the form
' =================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Close the form if the switchboard is not open
    If fxnSwitchboardIsOpen = False Then
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this form to function properly.", , "Cannot open the form ..."
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbQuery_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    Me.ActiveControl.Undo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbQuery_AfterUpdate()
    On Error GoTo Err_Handler

    ' Exit if no query selected
    If IsNull(Me.cmbQuery) Or Me.cmbQuery Like "qsub*" Then
        Me.txtUnfilteredFlag = "?"
        Me.txtUnfilteredFlag.forecolor = 0          'black
        Me.txtUnfilteredFlag.backcolor = 8454143    'yellow
        Me.subResults.SourceObject = ""
        Me.txtDesc = ""
        If IsNull(Me.cmbQuery) Then GoTo Exit_Procedure
    End If

    ' Update the description
    Me.txtDesc = ""

    Dim qdf As DAO.QueryDef
    Dim qdfs As DAO.QueryDefs
    Set qdfs = DBEngine(0)(0).QueryDefs

    On Error Resume Next
    For Each qdf In qdfs
        If qdf.Name = Me.cmbQuery.Value Then
            Me.txtDesc = qdf.Properties("Description")
        End If
    Next qdf
    Me.Repaint

    On Error GoTo Err_Handler
    ' Bind the subform to the newly-selected object
    Me.subResults.SourceObject = "Query." & Me.cmbQuery.Value

    ' Update the visual flag to indicate whether or not the query returns filtered results
    '   Note: suffix of "_X" means that the query cannot accept parameters (e.g., crosstab)
    If Right(Me.cmbQuery.Value, 2) = "_X" Then
        Me.txtUnfilteredFlag = "No"
        Me.txtUnfilteredFlag.forecolor = 16777215   'white
        Me.txtUnfilteredFlag.backcolor = 255        'red
    ElseIf Left(Me.cmbQuery.Value, 4) <> "qsub" Then
        Me.txtUnfilteredFlag = "Yes"
        Me.txtUnfilteredFlag.forecolor = 16777215   'white
        Me.txtUnfilteredFlag.backcolor = 4227072    'green
    End If

    ' Set focus to the subform to allow scrolling, etc.
    Me.subResults.SetFocus

Exit_Procedure:
    On Error Resume Next
    Set qdfs = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cmbQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmbQTypeFilter_AfterUpdate()
    On Error GoTo Err_Handler

    ' Update the query set
    If IsNull(Me.cmbQTypeFilter) Then Me.cmbQTypeFilter = "*"
    Me.cmbQuery.Requery
    Me.cmbQuery.SetFocus
    Me.cmbQuery.Dropdown

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdOpenBrowser_Click()
    On Error GoTo Err_Handler

    Set gvarRefForm = Me.Form
    Set gvarRefCtl = Me.subResults
    ' Open to a blank record - to distinguish from opening to the selected record in the subform
    DoCmd.OpenForm "frm_Data_Browser", , , , acFormAdd, , "off"

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox Err.Number & ": " & Err.Description
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdRequery_Click()
    On Error GoTo Err_Handler

    ' Bail out if no query is currently selected
    If IsNull(Me.cmbQuery) Then GoTo Exit_Procedure

    ' Requery the selected record in the recordset, and update the subform
    Me.subResults.Requery
    Me.subResults.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' The next set of procedures filters the recordset depending on user input

Private Sub cmdFiltersOff_Click()
    On Error GoTo Err_Handler

    ' Turn off the filters
    Me.cmdRequery.SetFocus
    ' Undo the filter toggles
    Me.togFilterByPark = False
    Me.togFilterByStatus = False
    Me.togFilterByType = False
    Me.togFilterBySite = False
    Me.togFilterByLoc = False
    Me.togFilterByRegion = False
    Me.togFilterByStratum = False
    Me.togFilterByPanelType = False
    Me.togFilterByPanelName = False
    Me.togFilterByYear = False
    Me.togFilterByRange = False
    ' Non-standard fields
    'Me.togFilterByWatershed = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' Location filter controls

Private Sub cmbParkFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByPark = Not IsNull(Me.cmbParkFilter)
    ' Turn off any filters that depend on park
    Me.cmbSiteFilter = Null
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
    'Me.cmbWatershedFilter.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub togFilterByPark_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbParkFilter) = True Then Me.togFilterByPark = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

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

Private Sub togFilterByType_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbTypeFilter) = True Then Me.togFilterByType = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

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

Private Sub togFilterByStatus_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbStatusFilter) = True Then Me.togFilterByStatus = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbSiteFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterBySite = Not IsNull(Me.cmbSiteFilter)
    ' Turn off any filters that depend on site
    Me.cmbLocFilter = Null
    Me.togFilterByLoc = False
    fxnFilterRecords
    Me.togFilterBySite.SetFocus
    Me.cmbLocFilter.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub togFilterBySite_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbSiteFilter) = True Then Me.togFilterBySite = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

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

Private Sub togFilterByLoc_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbLocFilter) = True Then Me.togFilterByLoc = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

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

Private Sub togFilterByStratum_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbStratumFilter) = True Then Me.togFilterByStratum = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

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

Private Sub togFilterByRegion_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbRegionFilter) = True Then Me.togFilterByRegion = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

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

Private Sub togFilterByPanelType_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbPanelTypeFilter) = True Then Me.togFilterByPanelType = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

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

Private Sub togFilterByPanelName_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbPanelNameFilter) = True Then Me.togFilterByPanelName = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' Event filter controls

Private Sub cmbYearFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByYear = Not IsNull(Me.cmbYearFilter)
    If Me.togFilterByYear = True Then Me.togFilterByRange = False
    fxnFilterRecords
    Me.togFilterByYear.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub togFilterByYear_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbYearFilter) Then Me.togFilterByYear = False
    If Me.togFilterByYear = True Then Me.togFilterByRange = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub togFilterByRange_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.txtStartDateFilter) And IsNull(Me.txtEndDateFilter) _
        Then Me.togFilterByRange = False
    If Me.togFilterByRange = True Then Me.togFilterByYear = False
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub optgScope_AfterUpdate()
    On Error GoTo Err_Handler

    If Me.optgScope = 1 Then
        If MsgBox("Warning: The summary results may be based on data" & vbCrLf & _
            "that have not yet passed the quality review." & vbCrLf & vbCrLf & _
            "As such the results should be considered provisional" & vbCrLf & _
            "and should only be shared or reported on in a way" & vbCrLf & _
            "that clearly indicates this.", vbExclamation + vbOKCancel + vbDefaultButton2, _
            "Include uncertified data?") = vbCancel Then
            Me.optgScope = 0
            Me.labCertOnly.fontBold = True
            Me.labBoth.fontBold = False
            Me.labBoth.forecolor = 0
        Else
            Me.labCertOnly.fontBold = False
            Me.labBoth.fontBold = True
            Me.labBoth.forecolor = 255
        End If
    Else
        Me.labCertOnly.fontBold = True
        Me.labBoth.fontBold = False
        Me.labBoth.forecolor = 0
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub optgExcluded_AfterUpdate()
    On Error GoTo Err_Handler

    If Me.optgExcluded = 1 Then
        If MsgBox("Include all sampling events in query results, even" & vbCrLf & _
            "those that have been flagged for exclusion from" & vbCrLf & _
            "summary output?" & vbCrLf & vbCrLf & _
            "Note that this may change summary statistics already" & vbCrLf & _
            "reported on for prior years.", _
            vbExclamation + vbOKCancel + vbDefaultButton2, _
            "Override sampling event exclusion flags?") = vbCancel Then
            Me.optgExcluded = 0
            Me.labExclude.fontBold = True
            Me.labInclude.fontBold = False
            Me.labInclude.forecolor = 0
        Else
            Me.labExclude.fontBold = False
            Me.labInclude.fontBold = True
            Me.labInclude.forecolor = 255
        End If
    Else
        Me.labExclude.fontBold = True
        Me.labInclude.fontBold = False
        Me.labInclude.forecolor = 0
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdViewExcluded_Click()
    On Error GoTo Err_Handler

    ' Open the query to view event records flagged for exclusion from summaries
    DoCmd.OpenQuery "qsub_Excluded_events", acViewNormal, acReadOnly

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cmbQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' =================================
' The next set of procedures relate to manipulating the selected query/results

Private Sub cmdChart_Click()
    On Error GoTo Err_Handler

    ' Open the selected query as a pivot chart after checking that a query is selected
    If IsNull(Me.cmbQuery) = False Then
        DoCmd.OpenQuery Me.cmbQuery.Value, acViewPivotChart, acReadOnly
        DoCmd.Maximize
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cmbQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdPivotTable_Click()
    On Error GoTo Err_Handler

    ' Open the selected query as a pivot table after checking that a query is selected
    If IsNull(Me.cmbQuery) = False Then
        DoCmd.OpenQuery Me.cmbQuery.Value, acViewPivotTable, acReadOnly
        DoCmd.Maximize
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cmbQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdCloseup_Click()
    On Error GoTo Err_Handler

    ' Open the selected query in a new window after checking that a query is selected
    If IsNull(Me.cmbQuery) = False Then
        DoCmd.OpenQuery Me.cmbQuery.Value, acViewNormal, acReadOnly
        DoCmd.Maximize
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cmbQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdExportExcel_Click()
    On Error GoTo Err_Handler

    Dim strQryName As String
    Dim strInitFile As String
    Dim strSaveFile As String

    ' Bail out if no query is currently selected
    If IsNull(Me.cmbQuery) Then GoTo Exit_Procedure

    strQryName = Me.cmbQuery

    strInitFile = Application.CurrentProject.Path & "\" & _
        strQryName & "_" & CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".xls"
    ' Open the save file dialog and update to the actual name given by the user
    strSaveFile = fxnSaveFile(strInitFile, "Microsoft Excel (*.xls)", "*.xls")
    DoCmd.OutputTo acOutputQuery, strQryName, acFormatXLS, strSaveFile, True
    'MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdExportText_Click()
    On Error GoTo Err_Handler

    Dim strQryName As String
    Dim strInitFile As String
    Dim strSaveFile As String

    ' Bail out if no query is currently selected
    If IsNull(Me.cmbQuery) Then GoTo Exit_Procedure

    strQryName = Me.cmbQuery

    strInitFile = Application.CurrentProject.Path & "\" & _
        strQryName & "_" & CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".txt"
    ' Open the save file dialog and update to the actual name given by the user
    strSaveFile = fxnSaveFile(strInitFile, "Microsoft Excel (*.txt)", "*.txt")
    DoCmd.OutputTo acOutputQuery, strQryName, acFormatTXT, strSaveFile, True
    'MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdDesign_Click()
    On Error GoTo Err_Handler

    ' Open the selected query in design view after checking that a query is selected
    If IsNull(Me.cmbQuery) = False Then _
        DoCmd.OpenQuery Me.cmbQuery.Value, acViewDesign, acReadOnly

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cmbQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' =================================
' FUNCTION:     fxnFilterRecords
' Description:  Filter the records by the indicated field
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 5, 2006
' Revisions:    JRB, 1/5/2010 - adapted to summarization tool, mainly for formatting filters
'               JRB, 7/10/2013 - updated by adding site filters
' =================================

Private Function fxnFilterRecords()
    On Error GoTo Err_Handler

    Dim bFilterOn As Boolean

    bFilterOn = False

    ' If any toggles are on, the filter is on
    If Me.togFilterByPark Or Me.togFilterByType Or Me.togFilterByStatus Or _
        Me.togFilterBySite Or Me.togFilterByLoc Or Me.togFilterByStratum Then bFilterOn = True
    ' And for loc filters that allow null values ...
    If Me.togFilterByRegion Or Me.togFilterByPanelType Or _
        Me.togFilterByPanelName Then bFilterOn = True
    ' And for event filters
    If Me.togFilterByYear Or Me.togFilterByRange Then bFilterOn = True
    ' Non-standard fields
    'If Me.togFilterByWatershed Then bFilterOn = True

Reformat_controls:
    ' Enable/disable the command button accordingly
    Me.cmdFiltersOff.Enabled = bFilterOn

    ' Make the labels bold or not depending on filter settings
    Me.labParkFilter.fontBold = Me.togFilterByPark
    Me.labTypeFilter.fontBold = Me.togFilterByType
    Me.labStatusFilter.fontBold = Me.togFilterByStatus
    Me.labSiteFilter.fontBold = Me.togFilterBySite
    Me.labLocFilter.fontBold = Me.togFilterByLoc
    Me.labStratumFilter.fontBold = Me.togFilterByStratum
    Me.labRegionFilter.fontBold = Me.togFilterByRegion
    Me.labPanelTypeFilter.fontBold = Me.togFilterByPanelType
    Me.labPanelNameFilter.fontBold = Me.togFilterByPanelName
    Me.labYearFilter.fontBold = Me.togFilterByYear
    Me.labStartDateFilter.fontBold = Me.togFilterByRange
    Me.labEndDateFilter.fontBold = Me.togFilterByRange
    ' Update the font colors if filtering on that field
    If Me.togFilterByPark Then Me.cmbParkFilter.forecolor = 16711680 _
        Else Me.cmbParkFilter.forecolor = 0
    If Me.togFilterByType Then Me.cmbTypeFilter.forecolor = 16711680 _
        Else Me.cmbTypeFilter.forecolor = 0
    If Me.togFilterByStatus Then Me.cmbStatusFilter.forecolor = 16711680 _
        Else Me.cmbStatusFilter.forecolor = 0
    If Me.togFilterBySite Then Me.cmbSiteFilter.forecolor = 16711680 _
        Else Me.cmbSiteFilter.forecolor = 0
    If Me.togFilterByLoc Then Me.cmbLocFilter.forecolor = 16711680 _
        Else Me.cmbLocFilter.forecolor = 0
    If Me.togFilterByStratum Then Me.cmbStratumFilter.forecolor = 16711680 _
        Else Me.cmbStratumFilter.forecolor = 0
    If Me.togFilterByRegion Then Me.cmbRegionFilter.forecolor = 16711680 _
        Else Me.cmbRegionFilter.forecolor = 0
    If Me.togFilterByPanelType Then Me.cmbPanelTypeFilter.forecolor = 16711680 _
        Else Me.cmbPanelTypeFilter.forecolor = 0
    If Me.togFilterByPanelName Then Me.cmbPanelNameFilter.forecolor = 16711680 _
        Else Me.cmbPanelNameFilter.forecolor = 0
    If Me.togFilterByYear Then Me.cmbYearFilter.forecolor = 16711680 _
        Else Me.cmbYearFilter.forecolor = 0
    If Me.togFilterByRange Then
        Me.txtStartDateFilter.forecolor = 16711680
        Me.txtEndDateFilter.forecolor = 16711680
    Else
        Me.txtStartDateFilter.forecolor = 0
        Me.txtEndDateFilter.forecolor = 0
    End If
    ' Non-standard fields
    'Me.labWatershedFilter.FontBold = Me.togFilterByWatershed
    'If Me.togFilterByWatershed Then Me.cmbWatershedFilter.ForeColor = 16711680 _
    '    Else: Me.cmbWatershedFilter.ForeColor = 0

Exit_Procedure:
    Exit Function

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
        "Error encountered (#" & Err.Number & " - fxnFilterRecords)"
    Resume Exit_Procedure

End Function
