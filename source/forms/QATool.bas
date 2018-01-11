Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =2
    GridX =24
    GridY =24
    Width =14415
    DatasheetFontHeight =10
    ItemSuffix =697
    Left =4800
    Top =3600
    Right =15735
    Bottom =11415
    DatasheetGridlinesColor =12632256
    Filter ="[Query_name] = \"qa_a111_Overview_transect_pt_duplicates\" AND [Time_frame] = \""
        "2014\" AND [Data_scope] = 0"
    RecSrcDt = Begin
        0xdef19da9b06be340
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="tbl_QA_Results"
    Caption =" Data Validation and Quality Review Tool"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            OldBorderStyle =1
            TextAlign =1
            FontWeight =700
            BackColor =8388608
            BorderColor =8388608
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =2
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderWidth =2
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            BorderColor =16776960
        End
        Begin CommandButton
            FontItalic = NotDefault
            FontSize =8
            ForeColor =-2147483630
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =4
            BorderWidth =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =8388608
        End
        Begin CheckBox
            SpecialEffect =4
            BorderWidth =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =8388608
        End
        Begin OptionGroup
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
        End
        Begin BoundObjectFrame
            BorderLineStyle =0
            BackStyle =0
            BorderColor =16776960
        End
        Begin TextBox
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin ListBox
            BorderLineStyle =0
            BackColor =8421376
            ForeColor =16777215
            BorderColor =16776960
            FontName ="Arial"
        End
        Begin ComboBox
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =16776960
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =16776960
        End
        Begin ToggleButton
            FontItalic = NotDefault
            FontSize =8
            ForeColor =-2147483630
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin Tab
            FontItalic = NotDefault
            BackStyle =0
            FontWeight =700
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1020
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =0
                    Left =120
                    Width =3480
                    Height =300
                    FontSize =11
                    FontWeight =500
                    BackColor =16777215
                    BorderColor =8355711
                    Name ="lblTitle"
                    Caption ="QA Tool"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =300
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =0
                    Left =120
                    Top =360
                    Width =6840
                    Height =315
                    FontSize =11
                    FontWeight =400
                    BackColor =16777215
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Enter the contact information and click save."
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =360
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =675
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    OverlapFlags =93
                    Left =13440
                    Top =60
                    Width =720
                    Height =354
                    ForeColor =0
                    Name ="btnCloseX"
                    Caption ="Close"
                    ControlTipText ="Close the form"

                    LayoutCachedLeft =13440
                    LayoutCachedTop =60
                    LayoutCachedWidth =14160
                    LayoutCachedHeight =414
                    UseTheme =255
                    BackColor =14136213
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =14136213
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    HoverColor =65280
                    PressedColor =9592887
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeColor =4210752
                    HoverForeThemeColorIndex =0
                    HoverForeTint =75.0
                    PressedForeColor =4210752
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =85
                    Left =10980
                    Top =60
                    Width =1914
                    Height =355
                    TabIndex =1
                    BackColor =16777215
                    BorderColor =0
                    Name ="optgMode"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="Change the form mode"

                    LayoutCachedLeft =10980
                    LayoutCachedTop =60
                    LayoutCachedWidth =12894
                    LayoutCachedHeight =415
                    Begin
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =12060
                            Top =144
                            OptionValue =1
                            BorderColor =0
                            Name ="optEditMode"

                            LayoutCachedLeft =12060
                            LayoutCachedTop =144
                            LayoutCachedWidth =12320
                            LayoutCachedHeight =384
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =12294
                                    Top =120
                                    Width =390
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    Name ="lblEditMode"
                                    Caption ="Edit"
                                    ControlTipText ="Edit mode"
                                    LayoutCachedLeft =12294
                                    LayoutCachedTop =120
                                    LayoutCachedWidth =12684
                                    LayoutCachedHeight =390
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =11100
                            Top =150
                            TabIndex =1
                            OptionValue =0
                            BorderColor =0
                            Name ="optViewMode"

                            LayoutCachedLeft =11100
                            LayoutCachedTop =150
                            LayoutCachedWidth =11360
                            LayoutCachedHeight =390
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =11334
                                    Top =120
                                    Width =495
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    Name ="lblViewMode"
                                    Caption ="View"
                                    ControlTipText ="View mode"
                                    LayoutCachedLeft =11334
                                    LayoutCachedTop =120
                                    LayoutCachedWidth =11829
                                    LayoutCachedHeight =390
                                End
                            End
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =93
                    Left =9360
                    Top =540
                    Width =4800
                    Height =355
                    TabIndex =3
                    BackColor =16777215
                    BorderColor =0
                    Name ="optgScope"
                    DefaultValue ="0"
                    ControlTipText ="Scope of the data included in the validation queries: uncertified events, certif"
                        "ied events, or both?"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =540
                    LayoutCachedWidth =14160
                    LayoutCachedHeight =895
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =0
                            OldBorderStyle =0
                            OverlapFlags =215
                            TextAlign =0
                            Left =9420
                            Top =600
                            Width =945
                            Height =255
                            FontWeight =400
                            BackColor =13025979
                            BorderColor =0
                            Name ="lblIncludeCertified"
                            Caption ="Data scope:"
                            LayoutCachedLeft =9420
                            LayoutCachedTop =600
                            LayoutCachedWidth =10365
                            LayoutCachedHeight =855
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =10500
                            Top =624
                            OptionValue =0
                            BorderColor =0
                            Name ="optUncertOnly"

                            LayoutCachedLeft =10500
                            LayoutCachedTop =624
                            LayoutCachedWidth =10760
                            LayoutCachedHeight =864
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =10740
                                    Top =600
                                    Width =1050
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    Name ="lblUncertOnly"
                                    Caption ="Uncert. only"
                                    ControlTipText ="Run queries only on uncertified events"
                                    LayoutCachedLeft =10740
                                    LayoutCachedTop =600
                                    LayoutCachedWidth =11790
                                    LayoutCachedHeight =870
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =12000
                            Top =630
                            OptionValue =1
                            BorderColor =0
                            Name ="optBoth"

                            LayoutCachedLeft =12000
                            LayoutCachedTop =630
                            LayoutCachedWidth =12260
                            LayoutCachedHeight =870
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =12240
                                    Top =600
                                    Width =480
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    Name ="lblBoth"
                                    Caption ="Both"
                                    LayoutCachedLeft =12240
                                    LayoutCachedTop =600
                                    LayoutCachedWidth =12720
                                    LayoutCachedHeight =870
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =12960
                            Top =630
                            OptionValue =2
                            BorderColor =0
                            Name ="optCertOnly"

                            LayoutCachedLeft =12960
                            LayoutCachedTop =630
                            LayoutCachedWidth =13220
                            LayoutCachedHeight =870
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =13200
                                    Top =600
                                    Width =870
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    Name ="lblCertOnly"
                                    Caption ="Cert. only"
                                    ControlTipText ="Run queries only on certified events"
                                    LayoutCachedLeft =13200
                                    LayoutCachedTop =600
                                    LayoutCachedWidth =14070
                                    LayoutCachedHeight =870
                                End
                            End
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7500
                    Top =540
                    Width =1620
                    TabIndex =2
                    BackColor =16777215
                    BorderColor =0
                    ForeColor =0
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cbxTimeframe"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Forms]![frm_Switchboard]![cTimeframe] AS Timeframe FROM tbl_QA_Results  "
                        "UNION SELECT tbl_QA_Results.Time_frame FROM tbl_QA_Results GROUP BY tbl_QA_Resul"
                        "ts.Time_frame ORDER BY Timeframe DESC;"

                    LayoutCachedLeft =7500
                    LayoutCachedTop =540
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =0
                            OldBorderStyle =0
                            OverlapFlags =215
                            TextAlign =0
                            Left =4860
                            Top =540
                            Width =2520
                            Height =255
                            FontWeight =400
                            BackColor =13025979
                            BorderColor =0
                            Name ="lblTimeframe"
                            Caption ="Time frame of data being certified:"
                            LayoutCachedLeft =4860
                            LayoutCachedTop =540
                            LayoutCachedWidth =7380
                            LayoutCachedHeight =795
                        End
                    End
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7440
                    Width =720
                    FontSize =11
                    FontWeight =400
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnUndo"
                    Caption ="Edit"
                    FontName ="Calibri"
                    ControlTipText ="Undo/Clear values"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x505050f0404040ff202820ff000800ff00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x606060ff706870ff404040ff000800ff00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x605860ff909090ff606060ff302830ff00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x505850ffb0a8b0ff808080ff404840ff00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000002018200020182020 ,
                        0x605850ffc0c0c0ffa0a0a0ff404040ff20182030201820000000000000000000 ,
                        0x00000000000000000000000000000000000000002018200020182020505850ff ,
                        0xa0a0a0ffd0d0d0ffb0b0b0ff707070ff201820ff201820302018200000000000 ,
                        0x000000000000000000000000000000002018200020182020706870ffc0b8c0ff ,
                        0xe0e8e0ffe0e0e0ffc0c0c0ff909890ff605860ff201820ff2018203020182000 ,
                        0x0000000000000000000000002018200020182020707070ffc0c0c0fff0e8f0ff ,
                        0xfff8fffff0f0f0ffd0d8d0ffc0c0c0ffa098a0ff605860ff101810ff20182030 ,
                        0x20182000000000000000000020182020808080ffd0d0d0fff0f0f0ffffffffff ,
                        0xfffffffffff8ffffe0e8e0ffd0d8d0ffc0b8c0ff909090ff505050ff201820ff ,
                        0x201820300000000000000000808080ffd0d0d0fff0f0f0fffff8fffffff8ffff ,
                        0xf0f8f0fff0f0f0ffe0e8e0ffd0d0d0ffc0c0c0ffa098a0ff606860ff505850ff ,
                        0x101810ff0000000000000000b0b8b0ffc0c8c0ffd0d0d0ffd0d0d0ffc0c0c0ff ,
                        0xc0b8c0ffb0b0b0ffa0a8a0ffa0a0a0ffa098a0ff909090ff707870ff606060ff ,
                        0x504850ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =7440
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =14136213
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =14136213
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =40.0
                    PressedColor =9592887
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeColor =4210752
                    HoverForeThemeColorIndex =0
                    HoverForeTint =75.0
                    PressedForeColor =4210752
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =8760
                    Top =120
                    Width =600
                    Height =240
                    TabIndex =5
                    Name ="Toggle680"

                    LayoutCachedLeft =8760
                    LayoutCachedTop =120
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =360
                    UseTheme =255
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =13440
                    Top =300
                    Width =900
                    FontSize =11
                    FontWeight =400
                    TabIndex =6
                    ForeColor =4210752
                    Name ="btnClose"
                    Caption =" Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Close this form"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000060585090605850ff605850ff ,
                        0x605850ff605850900000000060585090605850ff605850ff605850ff60585090 ,
                        0x0000000000000000000000000000000000000000605850ffffffffb0ffffffb0 ,
                        0xffffffb0605850ff60585090605850ffffffffb0ffffffb0ffffffb0605850ff ,
                        0x000000000000000000000000000000000000000060585090605850ffffffffb0 ,
                        0xffffffb0ffffffb0605850ffffffffb0ffffffb0ffffffb0605850ff60585090 ,
                        0x00000000000000000000000000000000000000000000000060585090605850ff ,
                        0xffffffb0ffffffb0ffffffb0ffffffb0ffffffb0605850ff6058509000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000060585090 ,
                        0x605850ffffffffb0ffffffb0ffffffb0605850ff605850900000000000000000 ,
                        0x00000000000000000000000000000000000000000000000060585090605850ff ,
                        0xffffffb0ffffffb0ffffffb0ffffffb0ffffffb0605850ff6058509000000000 ,
                        0x000000000000000000000000000000000000000060585090605850ffffffffb0 ,
                        0xffffffb0ffffffb0605850ffffffffb0ffffffb0ffffffb0605850ff60585090 ,
                        0x0000000000000000000000000000000000000000605850ffffffffb0ffffffb0 ,
                        0xffffffb0605850ff60585090605850ffffffffb0ffffffb0ffffffb0605850ff ,
                        0x000000000000000000000000000000000000000060585090605850ff605850ff ,
                        0x605850ff605850900000000060585090605850ff605850ff605850ff60585090 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =13440
                    LayoutCachedTop =300
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =660
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =14136213
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =14136213
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =40.0
                    PressedColor =9592887
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeColor =4210752
                    HoverForeThemeColorIndex =0
                    HoverForeTint =75.0
                    PressedForeColor =4210752
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =11550
            BackColor =13025979
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Top =60
                    Width =14415
                    Height =11490
                    Name ="PageTabs"
                    OnChange ="[Event Procedure]"

                    LayoutCachedTop =60
                    LayoutCachedWidth =14415
                    LayoutCachedHeight =11550
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =120
                            Top =465
                            Width =14160
                            Height =10950
                            Name ="tabResults"
                            Caption =" Results Summary"
                            LayoutCachedLeft =120
                            LayoutCachedTop =465
                            LayoutCachedWidth =14280
                            LayoutCachedHeight =11415
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextAlign =0
                                    Left =120
                                    Top =465
                                    Width =3300
                                    Height =423
                                    FontWeight =400
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="lblOverview"
                                    Caption ="* Double-click on the label to change sort order.  Click on a query name to open"
                                        "."
                                    ControlTipText ="View mode"
                                    LayoutCachedLeft =120
                                    LayoutCachedTop =465
                                    LayoutCachedWidth =3420
                                    LayoutCachedHeight =888
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =223
                                    Left =9900
                                    Top =525
                                    Width =1500
                                    Height =300
                                    Name ="btnRefreshX"
                                    Caption ="Refresh results"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Run the validation queries and refresh the results summary"

                                    LayoutCachedLeft =9900
                                    LayoutCachedTop =525
                                    LayoutCachedWidth =11400
                                    LayoutCachedHeight =825
                                    UseTheme =255
                                    BackColor =14136213
                                    BackThemeColorIndex =4
                                    BackTint =60.0
                                    BorderColor =14136213
                                    BorderThemeColorIndex =4
                                    BorderTint =60.0
                                    HoverColor =65280
                                    PressedColor =9592887
                                    PressedThemeColorIndex =4
                                    PressedShade =75.0
                                    HoverForeColor =4210752
                                    HoverForeThemeColorIndex =0
                                    HoverForeTint =75.0
                                    PressedForeColor =4210752
                                    PressedForeThemeColorIndex =0
                                    PressedForeTint =75.0
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =223
                                    Left =11640
                                    Top =525
                                    Width =2100
                                    Height =300
                                    TabIndex =1
                                    Name ="btnViewSummaryX"
                                    Caption ="View summary report"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="View the quality review results as a report"

                                    LayoutCachedLeft =11640
                                    LayoutCachedTop =525
                                    LayoutCachedWidth =13740
                                    LayoutCachedHeight =825
                                    UseTheme =255
                                    BackColor =14136213
                                    BackThemeColorIndex =4
                                    BackTint =60.0
                                    BorderColor =14136213
                                    BorderThemeColorIndex =4
                                    BorderTint =60.0
                                    HoverColor =65280
                                    PressedColor =9592887
                                    PressedThemeColorIndex =4
                                    PressedShade =75.0
                                    HoverForeColor =4210752
                                    HoverForeThemeColorIndex =0
                                    HoverForeTint =75.0
                                    PressedForeColor =4210752
                                    PressedForeThemeColorIndex =0
                                    PressedForeTint =75.0
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin Subform
                                    CanShrink = NotDefault
                                    OverlapFlags =247
                                    Left =120
                                    Top =915
                                    Width =14160
                                    Height =10053
                                    TabIndex =2
                                    BorderColor =0
                                    Name ="subResults"
                                    SourceObject ="Form.QAToolResults"
                                    LinkChildFields ="Query_name;Time_frame;Data_scope"
                                    LinkMasterFields ="Query_name;Time_frame;Data_scope"

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =915
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =10968
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    RowSourceTypeInt =1
                                    SpecialEffect =2
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =5490
                                    Top =562
                                    Width =1170
                                    TabIndex =3
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    Name ="cbxTypeFilter"
                                    RowSourceType ="Value List"
                                    RowSource ="1;Critical;2;Warning;3;Information"
                                    ColumnWidths ="0;2160"
                                    StatusBarText ="Filter by query type"
                                    AfterUpdate ="[Event Procedure]"
                                    ControlTipText ="Filter by query type"

                                    LayoutCachedLeft =5490
                                    LayoutCachedTop =562
                                    LayoutCachedWidth =6660
                                    LayoutCachedHeight =802
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =4260
                                            Top =555
                                            Width =1110
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="lblTypeFilter"
                                            Caption ="Query type:"
                                            LayoutCachedLeft =4260
                                            LayoutCachedTop =555
                                            LayoutCachedWidth =5370
                                            LayoutCachedHeight =795
                                        End
                                    End
                                End
                                Begin ToggleButton
                                    FontItalic = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    Left =6780
                                    Top =525
                                    Width =480
                                    Height =300
                                    FontWeight =400
                                    TabIndex =4
                                    ForeColor =0
                                    Name ="tglFilterByType"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    Caption ="Filter on"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Turn the type filter on or off"
                                    ImageData = Begin
                                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x505050f0404040ff202820ff000800ff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x606060ff706870ff404040ff000800ff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x605860ff909090ff606060ff302830ff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x505850ffb0a8b0ff808080ff404840ff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000002018200020182020 ,
                                        0x605850ffc0c0c0ffa0a0a0ff404040ff20182030201820000000000000000000 ,
                                        0x00000000000000000000000000000000000000002018200020182020505850ff ,
                                        0xa0a0a0ffd0d0d0ffb0b0b0ff707070ff201820ff201820302018200000000000 ,
                                        0x000000000000000000000000000000002018200020182020706870ffc0b8c0ff ,
                                        0xe0e8e0ffe0e0e0ffc0c0c0ff909890ff605860ff201820ff2018203020182000 ,
                                        0x0000000000000000000000002018200020182020707070ffc0c0c0fff0e8f0ff ,
                                        0xfff8fffff0f0f0ffd0d8d0ffc0c0c0ffa098a0ff605860ff101810ff20182030 ,
                                        0x20182000000000000000000020182020808080ffd0d0d0fff0f0f0ffffffffff ,
                                        0xfffffffffff8ffffe0e8e0ffd0d8d0ffc0b8c0ff909090ff505050ff201820ff ,
                                        0x201820300000000000000000808080ffd0d0d0fff0f0f0fffff8fffffff8ffff ,
                                        0xf0f8f0fff0f0f0ffe0e8e0ffd0d0d0ffc0c0c0ffa098a0ff606860ff505850ff ,
                                        0x101810ff0000000000000000b0b8b0ffc0c8c0ffd0d0d0ffd0d0d0ffc0c0c0ff ,
                                        0xc0b8c0ffb0b0b0ffa0a8a0ffa0a0a0ffa098a0ff909090ff707870ff606060ff ,
                                        0x504850ff00000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000
                                    End

                                    LayoutCachedLeft =6780
                                    LayoutCachedTop =525
                                    LayoutCachedWidth =7260
                                    LayoutCachedHeight =825
                                    UseTheme =255
                                    BackColor =14136213
                                    BackThemeColorIndex =4
                                    BackTint =60.0
                                    BorderColor =14136213
                                    BorderThemeColorIndex =4
                                    BorderTint =60.0
                                    HoverColor =65280
                                    PressedColor =9592887
                                    PressedThemeColorIndex =4
                                    PressedShade =75.0
                                    HoverForeColor =4210752
                                    HoverForeThemeColorIndex =0
                                    HoverForeTint =75.0
                                    PressedForeColor =4210752
                                    PressedForeThemeColorIndex =0
                                    PressedForeTint =75.0
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
                                    SpecialEffect =2
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8220
                                    Top =562
                                    Width =900
                                    TabIndex =5
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    Name ="cbxDoneFilter"
                                    RowSourceType ="Value List"
                                    RowSource ="True;False"
                                    StatusBarText ="Filter by the 'Done' flag"
                                    AfterUpdate ="[Event Procedure]"
                                    ControlTipText ="Filter by the 'Done' flag"

                                    LayoutCachedLeft =8220
                                    LayoutCachedTop =562
                                    LayoutCachedWidth =9120
                                    LayoutCachedHeight =802
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =7500
                                            Top =562
                                            Width =600
                                            Height =228
                                            FontWeight =400
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="lblDoneFilter"
                                            Caption ="Done:"
                                            LayoutCachedLeft =7500
                                            LayoutCachedTop =562
                                            LayoutCachedWidth =8100
                                            LayoutCachedHeight =790
                                        End
                                    End
                                End
                                Begin ToggleButton
                                    FontItalic = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    Left =9240
                                    Top =525
                                    Width =480
                                    Height =300
                                    FontWeight =400
                                    TabIndex =6
                                    ForeColor =0
                                    Name ="tglFilterByDone"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    Caption ="Filter on"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Turn the 'Done' filter on or off"
                                    ImageData = Begin
                                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x505050f0404040ff202820ff000800ff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x606060ff706870ff404040ff000800ff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x605860ff909090ff606060ff302830ff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x505850ffb0a8b0ff808080ff404840ff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000002018200020182020 ,
                                        0x605850ffc0c0c0ffa0a0a0ff404040ff20182030201820000000000000000000 ,
                                        0x00000000000000000000000000000000000000002018200020182020505850ff ,
                                        0xa0a0a0ffd0d0d0ffb0b0b0ff707070ff201820ff201820302018200000000000 ,
                                        0x000000000000000000000000000000002018200020182020706870ffc0b8c0ff ,
                                        0xe0e8e0ffe0e0e0ffc0c0c0ff909890ff605860ff201820ff2018203020182000 ,
                                        0x0000000000000000000000002018200020182020707070ffc0c0c0fff0e8f0ff ,
                                        0xfff8fffff0f0f0ffd0d8d0ffc0c0c0ffa098a0ff605860ff101810ff20182030 ,
                                        0x20182000000000000000000020182020808080ffd0d0d0fff0f0f0ffffffffff ,
                                        0xfffffffffff8ffffe0e8e0ffd0d8d0ffc0b8c0ff909090ff505050ff201820ff ,
                                        0x201820300000000000000000808080ffd0d0d0fff0f0f0fffff8fffffff8ffff ,
                                        0xf0f8f0fff0f0f0ffe0e8e0ffd0d0d0ffc0c0c0ffa098a0ff606860ff505850ff ,
                                        0x101810ff0000000000000000b0b8b0ffc0c8c0ffd0d0d0ffd0d0d0ffc0c0c0ff ,
                                        0xc0b8c0ffb0b0b0ffa0a8a0ffa0a0a0ffa098a0ff909090ff707870ff606060ff ,
                                        0x504850ff00000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000
                                    End

                                    LayoutCachedLeft =9240
                                    LayoutCachedTop =525
                                    LayoutCachedWidth =9720
                                    LayoutCachedHeight =825
                                    UseTheme =255
                                    BackColor =14136213
                                    BackThemeColorIndex =4
                                    BackTint =60.0
                                    BorderColor =14136213
                                    BorderThemeColorIndex =4
                                    BorderTint =60.0
                                    HoverColor =65280
                                    PressedColor =9592887
                                    PressedThemeColorIndex =4
                                    PressedShade =75.0
                                    HoverForeColor =4210752
                                    HoverForeThemeColorIndex =0
                                    HoverForeTint =75.0
                                    PressedForeColor =4210752
                                    PressedForeThemeColorIndex =0
                                    PressedForeTint =75.0
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =223
                                    Left =11520
                                    Top =720
                                    Width =2460
                                    FontSize =11
                                    FontWeight =400
                                    TabIndex =7
                                    ForeColor =4210752
                                    Name ="btnViewSummary"
                                    Caption ="View Summary Report"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    ControlTipText ="View the quality review results as a report"
                                    GridlineColor =10921638
                                    ImageData = Begin
                                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000000000000000000000000000d0c8d0ff505050ff306830ff607060ff ,
                                        0x506050ff405040ff304030ff203020ff202020ff404840ff808880ff00000000 ,
                                        0x00000000000000000000000000000000303030ffd0d8d0ff407040ff70b870ff ,
                                        0x60a060ff609860ff509050ff508850ff507850ff406040ff505050ff00000000 ,
                                        0x00000000000000000000000000000000b0b8b0ff605860ff407850ff80c080ff ,
                                        0x70b870ff70b070ff60a860ff50a050ff509050ff507850ff101810ff00000000 ,
                                        0x00000000000000000000000000000000202820ffffffffff508850ff80c890ff ,
                                        0x70c080ff70b870ff60b070ff60a860ff509850ff508050ff101810ff00000000 ,
                                        0x00000000000000000000000000000000a0a8a0ff606060ff509060ff90c890ff ,
                                        0x80c880ff70c080ff70b870ff60b060ff60a860ff508850ff202020ff00000000 ,
                                        0x00000000000000000000000000000000202820ffffffffff509060ffa0d0a0ff ,
                                        0x90c890ff80c880ff70c080ff70b870ff60b060ff509860ff303830ff00000000 ,
                                        0x00000000000000000000000000000000b0b0b0ff707070ff609870ffa0d8b0ff ,
                                        0x408850ff509850ff50a860ff60b860ff60c070ff60a060ff405040ff00000000 ,
                                        0x00000000000000000000000000000000202020ffffffffff70a080ffb0e0c0ff ,
                                        0x408850ffb0e0c0ffb0e0c0ffb0e0c0ff60b860ff60a060ff506050ff00000000 ,
                                        0x00000000000000000000000000000000a0a8a0ff808080ff70b080ffc0e0c0ff ,
                                        0x408850ff408850ff408850ff408850ff60b860ff60a860ff607060ff00000000 ,
                                        0x00000000000000000000000000000000202820ffffffffff80b890ffc0e0c0ff ,
                                        0xc0e0c0ffb0d8b0ffa0d0a0ff90c890ff80c880ff70b870ff709870ff00000000 ,
                                        0x00000000000000000000000000000000e0e8e0ff909090ff90c8a0ff90c0a0ff ,
                                        0x80b890ff80b880ff70b080ff70a870ff60a070ff70a870ffa0c0a0ff00000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000
                                    End

                                    LayoutCachedLeft =11520
                                    LayoutCachedTop =720
                                    LayoutCachedWidth =13980
                                    LayoutCachedHeight =1080
                                    PictureCaptionArrangement =5
                                    ForeThemeColorIndex =0
                                    ForeTint =75.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    Shape =1
                                    Gradient =12
                                    BackColor =14136213
                                    BackThemeColorIndex =4
                                    BackTint =60.0
                                    BorderColor =14136213
                                    BorderThemeColorIndex =4
                                    BorderTint =60.0
                                    ThemeFontIndex =1
                                    HoverColor =65280
                                    HoverTint =40.0
                                    PressedColor =9592887
                                    PressedThemeColorIndex =4
                                    PressedShade =75.0
                                    HoverForeColor =4210752
                                    HoverForeThemeColorIndex =0
                                    HoverForeTint =75.0
                                    PressedForeColor =4210752
                                    PressedForeThemeColorIndex =0
                                    PressedForeTint =75.0
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =223
                                    Left =9780
                                    Top =720
                                    Width =1980
                                    FontSize =11
                                    FontWeight =400
                                    TabIndex =8
                                    ForeColor =4210752
                                    Name ="btnRefresh"
                                    Caption =" Refresh Results"
                                    FontName ="Calibri"
                                    ControlTipText ="Run the validation queries and refresh the results summary"
                                    GridlineColor =10921638
                                    ImageData = Begin
                                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0xb0886090805830f0804800ff905800ff804800ff805820d00000000000000000 ,
                                        0x000000000000000000000000000000000000000000000000b0803000a0805090 ,
                                        0x906020ffc07820ffa05800ffb08040c000000000000000000000000000000000 ,
                                        0x000000000000000000000000000000000000000000000000a0783010a06820ff ,
                                        0xe09830ffb06000ffa07030c0a070301000000000804800000000000000000000 ,
                                        0x0000000000000000000000000000000000000000e09030ffd08830ffe09840ff ,
                                        0xf0a840ffa05810ff704000ff603000ff0000000000000000503000ff00000000 ,
                                        0x000000000000000000000000000000000000000000000000e09830ffffc880ff ,
                                        0xf0b050ffd08030ff905810ff0000000000000000805020ffb06810ff804000ff ,
                                        0x000000000000000020a0f0ff4078e0ff4078e0ff4078e0ff4078e0ffe09830ff ,
                                        0xffc070ffa06820ff4070e0ff4070e0ff906830fff0a850ffe09020ffb06000ff ,
                                        0x704000ff0000000020a0f0ff50c0ffff50c0ffff4078e0ff50c0ffff50c0ffff ,
                                        0xc08030ff50c0ffff50c0ffffd08840ffe09030ffffb860ffffb040ffc07810ff ,
                                        0x904800ff704000ff20a0f0ff70d8ffff50c8ffff3080e0ff70d8ffffc0b0a0ff ,
                                        0xc0b0a0ffc0b0a0ffc0b8b0ffd0c8c0ffd0b8a0ffd07820ffffb050ffa06010ff ,
                                        0xa08050700000000020a8f0ff80d8ffff60d0ffff3080e0ff80d8ffff60d0ffff ,
                                        0xa098a0ff90e0ffffd0a880ffc09060ffc07830ffffb050ffb06820ffb08050a0 ,
                                        0x000000000000000020a8f0ff80e0ffff70d0f0ff3080f0ff80e0ffff70d0f0ff ,
                                        0xa098a0ffd08040ffd07030ffd07030ffd07830ffd08040ffd08860ff00000000 ,
                                        0x000000000000000020a8f0ff90e8ffff70d8f0ff3080f0ff90e8ffff70d8f0ff ,
                                        0xa098a0ff90e8ffffe0c8b0fff0d0d0fff0c8b0ffffd8c0ffd0a080ff00000000 ,
                                        0x000000000000000020a8f0f090e8fff070d8f0f03080f0f090e8fff070d8f0f0 ,
                                        0x3080f0f090e8fff070d8f0f03080f0f090e8fff070d8f0f04078e0f000000000 ,
                                        0x000000000000000020a8f0ffb0f0ffffb0f0ffff3088f0ffb0f0ffffb0f0ffff ,
                                        0x3088f0ffb0f0ffffb0f0ffff3088f0ffb0f0ffffb0f0ffff4078e0ff00000000 ,
                                        0x000000000000000020a8f0ff3090f0ff3088f0ff3088f0ff3088f0ff3088f0ff ,
                                        0x3080f0ff3080f0ff3080f0ff3080f0ff4080e0ff4078e0ff4078e0ff00000000 ,
                                        0x000000000000000020a8ffff80e8ffff80e8ffff80e8ffff80e8ffff70e0ffff ,
                                        0x70d8ffff70d0ffff60c8ffff60c0ffff50b8ffff50b0ffff4078e0ff00000000 ,
                                        0x000000000000000030a8ffff30a8ffff30a8ffff30a8ffff30a8ffff30a0ffff ,
                                        0x30a0ffff30a0ffff3098f0ff3098f0ff3090f0ff3090f0ff3090f0ff00000000 ,
                                        0x0000000000000000
                                    End

                                    LayoutCachedLeft =9780
                                    LayoutCachedTop =720
                                    LayoutCachedWidth =11760
                                    LayoutCachedHeight =1080
                                    PictureCaptionArrangement =5
                                    ForeThemeColorIndex =0
                                    ForeTint =75.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    Shape =1
                                    Gradient =12
                                    BackColor =14136213
                                    BackThemeColorIndex =4
                                    BackTint =60.0
                                    BorderColor =14136213
                                    BorderThemeColorIndex =4
                                    BorderTint =60.0
                                    ThemeFontIndex =1
                                    HoverColor =65280
                                    HoverTint =40.0
                                    PressedColor =9592887
                                    PressedThemeColorIndex =4
                                    PressedShade =75.0
                                    HoverForeColor =4210752
                                    HoverForeThemeColorIndex =0
                                    HoverForeTint =75.0
                                    PressedForeColor =4210752
                                    PressedForeThemeColorIndex =0
                                    PressedForeTint =75.0
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =465
                            Width =14160
                            Height =10950
                            Name ="tabQueries"
                            Caption =" View/Fix Query Results"
                            LayoutCachedLeft =120
                            LayoutCachedTop =465
                            LayoutCachedWidth =14280
                            LayoutCachedHeight =11415
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =255
                                    Left =8040
                                    Top =600
                                    Width =1320
                                    Height =317
                                    Name ="btnDesignViewX"
                                    Caption ="Design view"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open the selected query in design view"

                                    LayoutCachedLeft =8040
                                    LayoutCachedTop =600
                                    LayoutCachedWidth =9360
                                    LayoutCachedHeight =917
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =1260
                                    Top =615
                                    Width =6660
                                    Height =252
                                    TabIndex =1
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                                    Name ="cbxObject"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT MSysObjects.Name AS Query_name FROM MSysObjects WHERE (((MSysObjects.Name"
                                        ") Like \"qa_*\") AND ((MSysObjects.Type)=5)) ORDER BY MSysObjects.Name; "
                                    AfterUpdate ="[Event Procedure]"

                                    LayoutCachedLeft =1260
                                    LayoutCachedTop =615
                                    LayoutCachedWidth =7920
                                    LayoutCachedHeight =867
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =120
                                            Top =615
                                            Width =1110
                                            Height =270
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="lblObject"
                                            Caption ="Query name"
                                            LayoutCachedLeft =120
                                            LayoutCachedTop =615
                                            LayoutCachedWidth =1230
                                            LayoutCachedHeight =885
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =10080
                                    Top =615
                                    Width =2220
                                    Height =252
                                    TabIndex =2
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="tbxUser"
                                    ControlSource ="QA_user"
                                    OnDirty ="[Event Procedure]"

                                    LayoutCachedLeft =10080
                                    LayoutCachedTop =615
                                    LayoutCachedWidth =12300
                                    LayoutCachedHeight =867
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =9480
                                            Top =615
                                            Width =570
                                            Height =270
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labUser"
                                            Caption ="QA by"
                                            LayoutCachedLeft =9480
                                            LayoutCachedTop =615
                                            LayoutCachedWidth =10050
                                            LayoutCachedHeight =885
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =12360
                                    Top =615
                                    Width =1920
                                    Height =252
                                    TabIndex =3
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="tbxRemedy_date"
                                    ControlSource ="Remedy_date"
                                    Format ="mm/dd/yy"

                                    LayoutCachedLeft =12360
                                    LayoutCachedTop =615
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =867
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    EnterKeyBehavior = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    ScrollBars =2
                                    SpecialEffect =2
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =1260
                                    Top =975
                                    Width =13020
                                    Height =660
                                    TabIndex =4
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="tbxQueryDesc"
                                    ControlSource ="Query_description"
                                    StatusBarText ="Description of the query"
                                    OnDirty ="[Event Procedure]"

                                    LayoutCachedLeft =1260
                                    LayoutCachedTop =975
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =1635
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =120
                                            Top =975
                                            Width =1035
                                            Height =495
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="lblQueryDesc"
                                            Caption ="Query description"
                                            LayoutCachedLeft =120
                                            LayoutCachedTop =975
                                            LayoutCachedWidth =1155
                                            LayoutCachedHeight =1470
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    EnterKeyBehavior = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    ScrollBars =2
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =1260
                                    Top =1755
                                    Width =13020
                                    Height =810
                                    TabIndex =5
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="tbxRemedy"
                                    ControlSource ="Remedy_desc"
                                    StatusBarText ="Details about actions taken and/or not taken to resolve errors"

                                    LayoutCachedLeft =1260
                                    LayoutCachedTop =1755
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =2565
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =120
                                            Top =1755
                                            Width =810
                                            Height =495
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="lblRemedy"
                                            Caption ="Remedy details"
                                            LayoutCachedLeft =120
                                            LayoutCachedTop =1755
                                            LayoutCachedWidth =930
                                            LayoutCachedHeight =2250
                                        End
                                    End
                                End
                                Begin Subform
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    SpecialEffect =2
                                    Left =120
                                    Top =3120
                                    Width =14160
                                    Height =7845
                                    TabIndex =6
                                    BorderColor =0
                                    Name ="subQueryResults"

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =3120
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =10965
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =255
                                            TextAlign =0
                                            Left =120
                                            Top =2880
                                            Width =1212
                                            Height =252
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="lblQueryResults"
                                            Caption ="Query results"
                                            LayoutCachedLeft =120
                                            LayoutCachedTop =2880
                                            LayoutCachedWidth =1332
                                            LayoutCachedHeight =3132
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    FELineBreak = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3270
                                    Top =2745
                                    Width =606
                                    Height =255
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =7
                                    BackColor =8454143
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="tbxEditQuery"
                                    FontName ="Tahoma"
                                    AsianLineBreak =255

                                    LayoutCachedLeft =3270
                                    LayoutCachedTop =2745
                                    LayoutCachedWidth =3876
                                    LayoutCachedHeight =3000
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1440
                                            Top =2745
                                            Width =1770
                                            Height =255
                                            FontSize =9
                                            FontWeight =400
                                            BackColor =16777215
                                            BorderColor =0
                                            ForeColor =0
                                            Name ="lblEditQuery"
                                            Caption ="Edit results directly?"
                                            FontName ="Tahoma"
                                            LayoutCachedLeft =1440
                                            LayoutCachedTop =2745
                                            LayoutCachedWidth =3210
                                            LayoutCachedHeight =3000
                                        End
                                    End
                                End
                                Begin CommandButton
                                    Enabled = NotDefault
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =4620
                                    Top =2685
                                    Width =1080
                                    Height =317
                                    TabIndex =8
                                    ForeColor =0
                                    Name ="btnAutoFixX"
                                    Caption ="Auto-fix"
                                    StatusBarText ="Run a pre-built query to automatically fix all the records"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Run a pre-built query to automatically fix all the records"

                                    LayoutCachedLeft =4620
                                    LayoutCachedTop =2685
                                    LayoutCachedWidth =5700
                                    LayoutCachedHeight =3002
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =6240
                                    Top =2685
                                    Width =2040
                                    Height =317
                                    TabIndex =9
                                    ForeColor =0
                                    Name ="btnOpenRecordX"
                                    Caption ="Open selected record"
                                    StatusBarText ="Open the form / query / table specified in the query to the record selected in t"
                                        "he subform"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open the form / query / table specified in the query to the record selected in t"
                                        "he subform"

                                    LayoutCachedLeft =6240
                                    LayoutCachedTop =2685
                                    LayoutCachedWidth =8280
                                    LayoutCachedHeight =3002
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =8580
                                    Top =2685
                                    Height =317
                                    TabIndex =10
                                    ForeColor =0
                                    Name ="btnOpenBrowserX"
                                    Caption ="Data browser"
                                    StatusBarText ="Open the project data browser"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open the project data browser"

                                    LayoutCachedLeft =8580
                                    LayoutCachedTop =2685
                                    LayoutCachedWidth =10020
                                    LayoutCachedHeight =3002
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =10500
                                    Top =2626
                                    Width =426
                                    Height =426
                                    FontWeight =400
                                    TabIndex =11
                                    Name ="btnExportX"
                                    Caption ="Export to Excel"
                                    StatusBarText ="Export the results of the selected query to Excel"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadada0000000dadadadadd00000dadadadada ,
                                        0xad000dadadadadaddad0dadadadadadaadadadadad72727ddada2727272f272a ,
                                        0xadad727272f272addada27272f2727daadada272f27272addadada2f2727dada ,
                                        0xadada2f272727daddada2f27272727daadad72727d7272addada2727dad727da ,
                                        0xadadadadadadadad
                                    End
                                    FontName ="Tahoma"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Export the results of the selected query to Excel"

                                    LayoutCachedLeft =10500
                                    LayoutCachedTop =2626
                                    LayoutCachedWidth =10926
                                    LayoutCachedHeight =3052
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =11040
                                    Top =2625
                                    Width =426
                                    Height =426
                                    FontWeight =400
                                    TabIndex =12
                                    Name ="btnCloseupX"
                                    Caption ="Zoom"
                                    StatusBarText ="Open the selected query in a new window"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadada00adadadadadadad000adadadadadada ,
                                        0xa000adadadadadadda000a700007dadaada0000888800daddada07ee888870da ,
                                        0xada708e88888807ddad08e888888880aada088888888880ddad088888888e80a ,
                                        0xada088888888e80ddad70888888ee07aadad07888eee70addadad00888800ada ,
                                        0xadadad700007adad
                                    End
                                    FontName ="Tahoma"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Open the selected query in a new window"

                                    LayoutCachedLeft =11040
                                    LayoutCachedTop =2625
                                    LayoutCachedWidth =11466
                                    LayoutCachedHeight =3051
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =12600
                                    Top =2685
                                    Width =1020
                                    Height =317
                                    TabIndex =13
                                    ForeColor =0
                                    Name ="btnRequeryX"
                                    Caption ="Requery"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Requery the results set for the selected query"

                                    LayoutCachedLeft =12600
                                    LayoutCachedTop =2685
                                    LayoutCachedWidth =13620
                                    LayoutCachedHeight =3002
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    Left =8460
                                    Top =660
                                    Width =720
                                    FontSize =11
                                    FontWeight =400
                                    TabIndex =14
                                    ForeColor =4210752
                                    Name ="btnDesignView"
                                    Caption ="Design View"
                                    FontName ="Calibri"
                                    ControlTipText ="Open the selected query in design view"
                                    GridlineColor =10921638
                                    ImageData = Begin
                                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000002098c0ff107090ff106880ff106880ff006080ff ,
                                        0x006070ff005870ff005870ff005060ff005060ff004860ff004860ff004050ff ,
                                        0x004050ff004050ff0000000020a0c0ff80d0ffff80d0ffff80d0ffff70d0ffff ,
                                        0x70c8ffff60c0ffff60c0ffff50b8ffff50b0ffff40a8ffff30a0f0ff30a0f0ff ,
                                        0x3098f0ff004050ff0000000020a0d0ff80d0ffff2090b0ff80d0ffff2090b0ff ,
                                        0x70d0ffff2088b0ff60c0ffff1080a0ff50b8ffff1078a0ff40a8ffff107090ff ,
                                        0x30a0f0ff005060ff0000000020a0d0ff20a0d0ff2090b0ff20a0d0ff2090b0ff ,
                                        0x20a0c0ff2088b0ff2098c0ff1080a0ff1088b0ff1078a0ff1080a0ff107090ff ,
                                        0x107890ff007090ffd07040ffd07040ffd07040ffd06840ffc06030ffb05830ff ,
                                        0xa05020ffa04820ff904010ff904010ff903810ff000000000000000000000000 ,
                                        0x0000000000000000d07040ffffa080fff08050fff07840ffe07040ffe07030ff ,
                                        0x707070ff505850ff000000ffc07050fff0905010000000000000000000000000 ,
                                        0x0000000000000000d07040ffffb090ff903810ff904020ffb05020ffc06030ff ,
                                        0xa0a0a0ffffffffff5090b0ff101010ff30607030000000000000000000000000 ,
                                        0x0000000000000000d07850ffffb890ffa04820ffc0603000d07040fff08050ff ,
                                        0xa0a0a0ff90b8c0ff70d0e0ff5098b0ff101010ff306070300000000000000000 ,
                                        0x0000000000000000e07850ffffc0a0ffc06030ffd07050ffff8850ffff9860ff ,
                                        0xc08060ff50a0b0ff90e0f0ff60c0d0ff5098b0ff101010ff3060704000000000 ,
                                        0x0000000000000000e08060ffffc8a0ffd07040ffffa870ffffa070ffd07850ff ,
                                        0xf090502080a0b04050a0b0ff90e0f0ff60c0d0ff5098b0ff101010ff30607040 ,
                                        0x0000000000000000e08860ffffc8a0ffffb890ffffb080ffd07850fff0905020 ,
                                        0x000000000000000080a0b04060a8b0ff90e0f0ff60c0d0ff5098b0ff101010ff ,
                                        0x3058603000000000e09070ffffc8a0ffffb890ffe08850fff090502000000000 ,
                                        0x00000000000000000000000080a0b04070b0c0ff90e0f0ff70c8e0ff808880ff ,
                                        0x303890ff30388050e09870ffffc0a0ffe09070fff09050200000000000000000 ,
                                        0x0000000000000000000000000000000080a0b04080b0c0ffd0b8b0ff7088d0ff ,
                                        0x6070b0ff303890ffe09880ffe0a080fff0905020000000000000000000000000 ,
                                        0x000000000000000000000000000000000000000080a0b0406070b0ff7090e0ff ,
                                        0x6078d0ff6070b0ffe0a080fff090500000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000007080c0506070b0ff ,
                                        0x6070b0ff6078c030
                                    End

                                    LayoutCachedLeft =8460
                                    LayoutCachedTop =660
                                    LayoutCachedWidth =9180
                                    LayoutCachedHeight =1020
                                    ForeThemeColorIndex =0
                                    ForeTint =75.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    Shape =1
                                    Gradient =12
                                    BackColor =14136213
                                    BackThemeColorIndex =4
                                    BackTint =60.0
                                    BorderColor =14136213
                                    BorderThemeColorIndex =4
                                    BorderTint =60.0
                                    ThemeFontIndex =1
                                    HoverColor =65280
                                    HoverTint =40.0
                                    PressedColor =9592887
                                    PressedThemeColorIndex =4
                                    PressedShade =75.0
                                    HoverForeColor =4210752
                                    HoverForeThemeColorIndex =0
                                    HoverForeTint =75.0
                                    PressedForeColor =4210752
                                    PressedForeThemeColorIndex =0
                                    PressedForeTint =75.0
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =465
                            Width =14160
                            Height =10950
                            Name ="tabDataTables"
                            Caption =" Browse data tables"
                            LayoutCachedLeft =120
                            LayoutCachedTop =465
                            LayoutCachedWidth =14280
                            LayoutCachedHeight =11415
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListRows =20
                                    ListWidth =11520
                                    Left =840
                                    Top =615
                                    Width =4320
                                    Height =252
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"0\";\"0\""
                                    Name ="cbxTable"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tsys_Link_Tables.Link_table, tsys_Link_Tables.Description_text FROM tsys_"
                                        "Link_Tables WHERE (((tsys_Link_Tables.Link_table) Like \"tbl_*\" And (tsys_Link_"
                                        "Tables.Link_table)<>\"tbl_QA_Results\")) OR (((tsys_Link_Tables.Link_table)=\"tl"
                                        "u_Project_Crew\")) OR (((tsys_Link_Tables.Link_table)=\"tlu_Project_Taxa\")) OR "
                                        "(((tsys_Link_Tables.Link_table)=\"tlu_Park_Taxa\")); "
                                    ColumnWidths ="4320;7200"
                                    AfterUpdate ="[Event Procedure]"

                                    LayoutCachedLeft =840
                                    LayoutCachedTop =615
                                    LayoutCachedWidth =5160
                                    LayoutCachedHeight =867
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =180
                                            Top =615
                                            Width =585
                                            Height =270
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labTable"
                                            Caption ="Table:"
                                            LayoutCachedLeft =180
                                            LayoutCachedTop =615
                                            LayoutCachedWidth =765
                                            LayoutCachedHeight =885
                                        End
                                    End
                                End
                                Begin Subform
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    SpecialEffect =2
                                    Left =120
                                    Top =1263
                                    Width =14160
                                    Height =9705
                                    TabIndex =1
                                    BorderColor =0
                                    Name ="subDataTables"

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =1263
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =10968
                                End
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =0
                                    Left =5340
                                    Top =465
                                    Width =7716
                                    Height =699
                                    FontWeight =400
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labEditWarning"
                                    Caption =" Warning:  This is a last resort!  If possible, open the records needing fixes w"
                                        "ithin the data entry form.  Also, when making manual edits in data tables, pleas"
                                        "e be sure to update the updated_date and updated_by fields if they are present i"
                                        "n the table."
                                    ControlTipText ="View mode"
                                    LayoutCachedLeft =5340
                                    LayoutCachedTop =465
                                    LayoutCachedWidth =13056
                                    LayoutCachedHeight =1164
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =5880
                    Top =2400
                    Width =2220
                    FontSize =11
                    FontWeight =400
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnOpenRecord"
                    Caption ="Open Selected Record"
                    StatusBarText ="Open the form / query / table specified in the query to the record selected in t"
                        "he subform"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Open the form / query / table specified in the query to the record selected in t"
                        "he subform"
                    GridlineColor =10921638

                    LayoutCachedLeft =5880
                    LayoutCachedTop =2400
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =2760
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =14136213
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =14136213
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =40.0
                    PressedColor =9592887
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeColor =4210752
                    HoverForeThemeColorIndex =0
                    HoverForeTint =75.0
                    PressedForeColor =4210752
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Enabled = NotDefault
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =4140
                    Top =2640
                    Width =1140
                    FontSize =11
                    FontWeight =400
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnAutoFix"
                    Caption ="Auto-Fix"
                    StatusBarText ="Run a pre-built query to automatically fix all the records"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Run a pre-built query to automatically fix all the records"
                    GridlineColor =10921638

                    LayoutCachedLeft =4140
                    LayoutCachedTop =2640
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =3000
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =14136213
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =14136213
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =40.0
                    PressedColor =9592887
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeColor =4210752
                    HoverForeThemeColorIndex =0
                    HoverForeTint =75.0
                    PressedForeColor =4210752
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =8460
                    Top =2280
                    Width =1740
                    FontSize =11
                    FontWeight =400
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnOpenBrowser"
                    Caption =" Data Browser"
                    StatusBarText ="Open the project data browser"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Open the project data browser"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000b0a090ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff000000000000000000000000 ,
                        0x0000000000000000b0a090fffff8f0ffe0e0e0ffe0d8d0ffe0d8d0ffe0d0c0ff ,
                        0xe0c8c0ffd0c0b0ffd0c0b0ffd0b8a0ff604830ff000000000000000000000000 ,
                        0x0000000000000000b0a090ffffffffffd0b8b0ffd0b8a0ffd0b0a0ffb0a090ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ffb0a090fffffffffffffffffffff8f0fffff0f0ffb0a090ff ,
                        0xfff8f0ffe0e0e0ffe0d8d0ffe0d8d0ffe0d0c0ffe0c8c0ffd0c0b0ffd0c0b0ff ,
                        0xd0b8a0ff604830ffc0a8a0ffffffffffe0c8c0ffe0c8c0ffd0c0b0ffb0a090ff ,
                        0xffffffffd0b8b0ffd0b8a0ffd0b0a0ffffe8e0ffc09080ffc09080ffc09080ff ,
                        0xd0b8b0ff604830ffc0b0a0ffffffffffb0a090ff604830ff604830ffb0a090ff ,
                        0xfffffffffffffffffff8f0fffff0f0fffff0e0ffffe8e0ffffe0d0ffffd8d0ff ,
                        0xd0c0b0ff604830ffd0b0a0ffffffffffb0a090fffff8f0ffe0e0e0ffc0a8a0ff ,
                        0xffffffffe0c8c0ffe0c8c0ffd0c0b0fffff8f0ffc0a890ffc0a890ffc0a090ff ,
                        0xe0d0c0ff604830ffd0b8a0ffffffffffb0a090ffffffffffd0b8b0ffc0b0a0ff ,
                        0xfffffffffffffffffffffffffffffffffffffffffff8f0fffff0f0ffffe8e0ff ,
                        0xe0d8d0ff604830fff0a890fff0a880ffb0a090ffffffffffffffffffd0b0a0ff ,
                        0xffffffffe0c8c0ffe0c8c0ffe0c8c0ffffffffffd0b0a0ffd0b0a0ffd0a8a0ff ,
                        0xe0e0e0ff604830fff0a890ffffc0a0ffc0a8a0ffffffffffe0c8c0ffd0b8a0ff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffff8ffff ,
                        0xfff8f0ff604830fff0a890fff0a880ffc0b0a0fffffffffffffffffff0a890ff ,
                        0xf0a880fff0a080fff0a070ffe09870ffe09060ffe08860ffe08050ffe07840ff ,
                        0xe07840ffd06030ff0000000000000000d0b0a0ffffffffffe0c8c0fff0a890ff ,
                        0xffc0a0ffffc0a0ffffb890ffffb890ffffb090ffffa880fff0a070fff0a070ff ,
                        0xf09870ffd06830ff0000000000000000d0b8a0fffffffffffffffffff0a890ff ,
                        0xf0a880fff0a080fff0a080ffe09870ffe09060ffe08860ffe08850ffe08050ff ,
                        0xe07840ffe07840ff0000000000000000f0a890fff0a880fff0a080fff0a070ff ,
                        0xe09870ffe09060ffe08860ffe08050ffe07840ffe07840ffd06030ff00000000 ,
                        0x00000000000000000000000000000000f0a890ffffc0a0ffffc0a0ffffb890ff ,
                        0xffb890ffffb090ffffa880fff0a070fff0a070fff09870ffd06830ff00000000 ,
                        0x00000000000000000000000000000000f0a890fff0a880fff0a080fff0a080ff ,
                        0xe09870ffe09060ffe08860ffe08850ffe08050ffe07840ffe07840ff00000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =8460
                    LayoutCachedTop =2280
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =2640
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =14136213
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =14136213
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =40.0
                    PressedColor =9592887
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeColor =4210752
                    HoverForeThemeColorIndex =0
                    HoverForeTint =75.0
                    PressedForeColor =4210752
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =10440
                    Top =2160
                    Width =420
                    FontSize =11
                    FontWeight =400
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnExport"
                    Caption ="Export to Excel"
                    StatusBarText ="Export the results of the selected query to Excel"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Export the results of the selected query to Excel"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c08060f0905030ff905830ff904820ff ,
                        0x804020ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000c08060f0905030ffc07860ff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000c08870ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000208030ff207820ff207820ff ,
                        0x207820ff107020ff107020ff107010ff106810ff106810ff106810ff00000000 ,
                        0x0000000000000000000000000000000000000000208030ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff106810ff00000000 ,
                        0x0000000000000000000000000000000000000000308030ffffffffff205810ff ,
                        0x205810ff205810ff90e090ff309840ff308030ffffffffff106810ff00000000 ,
                        0x0000000000000000000000000000000000000000308830ffffffffffa0d8a0ff ,
                        0x205810ffa0e0a0ff30a060ff40a870ff90e080ffffffffff107010ff00000000 ,
                        0x0000000000000000000000000000000000000000308830ffffffffffc0e8c0ff ,
                        0x90c090ff205810ff40b080ffa0e090ff90e090ffffffffff107020ff00000000 ,
                        0x0000000000000000000000000000000000000000308830fff0f8f0ffd0f0d0ff ,
                        0xb0e0b0ff40a870ff40a870ffa0e0a0ffa0e0a0ffffffffff107020ff00000000 ,
                        0x0000000000000000000000000000000000000000308830ffe0f0e0ffe0f8e0ff ,
                        0x309040ff409860ff205810ff205810ffa0d8a0ffffffffff207820ff00000000 ,
                        0x0000000000000000000000000000000000000000308830d0c0e0c0ff308020ff ,
                        0x308030ffd0f0d0ffb0d0b0ff205810ff608050ffffffffff207820ff00000000 ,
                        0x00000000000000000000000000000000000000003088308080c090ffc0e0c0ff ,
                        0xe0f0e0fff0f8f0ffffffffffffffffffffffffffffffffff207820ff00000000 ,
                        0x00000000000000000000000000000000000000000000000030883080308830d0 ,
                        0x308830ff308830ff308830ff308830ff308030ff208030ff208030ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =10440
                    LayoutCachedTop =2160
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =2520
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =14136213
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =14136213
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =40.0
                    PressedColor =9592887
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeColor =4210752
                    HoverForeThemeColorIndex =0
                    HoverForeTint =75.0
                    PressedForeColor =4210752
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =11100
                    Top =2160
                    Width =420
                    FontSize =11
                    FontWeight =400
                    TabIndex =5
                    ForeColor =4210752
                    Name ="btnCloseup"
                    Caption ="Export to Excel"
                    StatusBarText ="Open the selected query in a new window"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Open the selected query in a new window"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000030506050204860ff303040ff ,
                        0x3040505000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000030506050305060ff4088a0ff3090b0ff ,
                        0x304050ff50809040000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000030506050305060ff3088b0ff40b8e0ff90e0f0ff ,
                        0x90d0e0ff6090a0ff0000000000000000000000006068602060606070606060b0 ,
                        0x606060f0505850f0505850b0405050c03080a0ff40b8e0ff90e0f0ff90e0f0ff ,
                        0x6098b0ff6088a050000000000000000070707040707060b0a09090ffc0b0a0ff ,
                        0xe0c8b0ffe0c8b0ffc0a8a0ff908080ff505850b080c0d0ffa0d8f0ff6098b0ff ,
                        0x6088a050000000000000000080787020707870b0b0b0b0ffffe8e0ffffe0d0ff ,
                        0xf0e0d0fff0d8c0fff0d0c0fff0d0b0ffb09890ff505850c0608890ff6088a050 ,
                        0x00000000000000000000000080787070b0a8a0fffff8f0fffff0e0ffffe8e0ff ,
                        0xffe8e0ffffe0d0fff0d8d0fff0d8c0fff0d0c0ff908880ff50606080f0d8c000 ,
                        0x000000000000000000000000808070b0d0d0d0fffff8fffffff8f0fffff0f0ff ,
                        0xfff0e0ffffe8e0fff0e0d0fff0e0d0fff0d8c0ffc0b0a0ff605860c0f0d8c000 ,
                        0x000000000000000000000000808080f0f0f0f0fffffffffffff8fffffff8f0ff ,
                        0xfff0f0fffff0e0ffffe8e0fff0e0d0fff0d8d0ffe0c8b0ff606060f0f0d8c000 ,
                        0x000000000000000000000000908080f0f0f0f0fffffffffffffffffffff8ffff ,
                        0xfff8f0fffff0f0ffffe8e0ffffe8e0ffffe0d0fff0d0c0ff606060f000000000 ,
                        0x000000000000000000000000908880b0e0d8d0ffffffffffffffffffffffffff ,
                        0xfff8f0fffff8f0fffff0f0fffff0e0ffffe0d0ffd0b8b0ff606860c000000000 ,
                        0x00000000000000000000000090888070b0b0b0ffffffffffffffffffffffffff ,
                        0xfffffffffff8f0fffff8f0fffff0f0ffffe8e0ffa09890ff6068607000000000 ,
                        0x00000000000000000000000090888020908880b0c0c0c0ffffffffffffffffff ,
                        0xfffffffffffffffffff8fffffff8f0ffc0b8b0ff707070b07070702000000000 ,
                        0x0000000000000000000000000000000090888040908880b0b0b0b0ffe0d8d0ff ,
                        0xf0f0f0fff0f8f0ffd0d8d0ffb0a8a0ff807870b0807870400000000000000000 ,
                        0x00000000000000000000000000000000000000009088802090888070908880b0 ,
                        0x908880f0808080f0808080b08080707080787020000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =11100
                    LayoutCachedTop =2160
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =2520
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =14136213
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =14136213
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =40.0
                    PressedColor =9592887
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeColor =4210752
                    HoverForeThemeColorIndex =0
                    HoverForeTint =75.0
                    PressedForeColor =4210752
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =12300
                    Top =2280
                    Width =1380
                    FontSize =11
                    FontWeight =400
                    TabIndex =6
                    ForeColor =4210752
                    Name ="btnRequery"
                    Caption =" Requery"
                    StatusBarText ="Requery the results set for the selected query"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Requery the results set for the selected query"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xb0886090805830f0804800ff905800ff804800ff805820d00000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b0803000a0805090 ,
                        0x906020ffc07820ffa05800ffb08040c000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000a0783010a06820ff ,
                        0xe09830ffb06000ffa07030c0a070301000000000804800000000000000000000 ,
                        0x0000000000000000000000000000000000000000e09030ffd08830ffe09840ff ,
                        0xf0a840ffa05810ff704000ff603000ff0000000000000000503000ff00000000 ,
                        0x000000000000000000000000000000000000000000000000e09830ffffc880ff ,
                        0xf0b050ffd08030ff905810ff0000000000000000805020ffb06810ff804000ff ,
                        0x000000000000000020a0f0ff4078e0ff4078e0ff4078e0ff4078e0ffe09830ff ,
                        0xffc070ffa06820ff4070e0ff4070e0ff906830fff0a850ffe09020ffb06000ff ,
                        0x704000ff0000000020a0f0ff50c0ffff50c0ffff4078e0ff50c0ffff50c0ffff ,
                        0xc08030ff50c0ffff50c0ffffd08840ffe09030ffffb860ffffb040ffc07810ff ,
                        0x904800ff704000ff20a0f0ff70d8ffff50c8ffff3080e0ff70d8ffffc0b0a0ff ,
                        0xc0b0a0ffc0b0a0ffc0b8b0ffd0c8c0ffd0b8a0ffd07820ffffb050ffa06010ff ,
                        0xa08050700000000020a8f0ff80d8ffff60d0ffff3080e0ff80d8ffff60d0ffff ,
                        0xa098a0ff90e0ffffd0a880ffc09060ffc07830ffffb050ffb06820ffb08050a0 ,
                        0x000000000000000020a8f0ff80e0ffff70d0f0ff3080f0ff80e0ffff70d0f0ff ,
                        0xa098a0ffd08040ffd07030ffd07030ffd07830ffd08040ffd08860ff00000000 ,
                        0x000000000000000020a8f0ff90e8ffff70d8f0ff3080f0ff90e8ffff70d8f0ff ,
                        0xa098a0ff90e8ffffe0c8b0fff0d0d0fff0c8b0ffffd8c0ffd0a080ff00000000 ,
                        0x000000000000000020a8f0f090e8fff070d8f0f03080f0f090e8fff070d8f0f0 ,
                        0x3080f0f090e8fff070d8f0f03080f0f090e8fff070d8f0f04078e0f000000000 ,
                        0x000000000000000020a8f0ffb0f0ffffb0f0ffff3088f0ffb0f0ffffb0f0ffff ,
                        0x3088f0ffb0f0ffffb0f0ffff3088f0ffb0f0ffffb0f0ffff4078e0ff00000000 ,
                        0x000000000000000020a8f0ff3090f0ff3088f0ff3088f0ff3088f0ff3088f0ff ,
                        0x3080f0ff3080f0ff3080f0ff3080f0ff4080e0ff4078e0ff4078e0ff00000000 ,
                        0x000000000000000020a8ffff80e8ffff80e8ffff80e8ffff80e8ffff70e0ffff ,
                        0x70d8ffff70d0ffff60c8ffff60c0ffff50b8ffff50b0ffff4078e0ff00000000 ,
                        0x000000000000000030a8ffff30a8ffff30a8ffff30a8ffff30a8ffff30a0ffff ,
                        0x30a0ffff30a0ffff3098f0ff3098f0ff3090f0ff3090f0ff3090f0ff00000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12300
                    LayoutCachedTop =2280
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =2640
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =14136213
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =14136213
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =40.0
                    PressedColor =9592887
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeColor =4210752
                    HoverForeThemeColorIndex =0
                    HoverForeTint =75.0
                    PressedForeColor =4210752
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =360
            BackColor =13025979
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
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
' Form:         QATool
' Level:        Framework form
' Version:      1.00
' Basis:        John Boetsch's frm_QA_Tool form
'
' Description:  Data quality review/validation related properties, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, December 15, 2016
' References:   -
' Revisions:    BLC - 12/15/2016 - 1.00 - initial version migrated from John's form,
'                                         converted to popup (Y)
' =================================

' ==================================================================
'                       ORIGINAL FORM HISTORY
' ==================================================================
' FORM NAME:    frm_QA_Tool
' Description:  Standard form for data quality review and validation
' Data source:  tbl_QA_Results
' Data access:  edit only, no deletions; opens to allow additions until a query is
'               selected, at which time additions are disallowed (see code in the subform)
' Pages:        pgResults, pgQueryViews, pgDataTables
' Functions:    fxnUpdateQAResults, fxnFilterRecords, fxnSetQueryFlag
' References:   fxnChangeDelimiter, fxnSaveFile, fxnSwitchboardIsOpen, fxnTableExists
' Source/date:  John R. Boetsch, Jan 2006
' Adapted/date: Bonnie L. Campbell, June 3, 2014
' Revisions:    JRB, May 16, 2006 - updated to use a subform for results, added conditional
'                   formatting and sort capability, and improved documentation
'               JRB, June 20, 2006 - added a button on pgResults to open the selected record
'                   in the data entry forms to maximize quality control during record fixes
'               JRB, 8/2/2006 - added additional error trapping to cmdOpenRecord
'               JRB, 10/5/2006 - fixed a problem with the refresh button giving a copy/save
'                   error message by saving the current record and turning off the form filter;
'                   added timeframe to fxnUpdateQAResults, and updated to save record before
'                   running the qa report
'               JRB, 11/14/2007 - revised the description and code in fxnUpdateQAResults
'               JRB, 12/17/2007 - added cbxTable_Enter to restore table pick list functionality
'                   regardless of back end in Access or SQL Server; added PageTabs change
'                   code to update and bookmark the last-selected subform record upon
'                   moving back to the first page; added code to handle multiple possible
'                   data time frames by adding an unbound ctl and linking the subform to this;
'                   added code to the results set report to also filter on data time frame;
'                   also added code to allow the user to flag records using the Is_done field
'               JRB, May 2008 - updated documentation
'               JRB, 6/18/2008 - updated Form_Open to check switchboard and enable/disable
'                   functionality based on application mode
'               JRB, 7/1/2008 - updated by adding blnRunQueries; added filter capability for
'                   Is_done and query type; added fxnFilterRecords
'               JRB, 9/17/2008 - added ref to frm_Progress_Meter (progress meter popup) in
'                   fxnUpdateQAResults
'               JRB, 9/19/2008 - added optgScope; changed txtTime_frame to cmbTimeframe;
'                   updated fxnUpdateQAResults to reflect both changes; updated call to
'                   rpt_QA_Results
'               JRB, 11/21/2008 - added txtEditQuery and fxnSetQueryFlag; updated to lock
'                   subQueryResults except when the query is named in a way that indicates
'                   its results are editable; updated cmdOpenRecord; updated cmdViewReport;
'                   added error traps to selObject and cmdDesignView; fixed a bug with opening
'                   the report and changing the filter values
'               JRB, 1/13/2009 - added save record to PageTabs_Change (copy/edit error)
'               JRB, 2/23/2009 - added cmdOpenBrowser; fixed a bug in selObject_AfterUpdate and
'                   updated fxnUpdateQAResults
'               JRB, 3/27/2009 - added cmdExport to allow quick results export to Excel
'               JRB, 5/1/2009 - updated cmdOpenBrowser to turn browser filters off by default;
'                   updated cmdExport_Click to default to current application path
'               JRB, 5/22/2009 - updated fxnFilterRecords
'               JRB, 6/10/2009 - updated cmdViewReport, cmdExport, fxnUpdateQAResults
'               JRB, 7/9/2009 - updated cbxTable to rely on tsys_Link_Tables, if present
'               JRB, 11/3/2009 - added cmdAutoFix and fxnEnableAutoFix
'               JRB, 2/8/2010 - updated fxnSetQueryFlag
'               JRB, 6/6/2011 - fixed a minor glitch by adding a call to fxnFilterRecords
'                   within PageTabs_Change (the front page filters were being ignored)
'               JRB, 1/31/2013 - resized panes, added cmdCloseup; set to use login rather than
'                   rely on session default user for Form_Dirty
'               --------------------------------------------------------------------------------------
'               BLC, 6/3/2014 - Adapted for NCPN WQ Utilities tool
'               BLC, 6/16/2014 - Updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 6/16/2014 - Modified to pull queries from tsys_Db_Templates
'               BLC, 8/22/2014 - Shifted blnRunQueries to mod_User & extended to project scope
'                    since used in setUserAccess (Dim -> Public), shifted fxnUpdateQAResults to mod_QA &
'                    renamed UpdateQAResults
' ==================================================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let Title(Value As String)
    If Len(Value) > 0 Then
        m_Title = Value

        'set the form title & caption
        Me.lblTitle.Caption = m_Title
        Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(Value)
    End If
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Directions(Value As String)
    If Len(Value) > 0 Then
        m_Directions = Value

        'set the form directions
        Me.lblDirections.Caption = m_Directions
    Else
        RaiseEvent InvalidDirections(Value)
    End If
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let CallingForm(Value As String)
    If Len(Value) > 0 Then
        m_CallingForm = Value
    Else
        RaiseEvent InvalidCallingForm(Value)
    End If
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' SUB:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 6/16/2014 - Updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
'               BLC, 8/22/2014 - Shifted user access level dictated field settings to setUserAccess
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool, added form toggle to minimize calling form
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    'default
    Me.CallingForm = "DbAdmin"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'restore calling form
    ToggleForm Me.CallingForm, -1

    Title = "QA Tool"
    Directions = "Choose the desired data timeframe && scope."
    lblDirections.ForeColor = lngLtBlue
'    btnComment.Caption = StringFromCodepoint(uComment)
'    btnComment.ForeColor = lngBlue
    
    'set hovers
    btnClose.HoverColor = lngGreen
    btnRefresh.HoverColor = lngGreen
    btnViewSummary.HoverColor = lngGreen
    btnDesignView.HoverColor = lngGreen
    btnAutoFix.HoverColor = lngGreen
    btnOpenBrowser.HoverColor = lngGreen
    btnExport.HoverColor = lngGreen
    btnCloseup.HoverColor = lngGreen
    btnRequery.HoverColor = lngGreen
    
    ' Close the form if the switchboard is not open
    If SwitchboardIsOpen = False Then
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this form to function properly.", , "Cannot open the form ..."
        DoCmd.CancelEvent
        GoTo Exit_Handler
    End If

    'set default app mode & initialize controls
    setUserAccess Me
    
    ' Initialize UI
    With Me
        ' Set form time frame to global time frame
        .cbxTimeframe = TempVars.item("Timeframe")
        
        .cbxDoneFilter = "False"
        .tglFilterByDone = True
    End With
    FilterRecords

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub Form_Load()
    On Error GoTo Err_Handler

    ' Requery the results subform to reflect updates if the user chose to run upon opening
    If blnRunQueries Then Me.subResults.Requery
    ' Turn off the form filter and move to a blank record so that no query record is visible
    Me.Filter = ""
    DoCmd.GoToRecord , , acNewRec

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 2105 'form was saved as not allowing new records
        MsgBox "Error #" & Err.Number & ": " & Err.Description & "" _
            & vbCrLf & "Form saved in wrong mode (QA Tool Load Error)" _
            & vbCrLf & "The form has been saved in a manner that does not permit new" _
            & vbCrLf & "records to be added. Contact the database administrator.", _
            , vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[QATool form])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxTimeframe_AfterUpdate
' Description:  actions after updating timeframe
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 6/16/2014 - Updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub cbxTimeframe_AfterUpdate()
    On Error GoTo Err_Handler

    If Me.cbxTimeframe <> TempVars.item("Timeframe") Then
        Me.btnRefresh.Enabled = False
        Me.optgMode.Enabled = False
    Else
        Select Case TempVars.item("UserAccessLevel")
          Case "admin", "power user"
            Me.btnRefresh.Enabled = True
            Me.optgMode.Enabled = True
          Case "data entry"
            Me.btnRefresh.Enabled = True
            Me.optgMode.Enabled = False
          Case Else
            ' leave them as is
        End Select
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTimeframe_AfterUpdate[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          optgMode_AfterUpdate
' Description:  actions after updating mode (view/edit)
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub optgMode_AfterUpdate()
    On Error GoTo Err_Handler

    ' Change the subform data mode depending on the user choice
    If Me.optgMode = 0 Then
    ' View mode
        Me.subQueryResults.Locked = True
        Me.tbxUser.Locked = True
        Me.tbxQueryDesc.Locked = True
        Me.tbxRemedy.Locked = True
        Me.subDataTables.Locked = True
        Me.Detail.BackColor = 13025979 ' steel blue (default)
    Else
    ' Edit mode
        ' Unlock the subform if an editable query
        If Me.tbxEditQuery = "OK" Then Me.subQueryResults.Locked = False _
            Else Me.subQueryResults.Locked = True
        Me.tbxUser.Locked = False
        Me.tbxQueryDesc.Locked = False
        Me.tbxRemedy.Locked = False
        Me.subDataTables.Locked = False
        Me.Detail.BackColor = 12574431 ' haystack
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - optgMode_AfterUpdate[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnClose_Click
' Description:  actions when close button is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub btnClose_Click()
    On Error GoTo Err_Handler

    DoCmd.Close , , acSaveNo

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Dirty
' Description:  actions when form data was modified
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub Form_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Note: this event is ignored on inserting a new record if BeforeInsert code exists

    ' Bail out if the refresh button is disabled (app mode or if selected timeframe <>
    '   db timeframe)
    If Me.btnRefresh.Enabled = False Then GoTo Exit_Handler

    ' Bail out if no object record is selected - keeps from adding bogus new records
    If IsNull(Me.cbxObject) Then
        DoCmd.CancelEvent
        GoTo Exit_Handler
    End If

    ' Once a user starts to make edits in the record, update the user field
    '   on the results summary page
    If SwitchboardIsOpen Then Me.tbxUser = Environ("Username")
    Me.tbxRemedy_date = Now()

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Dirty[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Close
' Description:  form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 15, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/15/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore calling form
    ToggleForm Me.CallingForm, 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   QATool Tab Functionality
' =================================
' ---------------------------------
' SUB:          PageTabs_Change
' Description:  tab change actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 8/22/2014 - updated UpdateQAResults function name
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub PageTabs_Change()
    On Error GoTo Err_Handler

    Dim rs As DAO.Recordset
    Dim strCriteria As String
    Dim varReturn As Variant

    ' Bail out if the refresh button is disabled (app mode or if selected timeframe <>
    '   db timeframe)
    If Me.btnRefresh.Enabled = False Then GoTo Exit_Handler

    ' If moving to the first page, and if a specific query record has been selected
    '   move the subform bookmark to the currently-selected record
    If Me.PageTabs = 0 And IsNull(Me.cbxObject) = False Then
        ' Save the current record, reset the form filter and query selector, reset the form
        '   to allow additions, and move to a blank record
        If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

        ' Run the function to update the current QA query record
        varReturn = UpdateQAResults(False, Me.cbxObject)
        Me.Requery
        FilterRecords
        strCriteria = "[Query_name] = """ & Me.cbxObject.Value & _
            """ AND [Time_frame] = """ & Me.cbxTimeframe & _
            """ AND [Data_scope] = " & Me.optgScope

        Set rs = Me.subResults.Form.RecordsetClone
        rs.FindFirst strCriteria
        If rs.NoMatch Then
            'MsgBox "No entry found.", vbInformation
        Else
            Me.subResults.Form.Bookmark = rs.Bookmark
        End If
    ElseIf Me.PageTabs = 1 And IsNull(Me.cbxObject) = False Then
        ' Call the function to update the query flag
        SetQueryFlag
        EnableAutoFix
    End If

Exit_Handler:
    On Error Resume Next
    'cleanup
    rs.Close
    Set rs = Nothing
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PageTabs_Change[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
' ---------  Tab Pages ------------
' =================================

' =================================
' PAGE NAME:    QA Results Summary Page (tabResults)
' Description:  shows an overview of validation query results
' Unbound ctls: none
' Subforms:     subResults - subform for showing the results summaries
' =================================

' ---------------------------------
' SUB:          cbxTypeFilter_AfterUpdate
' Description:  actions after updating type filter
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub cbxTypeFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.tglFilterByType = Not IsNull(Me.cbxTypeFilter)
    FilterRecords

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTypeFilter_AfterUpdate[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglFilterByType_AfterUpdate
' Description:  actions after updating filter by type toggle
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub tglFilterByType_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cbxTypeFilter) = False Then FilterRecords Else Me.tglFilterByType = False

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByType_AfterUpdate[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxDoneFilter_AfterUpdate
' Description:  actions after updating done filter
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub cbxDoneFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.tglFilterByDone = Not IsNull(Me.cbxDoneFilter)
    FilterRecords

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxDoneFilter_AfterUpdate[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglFilterByDone_AfterUpdate
' Description:  actions after updating done filter by type toggle
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub tglFilterByDone_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cbxDoneFilter) = False Then FilterRecords Else Me.tglFilterByDone = False

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByDone_AfterUpdate[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnRefresh_Click
' Description:  actions when refresh button is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 6/13/2014 - initial version
'               BLC, 8/25/2014 - updated UpdateQAResults function name (dropped fxn prefix)
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub btnRefresh_Click()
    On Error GoTo Err_Handler

    ' Save the current record, reset the form filter and query selector, reset the form
    '   to allow additions, and move to a blank record
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    Me.Filter = ""
    Me.FilterOn = False
    Me.cbxObject = Null
    Me.subQueryResults.SourceObject = ""
    Me.AllowAdditions = True
    DoCmd.GoToRecord , , acNewRec

    ' Set the form to view mode and call the event procedure for the form mode ctl
    Me.optgMode = 0
    optgMode_AfterUpdate
    Me.Repaint

    ' Refresh the validation query results (filtering requeries the subform)
    UpdateQAResults
    FilterRecords

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRefresh_Click[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnViewSummary_Click
' Description:  actions when refresh button is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub btnViewSummary_Click()
    On Error GoTo Err_Handler

    ' Generate the QA report
    Dim strRptName As String
    Dim strMsg As String
    Dim strFilter As String
    Dim strTimeframe As String
    Dim strScope As String
    Dim strInitFile As String
    Dim strSaveFile As String
    Dim varResponse As VbMsgBoxResult

    strRptName = "rpt_QA_Results"

    strMsg = "This will open the quality assurance report ..." & vbCrLf & vbCrLf & _
        "Would you like to limit report results to " & Me.cbxTimeframe & "?"
    varResponse = MsgBox(strMsg, vbYesNoCancel, "Quality assurance report")

    Select Case varResponse
      Case vbCancel
        GoTo Exit_Handler
      Case vbYes
        strTimeframe = Me.cbxTimeframe
        strFilter = "[Time_frame]=""" & strTimeframe & """"
      Case Else
        strTimeframe = Trim(InputBox("Enter the time frame to filter by" & vbCrLf & _
            "(or leave blank to show all):", "Filter by data time frame", _
            Me.cbxTimeframe))
        If strTimeframe <> "" Then
            strFilter = "[Time_frame]=""" & strTimeframe & """"
        Else
            strFilter = ""
        End If
    End Select

    ' Save the current record so that all changes are reflected in the report
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    Select Case Me.optgScope
      Case 0
        strScope = "Uncertifed event data only"
      Case 1
        strScope = "Both certified and uncertified events"
      Case 2
        strScope = "Certified event data only"
    End Select

    If MsgBox("Would you like to filter by the current data scope?" & _
        vbCrLf & vbCrLf & "   " & strScope, vbYesNo, "Filter by data scope?") = vbYes Then
        If strFilter <> "" Then strFilter = strFilter & " AND "
        strFilter = strFilter & "[Data_scope]=" & Me.optgScope
    End If

    ' Open the formatted report output, filtering on time frame
    DoCmd.OpenReport "rpt_QA_Results", acViewPreview, , strFilter
    If MsgBox("Would you like to save this report?", vbYesNo + vbDefaultButton2, _
        "Save report to a file?") = vbYes Then
        If strTimeframe <> "" Then
            ' Add timeframe to file name
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                strTimeframe & "_" & CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".snp"
        Else
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".snp"
        End If
        ' Open the save file dialog and update to the actual name given by the user
        strSaveFile = SaveFile(strInitFile, "Snapshot Viewer (*.snp)", "*.snp")
        DoCmd.OutputTo acOutputReport, strRptName, acFormatSNP, strSaveFile, True
        MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnViewSummary_Click[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
' PAGE NAME:    View/Fix Query Results (tabQueries)
' Description:  shows records returned by individual QA queries, provides the
'               user the opportunity to fix these
' Unbound ctls: cbxObject - combo box for selecting the query object by name
' Subforms:     subQueryResults - subform showing results of the selected query
' =================================

' ---------------------------------
' SUB:          cbxObject_AfterUpdate
' Description:  actions after updating object combobox
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 6/13/2014 - initial version
'               BLC, 8/22/2014 - updated UpdateQAResults function name
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub cbxObject_AfterUpdate()
    On Error GoTo Err_Handler

    Dim strCriteria As String
    Dim varReturn As Variant

    ' Bail out if the refresh button is disabled (app mode or if selected timeframe <>
    '   db timeframe)
    If Me.btnRefresh.Enabled = False Then GoTo Exit_Handler

    ' Exit if no query selected
    If IsNull(Me.cbxObject) Then
        MsgBox "Please pick from the list", vbOKOnly, "No Query Selected"
        Me.AllowAdditions = True
        DoCmd.GoToRecord , , acNewRec
        Me.tbxEditQuery = ""
        Me.tbxEditQuery.ForeColor = 0          'black
        Me.tbxEditQuery.BackColor = 8454143    'yellow
        GoTo Exit_Handler
    End If
    
    ' Bind the subform to the selected query
    Me.subQueryResults.SourceObject = "Query." & Me.cbxObject.Value
    ' Build the filter string and see if a record already exists
    strCriteria = "[Query_name] = """ & Me.cbxObject.Value & _
        """ AND [Time_frame] = """ & Me.cbxTimeframe & _
        """ AND [Data_scope] = " & Me.optgScope
    If DCount("*", "tbl_QA_Results", strCriteria) = 0 Then
        ' Run the function to update the current QA query record
        varReturn = UpdateQAResults(False, Me.cbxObject, True)
    End If
    ' Set the form to the selected record
    Me.Form.Filter = strCriteria
    Me.Form.FilterOn = True

    ' Call the function to update the query flag
    SetQueryFlag
    EnableAutoFix

    Dim qdf As DAO.QueryDef
    Dim qdfs As DAO.QueryDefs
    Set qdfs = DBEngine(0)(0).QueryDefs

    On Error Resume Next
    For Each qdf In qdfs
        If qdf.Name = Me.cbxObject.Value Then
            MsgBox ("This query returns (" & DCount("*", qdf.Name) & _
                ") records that meet the following criteria: " & _
                vbCrLf & vbCrLf & qdf.Properties("Description"))
        End If
    Next qdf

Exit_Handler:
    On Error Resume Next
    Set qdfs = Nothing
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "Error #" & Err.Number & ": " & Err.Description & _
            vbCrLf & "This query is no longer available in the application." & _
            vbCrLf & """" & Me.cbxObject & """", , "Query not found" _
            , vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxObject_AfterUpdate[QATool form])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxObject_AfterUpdate[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDesignView_Click
' Description:  design view button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 6/13/2014 - initial version
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub btnDesignView_Click()
    On Error GoTo Err_Handler

    ' Open the selected query in design view after checking that a query is selected
    If IsNull(Me.cbxObject) = False Then _
        DoCmd.OpenQuery Me.cbxObject.Value, acViewDesign, acReadOnly

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDesignView_Click[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnAutoFix_Click
' Description:  auto fix button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 6/13/2014 - initial version
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub btnAutoFix_Click()
    On Error GoTo Err_Handler

    Dim ctlAutoFix As Control
    Dim varAutoFix As Variant

    varAutoFix = Null

    On Error Resume Next
    Set ctlAutoFix = Forms!frm_QA_Tool.subQueryResults!varAutoFix
    varAutoFix = ctlAutoFix.Value
    On Error GoTo Err_Handler

    If IsNull(varAutoFix) Then
        MsgBox "There are no records selected, or no query is specified to fix the results."
    ElseIf Left(varAutoFix, 1) = "t" Then
    ' Object is a table - open in the next tab
        MsgBox "Object is not labeled as a query:" & vbCrLf & vbCrLf & _
            "  " & varAutoFix, , "No action taken"
    ElseIf Left(varAutoFix, 1) = "q" Then
    ' Object is a query - open on its own
        Dim qdf As DAO.QueryDef
        Dim qdfs As DAO.QueryDefs
        Set qdfs = DBEngine(0)(0).QueryDefs
        On Error Resume Next
        For Each qdf In qdfs
            If qdf.Name = varAutoFix Then
                If MsgBox("This will open/run the following query:" & vbCrLf & vbCrLf & _
                    """" & varAutoFix & """" & vbCrLf & vbCrLf & qdf.Properties("Description"), _
                    vbOKCancel, "Open or run query ...") = vbCancel Then
                    GoTo Exit_Handler
                End If
            End If
        Next qdf
        DoCmd.OpenQuery varAutoFix
        Me.subQueryResults.Requery
    End If

Exit_Handler:
    On Error Resume Next
    'cleanup
    Set ctlAutoFix = Nothing
    Set qdfs = Nothing
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 2427   ' No records in the subform
        ' Do nothing ...
      Case 2465   ' Needed field is not present in the record set
        MsgBox "No form is specified for fixing these results", , "Missing query field"
      Case 2467   ' No subform recordset
        MsgBox "No query result set"
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAutoFix_Click[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnOpenRecord_Click
' Description:  open record button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 6/13/2014 - initial version
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub btnOpenRecord_Click()
    On Error GoTo Err_Handler

    ' Opens the selected subform record in the object specified in the query
    '   to make use of quality control features of the front end during edits

    Dim ctlObject As Control
    Dim ctlFilter As Control
    Dim ctlArgs As Control
    Dim varObject As Variant
    Dim varFilter As Variant
    Dim varArgs As Variant

    varObject = Null
    varFilter = Null
    varArgs = Null
    
    On Error Resume Next
    Set ctlObject = Forms!frm_QA_Tool.subQueryResults!varObject
    varObject = ctlObject.Value
    Set ctlFilter = Forms!frm_QA_Tool.subQueryResults!varFilter
    varFilter = ctlFilter.Value
    Set ctlArgs = Forms!frm_QA_Tool.subQueryResults!varArgs
    varArgs = ctlArgs.Value
    On Error GoTo Err_Handler

    If IsNull(varObject) Then
        MsgBox "There are no records selected, or no form is specified."
    ElseIf Left(varObject, 1) = "t" Then
    ' Object is a table - open in the next tab
        Me.subDataTables.SourceObject = "Table." & varObject
        Me.cbxTable = varObject
        Me.tabDataTables.SetFocus
    ElseIf Left(varObject, 1) = "q" Then
    ' Object is a query - open on its own
        Dim qdf As DAO.QueryDef
        Dim qdfs As DAO.QueryDefs
        Set qdfs = DBEngine(0)(0).QueryDefs
        On Error Resume Next
        For Each qdf In qdfs
            If qdf.Name = varObject Then
                If MsgBox("This will open/run the following query:" & vbCrLf & vbCrLf & _
                    """" & varObject & """" & vbCrLf & vbCrLf & qdf.Properties("Description"), _
                    vbOKCancel, "Open or run query ...") = vbCancel Then
                    GoTo Exit_Handler
                End If
            End If
        Next qdf
        DoCmd.OpenQuery varObject
        Me.subQueryResults.Requery
    ElseIf IsNull(varFilter) Then
    ' Filter by form alone if no filter
        Select Case varObject
          Case "frm_Contacts"
            Set gvarRefContactCtl = Me.subQueryResults
          Case "fsub_Project_Taxa"
            Set gvarRefTaxonCtl = Me.subQueryResults
          Case Else
            Set gvarRefForm = Me.Form
            Set gvarRefCtl = Me.subQueryResults
        End Select
        DoCmd.OpenForm varObject, , , , , , varArgs
    Else
    ' Filter by form and filter
        Select Case varObject
          Case "frm_Contacts"
            Set gvarRefContactCtl = Me.subQueryResults
          Case "fsub_Project_Taxa"
            Set gvarRefTaxonCtl = Me.subQueryResults
          Case Else
            Set gvarRefForm = Me.Form
            Set gvarRefCtl = Me.subQueryResults
        End Select
        DoCmd.OpenForm varObject, , , varFilter, , , varArgs
    End If

Exit_Handler:
    On Error Resume Next
    'cleanup
    Set ctlArgs = Nothing
    Set ctlFilter = Nothing
    Set ctlObject = Nothing
    Set qdfs = Nothing
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 2427   ' No records in the subform
        ' Do nothing ...
      Case 2465   ' Needed field is not present in the record set
        MsgBox "No form is specified for fixing these results", , "Missing query field"
      Case 2467   ' No subform recordset
        MsgBox "No query result set"
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnOpenRecord_Click[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnOpenBrowser_Click
' Description:  open browser click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub btnOpenBrowser_Click()
    On Error GoTo Err_Handler

    Set gvarRefForm = Me.Form
    Set gvarRefCtl = Me.subQueryResults
    ' Open to a blank record - to distinguish from opening to the selected record in the subform
    DoCmd.OpenForm "frm_Data_Browser", , , , acFormAdd, , "off"

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnOpenBrowser_Click[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnExport_Click
' Description:  export button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub btnExport_Click()
    On Error GoTo Err_Handler

    Dim strQName As String
    Dim strSaveFile As String

    ' Bail out if no query is currently selected
    If IsNull(Me.cbxObject) Then GoTo Exit_Handler
    ' Requery the selected record in the recordset, and update the subform
    Me.subQueryResults.Requery
    strQName = Me.cbxObject
    strSaveFile = CurrentProject.Path & "\" & strQName & "_" & _
        CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".xls"
    DoCmd.OutputTo acOutputQuery, strQName, acFormatXLS, strSaveFile, True
    MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnExport_Click[QATool form])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------------------------
' SUB:          btnCloseup_Click
' Description:  Closeup button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub btnCloseup_Click()
    On Error GoTo Err_Handler

    ' Open the selected query in a new window after checking that a query is selected
    If IsNull(Me.cbxObject) = False Then
        If Me.tbxEditQuery = "OK" Then
            DoCmd.OpenQuery Me.cbxObject.Value, acViewNormal, acEdit
        Else
            DoCmd.OpenQuery Me.cbxObject.Value, acViewNormal, acReadOnly
        End If
        DoCmd.Maximize
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cbxObject & """", , "Object not found" & _
            vbCrLf & "Error encountered (#" & Err.Number & " - btnCloseup_Click[QATool form])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCloseup_Click[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnRequery_Click
' Description:  requery button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub btnRequery_Click()
    On Error GoTo Err_Handler

    'Dim varReturn As Variant

    ' Bail out if no query is currently selected
    If IsNull(Me.cbxObject) Then GoTo Exit_Handler
    ' Requery the selected record in the recordset, and update the subform
    Me.subQueryResults.Requery
    ' Run the function to update the current QA query record - commented out because this
    '   is done upon changing page tabs
    'varReturn = fxnUpdateQAResults(False, Me.cbxObject)

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRequery_Click[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxUser_Dirty
' Description:  user textbox after data entry actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub tbxUser_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Prompt user to confirm before allowing edits in the QA user control
    If MsgBox("Are you sure you want to change the user name?", _
        vbYesNo, "Please confirm ...") = vbNo Then
        DoCmd.CancelEvent
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxUser_Dirty[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxQueryDesc_Dirty
' Description:  query description after data entry actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub tbxQueryDesc_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Prompt user to confirm before allowing edits in query definition control
    If MsgBox("Are you sure you want to change the query definition?", _
        vbYesNo, "Please confirm ...") = vbNo Then
        DoCmd.CancelEvent
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxQueryDesc_Dirty[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
' PAGE NAME:    Browse Data Tables (tabDataTables)
' Description:  allows the user to select and view the contents of individual data
'               tables to make data revisions
' Unbound ctls: cbxTable - combo box for selecting the table object by name
' Subforms:     subDataTables - subform showing the contents of the selected table
' =================================

' ---------------------------------
' SUB:          cbxTable_AfterUpdate
' Description:  table combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub cbxTable_AfterUpdate()
    On Error GoTo Err_Handler

    ' Once a table is selected, bind the subform to this table
    If IsNull(Me.cbxTable) Then
    ' If none selected ...
        Me.subDataTables.SourceObject = ""
    Else
    ' If a table is selected ...
        If TableExists(Me.cbxTable) Then
            Me.subDataTables.SourceObject = "Table." & Me.cbxTable.Value
        Else
            MsgBox "Unable to find the selected table in the database ...", , _
                "Table not found"
        End If
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTable_AfterUpdate[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxTable_Enter
' Description:  table combobox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Sub cbxTable_Enter()
     On Error GoTo Err_Handler

    Dim strSysTable As String

    strSysTable = "tsys_Link_Tables"     ' System table listing linked tables

    ' If the system table does not exist, replace the row source with one that doesn't use it
    If TableExists(strSysTable) = False Then
        Me.cbxTable.RowSource = "SELECT MSysObjects.Name " & _
            "FROM MSysObjects " & _
            "WHERE (((MSysObjects.Name) Like 'tbl_*' " & _
            "And (MSysObjects.Name)<>'tbl_QA_Results')) " & _
            "OR (((MSysObjects.Name)='tlu_Project_Crew')) " & _
            "OR (((MSysObjects.Name)='tlu_Project_Taxa')) " & _
            "OR (((MSysObjects.Name)='tlu_Park_Taxa'));"
        Me.cbxTable.ColumnCount = 1
        Me.cbxTable.ListWidth = Me.cbxTable.Width
        Me.cbxTable.Requery
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTable_Enter[QATool form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     FilterRecords
' Description:  filter records by the indicated field
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 5, 2006
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               JRB, May 2008 - made code more robust and error-proof
'               JRB, 7/1/2008 - updated by filtering on the subform rather than the form
'               JRB, 5/22/2009 - updated filter AND clauses
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Function FilterRecords()
    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim bFilterOn As Boolean

    bFilterOn = False
    strFilter = ""

    ' Save the record (to trigger validation)
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    If Me.tglFilterByType Then
        bFilterOn = True
        strFilter = strFilter & "[Query_type] = """ & Me.cbxTypeFilter & """"
    End If
    If Me.tglFilterByDone Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Is_done] = " & Me.cbxDoneFilter & ""
    End If

    ' Apply the filter
    'Me.Filter = strFilter
    'Me.FilterOn = bFilterOn
    Me.subResults.Form.Filter = strFilter
    Me.subResults.Form.FilterOn = bFilterOn

    ' Make the labels bold or not depending on filter settings
    Me.lblTypeFilter.FontBold = Me.tglFilterByType
    Me.lblDoneFilter.FontBold = Me.tglFilterByDone

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 2001   ' Run time canceled event (validation error) - do nothing
        Me.tglFilterByType = False
        Me.tglFilterByDone = False
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FilterRecords[QATool form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SetQueryFlag
' Description:  Updates the flag to indicate whether or not the query results are editable
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, 10/7/2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               JRB, 2/8/2010 - updated flag from "X" to "_X" in of x as last letter in name
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Function SetQueryFlag()
    On Error GoTo Err_Handler

    ' Update the visual flag to indicate whether or not the query results are editable
    '   Note: suffix of "_X" means that the query results may be edited
    If Right(Me.cbxObject.Value, 2) = "_X" Then
        Me.tbxEditQuery = "OK"
        Me.tbxEditQuery.ForeColor = 16777215   'white
        Me.tbxEditQuery.BackColor = 4227072    'green
        ' Unlock the subform if in edit mode
        If Me.optgMode = 1 Then Me.subQueryResults.Locked = False _
            Else Me.subQueryResults.Locked = True
    Else
        Me.tbxEditQuery = "No"
        Me.tbxEditQuery.ForeColor = 16777215   'white
        Me.tbxEditQuery.BackColor = 255        'red
        ' Lock the subform
        Me.subQueryResults.Locked = True
    End If

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetQueryFlag[QATool form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     EnableAutoFix
' Description:  Enables or disables the control for running an action query to fix records
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, 11/3/2009
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool (initial)
'               Bonnie Campbell, December 15, 2016 - for NCPN tools
' Revisions:
'               BLC, 6/16/2014 - Updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 12/15/2016 - Adapted from NCPN WQ Utilities tool
' ---------------------------------
Private Function EnableAutoFix()
    On Error GoTo Err_Handler

    Dim ctlAutoFix As Control

    Me.btnAutoFix.Enabled = False

    ' The following looks for 'varAutoFix' field in the query results ...
    '   If it isn't there, it will throw a trapped error and the ctl will remain disabled
    Set ctlAutoFix = Forms!frm_QA_Tool.subQueryResults!varAutoFix

    ' If no error, the field is there ... enable the ctl if user has sufficient rights
    Select Case TempVars.item("UserAccessLevel")
      Case "admin", "power user"
        Me.btnAutoFix.Enabled = True
    End Select

Exit_Handler:
    On Error Resume Next
    'cleanup
    Set ctlAutoFix = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 2465, 2467
        ' Do nothing ...
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - EnableAutoFix[QATool form])"
    End Select
    Resume Exit_Handler
End Function


'' ---------------------------------
'' SUB:          btnOpenRecord_Click
'' Description:
'' Parameters:   -
'' Returns:      -
'' Throws:       -
'' References:   -
'' Source/date:  John R. Boetsch, May 2008
'' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
'' Revisions:    BLC, 6/13/2014 - XX
'' ---------------------------------
'Private Sub btnOpenRecord_Click()
'    On Error GoTo Err_Handler
'
'    ' Opens the selected subform record in the object specified in the query
'    '   to make use of quality control features of the front end during edits
'
'    Dim ctlObject As Control
'    Dim ctlFilter As Control
'    Dim ctlArgs As Control
'    Dim varObject As Variant
'    Dim varFilter As Variant
'    Dim varArgs As Variant
'
'    varObject = Null
'    varFilter = Null
'    varArgs = Null
'
'    On Error Resume Next
'    Set ctlObject = Forms!frm_QA_Tool.subQueryResults!varObject
'    varObject = ctlObject.Value
'    Set ctlFilter = Forms!frm_QA_Tool.subQueryResults!varFilter
'    varFilter = ctlFilter.Value
'    Set ctlArgs = Forms!frm_QA_Tool.subQueryResults!varArgs
'    varArgs = ctlArgs.Value
'    On Error GoTo Err_Handler
'
'    If IsNull(varObject) Then
'        MsgBox "There are no records selected, or no form is specified."
'    ElseIf Left(varObject, 1) = "t" Then
'    ' Object is a table - open in the next tab
'        Me.subDataTables.SourceObject = "Table." & varObject
'        Me.cbxTable = varObject
'        Me.pgDataTables.SetFocus
'    ElseIf Left(varObject, 1) = "q" Then
'    ' Object is a query - open on its own
'        Dim qdf As DAO.QueryDef
'        Dim qdfs As DAO.QueryDefs
'        Set qdfs = DBEngine(0)(0).QueryDefs
'        On Error Resume Next
'        For Each qdf In qdfs
'            If qdf.Name = varObject Then
'                If MsgBox("This will open/run the following query:" & vbCrLf & vbCrLf & _
'                    """" & varObject & """" & vbCrLf & vbCrLf & qdf.Properties("Description"), _
'                    vbOKCancel, "Open or run query ...") = vbCancel Then
'                    GoTo Exit_Handler
'                End If
'            End If
'        Next qdf
'        DoCmd.OpenQuery varObject
'        Me.subQueryResults.Requery
'    ElseIf IsNull(varFilter) Then
'    ' Filter by form alone if no filter
'        Select Case varObject
'          Case "frm_Contacts"
'            Set gvarRefContactCtl = Me.subQueryResults
'          Case "fsub_Project_Taxa"
'            Set gvarRefTaxonCtl = Me.subQueryResults
'          Case Else
'            Set gvarRefForm = Me.Form
'            Set gvarRefCtl = Me.subQueryResults
'        End Select
'        DoCmd.OpenForm varObject, , , , , , varArgs
'    Else
'    ' Filter by form and filter
'        Select Case varObject
'          Case "frm_Contacts"
'            Set gvarRefContactCtl = Me.subQueryResults
'          Case "fsub_Project_Taxa"
'            Set gvarRefTaxonCtl = Me.subQueryResults
'          Case Else
'            Set gvarRefForm = Me.Form
'            Set gvarRefCtl = Me.subQueryResults
'        End Select
'        DoCmd.OpenForm varObject, , , varFilter, , , varArgs
'    End If
'
'Exit_Handler:
'    On Error Resume Next
'    'cleanup
'    Set ctlArgs = Nothing
'    Set ctlFilter = Nothing
'    Set ctlObject = Nothing
'    Set qdfs = Nothing
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case 2427   ' No records in the subform
'        ' Do nothing ...
'      Case 2465   ' Needed field is not present in the record set
'        MsgBox "No form is specified for fixing these results", , "Missing query field"
'      Case 2467   ' No subform recordset
'        MsgBox "No query result set"
'      Case 3011, 7874   ' Object not found
'        MsgBox "The table, query or form is no longer available in the application.", , _
'            "Object not found"
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - btnOpenRecord_Click[QATool form])"
'    End Select
'    Resume Exit_Handler
'End Sub
