Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7860
    DatasheetFontHeight =11
    ItemSuffix =35
    Left =4065
    Top =3105
    Right =16815
    Bottom =14490
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x656d3fc92f05e540
    End
    Caption ="Events (Sampling Visits)"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =1620
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =60
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =420
                    Width =6840
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Choose the site location, then enter the sampling start date."
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1560
                    Top =1260
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblStartDate"
                    Caption ="Start Date"
                    GridlineColor =10921638
                    LayoutCachedLeft =1560
                    LayoutCachedTop =1260
                    LayoutCachedWidth =2805
                    LayoutCachedHeight =1575
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6780
                    Top =1200
                    Width =720
                    ForeColor =16711680
                    Name ="btnComment"
                    Caption ="comm"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6780
                    LayoutCachedTop =1200
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =1560
                    ForeThemeColorIndex =-1
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =3840
                    Top =60
                    Width =3900
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblContext"
                    Caption ="context"
                    GridlineColor =10921638
                    LayoutCachedLeft =3840
                    LayoutCachedTop =60
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =480
                    Top =840
                    Width =1065
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblLocation"
                    Caption ="Location"
                    GridlineColor =10921638
                    LayoutCachedLeft =480
                    LayoutCachedTop =840
                    LayoutCachedWidth =1545
                    LayoutCachedHeight =1155
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4860
                    Top =840
                    Width =1680
                    TabIndex =1
                    ForeColor =16711680
                    Name ="btnAddLocation"
                    Caption =" Add Location"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add a new location"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000009070603090706080 ,
                        0x906850c0806850f0806050f0805840c070584080705040300000000000000000 ,
                        0x0000000000000000000000000000000000000000a0786060907060e0907050ff ,
                        0x706840ff906840ffb06840ff906040ff805040ff705040e07050406000000000 ,
                        0x00000000000000000000000000000000a0806060a07860f0c08060ff807830ff ,
                        0x607820ff506040ff906040ffc06840ffb06040ff905840ff705040f070504060 ,
                        0x000000000000000000000000b0807030a07860e0d08860fff08860ffd08040ff ,
                        0x608820ff308820ff606830ff706840ffb06840ffb06040ff905840ff705040e0 ,
                        0x705040300000000000000000b0807080c08870fff09860fff09060ffc08860ff ,
                        0x508040ff409820ff408830ff707040ffb06840ff806040ff606040ff605840ff ,
                        0x705840800000000000000000b08070c0e09870ffffa070ffe09070ff707850ff ,
                        0x50a840ff50a830ff809830ffd07850ff707040ff506840ff307810ff606830ff ,
                        0x805840c00000000000000000b08870f0f0b890fff0a880ff607860ff609850ff ,
                        0x60b040ff90a040ffe08860ffd08050ff507050ff408830ff309010ff408020ff ,
                        0x806050f00000000000000000b08870f0ffb8a0ffa09880ff40a850ff60b850ff ,
                        0x909860ffe09870ffe09060ffe08860ff60a030ff40a020ff409820ff408820ff ,
                        0x806850f00000000000000000b08870c0e0b8a0ff90b880ff50c870ff50c870ff ,
                        0x80a070ffc0a070ffd09870ffe09060ffa09840ff50a830ff40a020ff508830ff ,
                        0x906850c00000000000000000b0907080b0a080ff90d0a0ff70e090ff90e0a0ff ,
                        0x80e8a0ffa0d090ffb0a870fff09870ffe09060ff50a830ff50a830ff708840ff ,
                        0x907060800000000000000000b0907030b09080e090c8a0ffa0e8b0ffd0f0c0ff ,
                        0xd0f8d0fff0f8d0ffb0e8b0ff809860ffd09840ffb09040ff609840ff907050e0 ,
                        0x90706030000000000000000000000000b0908070b0a080f0b0d8b0ffd0f0d0ff ,
                        0xe0f8d0ff80b070ff80d890ff50b070ff908840ffd08860ffa07860f0a0786060 ,
                        0x0000000000000000000000000000000000000000b0908070b09080e0c0c0a0ff ,
                        0xd0e0c0ffb0e8b0ff80e0a0ff60b870ffa08850ffa07860e0a080606000000000 ,
                        0x000000000000000000000000000000000000000000000000b0908030b0908080 ,
                        0xb09080c0b09080f0c09880f0b09080c0b0807080b08070300000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4860
                    LayoutCachedTop =840
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1200
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =-1
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =1560
                    Top =840
                    Width =3120
                    Height =315
                    ColumnOrder =0
                    FontSize =8
                    TabIndex =2
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x0100000070000000020000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff00000000000200000004000000070000000100000000000000 ,
                        0xfff2000000000000000000000000000000000000000000000000000000000000 ,
                        0x22002200000000002200220000000000
                    End
                    Name ="cbxLocation"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1440;0;0"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    ControlTipText ="Select location of this sampling visit"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1560
                    LayoutCachedTop =840
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =1155
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000000000000200000001 ,
                        0x00000000000000fff20000020000002200220000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5445
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6660
                    Top =60
                    Width =720
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnSave"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save Record"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000d0687050c06860ffb05850ffa05050ffa05050ff ,
                        0xa05050ff904850ff904840ff904840ff804040ff803840ff803840ff703840ff ,
                        0x703830ff0000000000000000d06870fff09090ffe08080ffb04820ff403020ff ,
                        0xc0b8b0ffc0b8b0ffd0c0c0ffd0c8c0ff505050ffa04030ffa04030ffa03830ff ,
                        0x703840ff0000000000000000d07070ffff98a0fff08880ffe08080ff705850ff ,
                        0x404030ff907870fff0e0e0fff0e8e0ff908070ffa04030ffa04040ffa04030ff ,
                        0x803840ff0000000000000000d07870ffffa0a0fff09090fff08880ff705850ff ,
                        0x000000ff404030fff0d8d0fff0e0d0ff807860ffb04840ffb04840ffa04040ff ,
                        0x804040ff0000000000000000d07880ffffa8b0ffffa0a0fff09090ff705850ff ,
                        0x705850ff705850ff705850ff706050ff806860ffc05850ffb05050ffb04840ff ,
                        0x804040ff0000000000000000e08080ffffb0b0ffffb0b0ffffa0a0fff09090ff ,
                        0xf08880ffe08080ffe07880ffd07070ffd06870ffc06060ffc05850ffb05050ff ,
                        0x904840ff0000000000000000e08890ffffb8c0ffffb8b0ffd06060ffc06050ff ,
                        0xc05850ffc05040ffb05030ffb04830ffa04020ffa03810ffc06060ffc05850ff ,
                        0x904840ff0000000000000000e09090ffffc0c0ffd06860ffffffffffffffffff ,
                        0xfff8f0fff0f0f0fff0e8e0fff0d8d0ffe0d0c0ffe0c8c0ffa03810ffc06060ff ,
                        0x904850ff0000000000000000e098a0ffffc0c0ffd07070ffffffffffffffffff ,
                        0xfffffffffff8f0fff0f0f0fff0e8e0fff0d8d0ffe0d0c0ffa04020ffd06860ff ,
                        0xa05050ff0000000000000000f0a0a0ffffc0c0ffe07870ffffffffffffffffff ,
                        0xfffffffffffffffffff8f0fff0f0f0fff0e8e0fff0d8d0ffb04830ffd07070ff ,
                        0xa05050ff0000000000000000f0a8a0ffffc0c0ffe08080ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffff8f0fff0f0f0fff0e8e0ffb05030ffe07880ff ,
                        0xa05050ff0000000000000000f0b0b0ffffc0c0fff08890ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffff8f0fff0f0f0ffc05040ff603030ff ,
                        0xb05850ff0000000000000000f0b0b0ffffc0c0ffff9090ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffff8f0ffc05850ffb05860ff ,
                        0xb05860ff0000000000000000f0b8b0fff0b8b0fff0b0b0fff0b0b0fff0a8b0ff ,
                        0xf0a0a0ffe098a0ffe09090ffe09090ffe08890ffe08080ffd07880ffd07870ff ,
                        0xd07070ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6660
                    LayoutCachedTop =60
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =420
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1620
                    Top =60
                    Width =1020
                    Height =315
                    FontSize =9
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxStartDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter the date of this sampling visit started"
                    ConditionalFormat = Begin
                        0x010000009e000000020000000100000000000000000000001a00000001000000 ,
                        0x00000000fff2000000000000030000001b0000001e0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078005300740061007200740044006100740065005d002e005b00 ,
                        0x560061006c00750065005d003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedTop =60
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000190000005b00 ,
                        0x7400620078005300740061007200740044006100740065005d002e005b005600 ,
                        0x61006c00750065005d003d002200220000000000000000000000000000000000 ,
                        0x000000000000000000030000000100000000000000ffffff0002000000220022 ,
                        0x00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =75
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =8355711
                    ForeColor =255
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =75
                    LayoutCachedWidth =960
                    LayoutCachedHeight =375
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5760
                    Top =60
                    Width =720
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnUndo"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Undo/Clear values"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000f0906060d0784080b0583010000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000e0785040f08850ffd07040ffa05830500000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000f0906020d0704060f08050ffd07050f0a050300000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c06840d0f08850ffc078508000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0c0b01000000000000000000000000090482040e07840ffe08860ffe0a08000 ,
                        0x00000000000000000000000000000000d07040ffd07040ffc06840ffb06030ff ,
                        0xb05830ff905030ff0000000000000000b0603020c06840ffe08050ffd0886080 ,
                        0x00000000000000000000000000000000d07850ffe07030fff08050fff09870ff ,
                        0xe09060fff0a08040000000000000000080402000c06840ffe07840f0e09870c0 ,
                        0x00000000000000000000000000000000d08050ffe08050fff09060fff0a070ff ,
                        0x904830b0b0603040000000000000000080402000c06840ffd07040f0e09870d0 ,
                        0x00000000000000000000000000000000d08860ffe09060fff09870fff08850f0 ,
                        0xb06040ffb06040ffb060307000000000b0805020a05830f0d07840f0e09070d0 ,
                        0x000000000000000000000000e0b09010c08060ffd09870e0d0886090d09070ff ,
                        0xd08050ffc07040ffc06840ffb06030c0b07040e0a06040ffe08050ffd0a080e0 ,
                        0x00000000000000000000000000000000c08860ffd0a0804000000000d08860c0 ,
                        0xd08860ffd08050f0c06840ffb06840ffb06030f0e07840f0e0a080f0d09880e0 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0a880c0e09880ffe09870f0e09070f0e09070e0e0a080f0e0a890f0f0b8a020 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000f0b89060f0b090c0f0b8a0e0f0c0a0c0f0c0a090f0c0b02000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5760
                    LayoutCachedTop =60
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =420
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Subform
                    CanShrink = NotDefault
                    OverlapFlags =215
                    Left =105
                    Top =945
                    Width =7650
                    Height =4380
                    TabIndex =4
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.EventsList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =945
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =5325
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =825
                    Width =7860
                    Height =4620
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =825
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =5445
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7560
                    Top =105
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxID"
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =7560
                    LayoutCachedTop =105
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =405
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =7
                    Left =3480
                    Top =60
                    Width =1740
                    Height =315
                    FontSize =8
                    TabIndex =5
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x010000009a000000020000000100000000000000000000001800000001000000 ,
                        0x00000000fff200000000000003000000190000001c0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078005300740061007200740044006100740065005d002e005600 ,
                        0x61006c00750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxSite"
                    ControlSource ="Site_ID"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;0;0;0;0;0;1440"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Select site of this sampling visit"
                    GridlineColor =10921638
                    SeparatorCharacters =2
                    AllowValueListEdits =0

                    LayoutCachedLeft =3480
                    LayoutCachedTop =60
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000170000005b00 ,
                        0x7400620078005300740061007200740044006100740065005d002e0056006100 ,
                        0x6c00750065003d00220022000000000000000000000000000000000000000000 ,
                        0x0000000000030000000100000000000000ffffff000200000022002200000000 ,
                        0x000000000000000000000000000000000000
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =630
                    Width =7860
                    Height =315
                    FontSize =9
                    LeftMargin =360
                    TopMargin =36
                    RightMargin =360
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =65535
                    Name ="lblMsg"
                    Caption ="msg"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedTop =630
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =945
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =4320
                    Top =450
                    Width =825
                    Height =600
                    FontSize =20
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =16772541
                    Name ="lblMsgIcon"
                    Caption ="icon"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =4320
                    LayoutCachedTop =450
                    LayoutCachedWidth =5145
                    LayoutCachedHeight =1050
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1725
                    Top =390
                    Width =840
                    Height =180
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblHintDate"
                    Caption ="M\\DD\\YYYY"
                    GridlineColor =10921638
                    LayoutCachedLeft =1725
                    LayoutCachedTop =390
                    LayoutCachedWidth =2565
                    LayoutCachedHeight =570
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' Form:         Events
' Level:        Application form
' Version:      1.15
' Basis:        Dropdown frm
'
' Description:  Events form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
'               BLC - 7/26/2016 - 1.01 - added GetRecords() calls
'               BLC - 8/2/2016  - 1.02 - use Me.CallingForm
'               BLC - 8/22/2016 - 1.03 - added m_SaveOK, Form_BeforeUpdate()
'               BLC - 8/23/2016 - 1.04 - changed ReadyForSave() to public for
'                                        mod_App_Data Upsert/SetRecord()
'               BLC - 9/1/2016  - 1.05 - btnSave_Click code cleanup, remove ClearForm()
'               BLC - 10/20/2016 - 1.06 - removed buttoncaption, selectedID, selectedvalue properties
'                                         revised to use GetContext()
'               BLC - 1/9/2017   - 1.07 - revised to update VegPlot calling form event list,
'                                         hid form title
'               BLC - 1/13/2017 - 1.08 - revised to use && vs. StringFromCodepoint(uAmpersand)
'               BLC - 10/16/2017 - 1.09 - added date hint & uncommented ClearForm(),
'                                         reset combobox data sources on btnUndo_Click()
'               BLC - 10/17/2017 - 1.10 - added form BeforeUpdate() & AfterUpdate()
'               BLC - 10/18/2017 - 1.11 - enable comment button when ID is set (>0)
'               BLC - 10/24/2017 - 1.12 - comment out unused cbxSite
'               BLC - 11/6/2017  - 1.13 - set so Save button only enabled for new events (tbxID = 0)
'               BLC - 11/10/2017 - 1.14 - added Form_GotFocus for updating cbxLocation after new location added
'               BLC - 11/24/2017 - 1.15 - revised comment button caption
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

Private m_SaveOK As Boolean 'ok to save record (prevents bound form from immediately updating)

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
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  OpenArgs passes only the calling form name
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 6/27/2016 - updated to use ToggleForm() & ClearForm()
'   BLC - 7/26/2016 - added GetRecords() for cbxSite, cbxLocation recordsets
'   BLC - 8/2/2016 - use Me.CallingForm
'   BLC - 1/13/2017 - revised to use && vs. StringFromCodepoint(uAmpersand) for directions
'   BLC - 10/16/2017 - added date hint, uncommented ClearForm()
'   BLC - 10/24/2017 - commented out unused cbxSite
'   BLC - 11/24/2017 - revised comment button caption
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "Main"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize calling form
    ToggleForm Me.CallingForm, -1
    
    'set context - based on TempVars
    lblContext.ForeColor = lngLime
    lblContext.Caption = GetContext()
    
    Title = "Events (Sampling Visits)"
    lblTitle.Caption = "" 'hide second title
    Directions = "Choose the site location, then enter the sampling start date."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
    btnComment.Caption = StringFromCodepoint(uComment)
    btnComment.ForeColor = lngBlue

    
    'set hint
    Me.lblHintDate.Caption = "M\DD\YYYY"
    Me.lblHintDate.ForeColor = lngBlue
    
    'set hover
    btnComment.HoverColor = lngGreen
    btnSave.HoverColor = lngGreen
    btnUndo.HoverColor = lngGreen
      
    'defaults
    tbxIcon.ForeColor = lngRed
    btnComment.Enabled = False
    btnSave.Enabled = False
    tbxStartDate.BackColor = lngYellow
    lblMsgIcon.Caption = ""
    lblMsg.Caption = ""
  
    'ID default -> value used only for edits of existing table values
    tbxID.DefaultValue = 0
    
    'clear form datasource in case it was saved (to keep unbound)
    Me.RecordSource = ""
    
    'set data sources
    'Set cbxSite.Recordset = GetRecords("s_site_by_park_river_segment")
    Set cbxLocation.Recordset = GetRecords("s_location_by_park_river_segment")
    
    'initialize values
    ClearForm Me
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Events form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'eliminate NULLs
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Events form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 10/20/2016 - code cleanup
'   BLC - 10/18/2017 - enable comment button when ID is set (>0)
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
'FIX!
    'default
    btnSave.Enabled = False

'    If tbxID = 0 Then btnSave.Enabled = True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[Events form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_BeforeUpdate
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 22, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/22/2016 - initial version
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler
              
    If Not m_SaveOK Then
        Cancel = True
    End If
    'Cancel = True

    Me.lblMsg.Caption = StringFromCodepoint(uRArrow) & " Updating record..."

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[Events form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_AfterUpdate
' Description:  form after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 17, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/17/2017 - initial version
' ---------------------------------
Private Sub Form_AfterUpdate()
On Error GoTo Err_Handler
              
    Me.lblMsg.Caption = ""

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[Events form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxLocation_GotFocus
' Description:  combobox got focus actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 10, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/10/2017 - initial version
' ---------------------------------
Private Sub cbxLocation_GotFocus()
On Error GoTo Err_Handler
    
    'refresh Location - handles update after adding location
    Me.cbxLocation.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxLocation_GotFocus[Events form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxStartDate_AfterUpdate
' Description:  Dropdown after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub tbxStartDate_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxStartDate_AfterUpdate[Events form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnUndo_Click
' Description:  Undo button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 10/16/2017 - reset combobox data sources
'   BLC - 10/24/2017 - comment out unused cbxSite
' ---------------------------------
Private Sub btnUndo_Click()
On Error GoTo Err_Handler
    
    ClearForm Me
    
    'reset data sources
    'Set cbxSite.Recordset = GetRecords("s_site_by_park_river_segment")
    Set cbxLocation.Recordset = GetRecords("s_location_by_park_river_segment")
            
    cbxLocation.ControlSource = ""  'clear from Location_ID
    cbxLocation.Value = ""
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUndo_Click[Events form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSave_Click
' Description:  Save button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 9/1/2016  - cleanup commented code
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    'set enable btnSave_Click save
    m_SaveOK = True
    
'    'pre-save form
'    Me![list].Form.Dirty = False
    
    UpsertRecord Me
    
    Me![list].Form.Requery
    
    'revert to disable non-btnSave_Click save
    m_SaveOK = False
    
    'clear fields
    ClearForm Me
        
    cbxLocation.ControlSource = ""  'clear from Location_ID
    cbxLocation.Value = ""
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[Events form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAddLocation_Click
' Description:  Add location button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 24, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/24/2017 - initial version
' ---------------------------------
Private Sub btnAddLocation_Click()
On Error GoTo Err_Handler
    
    'open comment form
    DoCmd.OpenForm "Location", acNormal, , , , , "Event|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddLocation_Click[Event form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnComment_Click
' Description:  Undo button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub btnComment_Click()
On Error GoTo Err_Handler
    
    'open comment form
    DoCmd.OpenForm "Comment", acNormal, , , , , "event|" & tbxID & "|255"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[Events form])"
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
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 8/2/2016 - use Me.CallingForm
'   BLC - 1/9/2017 - revise to update VegPlot form's event combobox
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore calling form
    ToggleForm Me.CallingForm, 0
    
    'update events combobox for VegPlot
    Select Case Me.CallingForm
        Case "VegPlot"
'   show #DELETED vs. refreshing combo
'            Forms("VegPlot").Controls("cbxEvent").Requery
'            Forms("VegPlot").Controls("cbxEvent").Refresh
        Set Forms("VegPlot").Controls("cbxEvent").Recordset = GetRecords("s_events_by_park_river")
        Forms("VegPlot").Controls("cbxEvent").Requery
    End Select
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[Events form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ReadyForSave
' Description:  Check if form values are ready to save
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 8/23/2016 - changed ReadyForSave() to public for mod_App_Data Upsert/SetRecord()
'   BLC - 11/6/2017 - revised so Save button is only enabled for new events (tbxID = 0)
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    'requires: site ID, location ID, protocol ID, start date
    If IsDate(tbxStartDate) Then
        isOK = True
    End If
    
    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    'enable save button only for new events (tbxID = 0)
    If tbxID = 0 Then btnSave.Enabled = isOK
    
    'refresh form
'    Me.Requery
    
    'enable comment if ID > 0
    If tbxID > 0 Then btnComment.Enabled = True
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[Events form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          RunReadyForSave
' Description:  Run ready for save check from another form (public method)
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/26/2016 - initial version
' ---------------------------------
Public Sub RunReadyForSave()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RunReadyForSave[Events form])"
    End Select
    Resume Exit_Handler
End Sub
