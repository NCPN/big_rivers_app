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
    Width =8220
    DatasheetFontHeight =11
    ItemSuffix =60
    Left =3360
    Top =2775
    Right =23250
    Bottom =14625
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x20f1f0c46fcee440
    End
    Caption ="Contact"
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
        Begin ToggleButton
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackColor =-1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-1
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverColor =0
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedColor =0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeColor =0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeColor =0
            PressedForeThemeColorIndex =1
        End
        Begin FormHeader
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
                    Caption ="Contact"
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
                    Caption ="Enter the contact information and click save."
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
                    Left =3120
                    Top =1080
                    Width =1500
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblLast"
                    Caption ="Last"
                    GridlineColor =10921638
                    LayoutCachedLeft =3120
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =1065
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblNumber"
                    Caption ="First"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =1065
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =1380
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =2700
                    Top =1065
                    Width =360
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblSampleDate"
                    Caption ="MI"
                    GridlineColor =10921638
                    LayoutCachedLeft =2700
                    LayoutCachedTop =1065
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =1380
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6660
                    Top =900
                    Width =720
                    ForeColor =16711680
                    Name ="btnComment"
                    Caption ="í ½í·©"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6660
                    LayoutCachedTop =900
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =1260
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =8010
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    Left =240
                    Top =960
                    Width =7620
                    Height =2100
                    FontSize =18
                    FontWeight =300
                    LeftMargin =288
                    TopMargin =72
                    BackColor =11262179
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblWork"
                    Caption ="Work\015\012Info"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =960
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =3060
                    BackThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6660
                    Top =60
                    Width =720
                    TabIndex =10
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
                    TabIndex =13
                    BorderColor =8355711
                    ForeColor =690698
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
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =5760
                    Top =60
                    Width =720
                    TabIndex =11
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =105
                    Top =3510
                    Width =7995
                    Height =4380
                    TabIndex =12
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.ContactList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =3510
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =7890
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =3390
                    Width =8220
                    Height =4620
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =3390
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =8010
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
                    TabIndex =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxID"
                    ControlSource ="Contact.ID"
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
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2580
                    Top =60
                    Width =360
                    Height =315
                    TabIndex =1
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxMI"
                    ControlSource ="MiddleInitial"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="Enter your middle initial"
                    ConditionalFormat = Begin
                        0x010000009a000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000180000001b0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d002200220000000000000022002200000000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =60
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Height =315
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFirst"
                    ControlSource ="FirstName"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter your first name"
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3060
                    Top =60
                    Width =1680
                    Height =315
                    TabIndex =2
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLast"
                    ControlSource ="LastName"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter your last name"
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3060
                    LayoutCachedTop =60
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1860
                    Top =2100
                    Width =2040
                    Height =315
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxOrganization"
                    ControlSource ="Organization"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter your organization abbreviation (NCPN, etc.)"
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1860
                    LayoutCachedTop =2100
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =2415
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =540
                    Top =2100
                    Width =1260
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblOrganization"
                    Caption ="Organization"
                    GridlineColor =10921638
                    LayoutCachedLeft =540
                    LayoutCachedTop =2100
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =2415
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1860
                    Top =2535
                    Width =2040
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPosition"
                    ControlSource ="PositionTitle"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter your position"
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1860
                    LayoutCachedTop =2535
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =2850
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =540
                    Top =2535
                    Width =1260
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPosition"
                    Caption ="Position"
                    GridlineColor =10921638
                    LayoutCachedLeft =540
                    LayoutCachedTop =2535
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =2850
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =4380
                    Top =1140
                    Width =3300
                    Height =315
                    TabIndex =4
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxEmail"
                    ControlSource ="Email"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter your work email address (first_last@nps.gov, etc.)"
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =1140
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1455
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =3060
                    Top =1140
                    Width =1260
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblEmail"
                    Caption ="Email"
                    GridlineColor =10921638
                    LayoutCachedLeft =3060
                    LayoutCachedTop =1140
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1455
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5640
                    Top =2100
                    Width =2040
                    Height =315
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPhone"
                    ControlSource ="WorkPhone"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter your work phone #"
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =2100
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =2415
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =4320
                    Top =2100
                    Width =1260
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPhone"
                    Caption ="Phone"
                    GridlineColor =10921638
                    LayoutCachedLeft =4320
                    LayoutCachedTop =2100
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =2415
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5640
                    Top =2535
                    Width =2040
                    Height =315
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxExtension"
                    ControlSource ="WorkExtension"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter your work extension (if any)"
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =2535
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =2850
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =4320
                    Top =2535
                    Width =1260
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblExtension"
                    Caption ="Extension"
                    GridlineColor =10921638
                    LayoutCachedLeft =4320
                    LayoutCachedTop =2535
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =2850
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =4380
                    Top =1575
                    Width =3300
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUsername"
                    ControlSource ="Username"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter your username (AD)"
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =1575
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1890
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =3060
                    Top =1575
                    Width =1260
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblUsername"
                    Caption ="Username"
                    GridlineColor =10921638
                    LayoutCachedLeft =3060
                    LayoutCachedTop =1575
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1890
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =6660
                    Top =540
                    Width =720
                    TabIndex =15
                    ForeColor =4210752
                    Name ="btnSetRole"
                    Caption ="Edit"
                    ControlTipText ="Set user's role"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b04830ffb03830ffa03030ff903020ff903030ff00000000 ,
                        0x0000000040a040ff309020ff308020ff307020ff206020ff0000000000000000 ,
                        0x0000000000000000d04010ffd05020ffd05020ffc04010ff900800ff00000000 ,
                        0x0000000020a800ff30b820ff30b020ff20a800ff006800ff0000000000000000 ,
                        0x0000000000000000f06020fff09060fff08850ffe06020ff901000ff5058a0ff ,
                        0x505880ff30c810ff70e860ff70e060ff40c820ff107000ff404090ff504890ff ,
                        0x504880ff606890fff06830fff09870fff08860ffe06020ffa02010ff0088e0ff ,
                        0x0070e0ff40d030ff90f070ff80e870ff40c820ff107000ff1020c0ff0008e0ff ,
                        0x0000c0ff100860ffff7830fff0a890fff0a080ffd06830ff801810ff00c0ffff ,
                        0x00a0ffff40d030ffa0f890ffa0f080ff50c840ff106800ff4058f0ff3030ffff ,
                        0x0008d0ff000060ffff8040ffffd8c0ffffc8a0ffe08050ffa02010ff00c8ffff ,
                        0x00a8ffff40d840ffd0ffc0ffc0ffc0ff70d860ff108000ff8080ffff5058f0ff ,
                        0x0010c0ff000050ffffb080ffffb890ffffa880fff07030ff905840ff00c8ffff ,
                        0x10b0ffff30b880ffa0f890ffa0f890ff50d840ff20a030ff8088ffff6068f0ff ,
                        0x0010c0ff000850ff00000000d04020ffd01800ffa01810ff30a0c0ff30f0ffff ,
                        0x10c0ffff0070a0ff109800ff10a800ff008800ff306890ff9090ffff7070ffff ,
                        0x0010c0ff202870ffe07050ffd05830ffd06830ffc02810ffa04030ffa0ffffff ,
                        0x20e0ffff009840ff40b820ff50c030ff10a010ff108820ffc0d0f0ff9090ffff ,
                        0x1020d0ff202090ffe06020ffffc8a0ffffb890ffe07040ffc01800ff40c8f0ff ,
                        0x10b0ffff30c030ffb0f8a0ffb0f8a0ff60d050ff109800ff6070e0ff4040ffff ,
                        0x1018e0ff7078e0fff08850ffffd0b0ffffd8b0fff07040ffa02010ff0058c0ff ,
                        0x0048b0ff30b850ffc0ffa0ffd0ffd0ff70e050ff109000ff0000b0ff0000b0ff ,
                        0x3038b0ff0000000000000000f08860fff07040ffc04020ff306890ff00a8f0ff ,
                        0x0078e0ff007090ff50d050ff60e840ff30b010ff104870ff3030e0ff2020e0ff ,
                        0x0000b0ff5050b0ff000000000000000000000000b0a8c0ff20d8ffff80ffffff ,
                        0x20d0ffff0060e0ff70b0b0ff00000000a0c0c0ff5048ffffc0c0ffff8080ffff ,
                        0x1018d0ff2020a0ff000000000000000000000000b0e0ffff20c8ffff90ffffff ,
                        0x30e0ffff0070d0ffa0b8e0ff00000000d0d8ffff4048ffffc0c0ffff8090ffff ,
                        0x1010d0ff5050c0ff0000000000000000000000000000000080d0ffff10b0ffff ,
                        0x1090f0ff70a8e0ff000000000000000000000000a0a8ffff3038f0ff2028f0ff ,
                        0x5058e0ff00000000
                    End

                    LayoutCachedLeft =6660
                    LayoutCachedTop =540
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =900
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
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2580
                    Top =480
                    Width =2880
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"30\""
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxUserRole"
                    ControlSource ="Access_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT * FROM Access; "
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Choose the contact's application role"
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =480
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =795
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =900
                    Top =480
                    Width =1605
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblUserRole"
                    Caption ="Application Role"
                    GridlineColor =10921638
                    LayoutCachedLeft =900
                    LayoutCachedTop =480
                    LayoutCachedWidth =2505
                    LayoutCachedHeight =795
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =3165
                    Width =8220
                    Height =315
                    FontSize =9
                    LeftMargin =360
                    TopMargin =36
                    RightMargin =360
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblMsg"
                    Caption ="msg"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedTop =3165
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =3480
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =4320
                    Top =2985
                    Width =825
                    Height =600
                    FontSize =20
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblMsgIcon"
                    Caption ="Icon"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =4320
                    LayoutCachedTop =2985
                    LayoutCachedWidth =5145
                    LayoutCachedHeight =3585
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
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
' Form:         Contact
' Level:        Application form
' Version:      1.06
' Basis:        Dropdown form
'
' Description:  Contact form object related properties, Contact, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, June 20, 2016
' References:   -
' Revisions:    BLC - 6/20/2016 - 1.00 - initial version
'               BLC - 6/24/2016 - 1.01 - replaced Exit_Function > Exit_Handler
'               BLC - 6/27/2016 - 1.02 - adjusted for ToggleForm(),
'                                        shifted ClearForm() to mod_Forms
'               BLC - 8/23/2016 - 1.03 - changed ReadyForSave() to public for
'                                        mod_App_Data Upsert/SetRecord()
'               BLC - 8/29/2016 - 1.04 - code cleanup btnSave_Click() uses UpsertRecord()
'               BLC - 10/17/2016 - 1.05 - revise to restore calling form vs. Main/DbAdmin only
'               BLC - 10/25/2016 - 1.06 - revised to remove strParent calls, Main calling form
'                                         passed via ClickAction()
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
Private m_ContinueSave As Boolean 'ok to proceed saving (prevent Property Not Found error?)

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidLabel(Value As String)
Public Event InvalidCaption(Value As String)
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
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
'   BLC - 6/27/2016 - adjusted for ToggleForm()
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    'default
    Me.CallingForm = "DbAdmin"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize Calling Form
    ToggleForm Me.CallingForm, -1

    Title = "Contact"
    Directions = "Enter the contact information and click save."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
    btnComment.Caption = StringFromCodepoint(uComment)
    btnComment.ForeColor = lngBlue
    lblWork.Caption = "Work" & vbCrLf & "Info"
    
    'set hover
    btnComment.HoverColor = lngGreen
    btnSave.HoverColor = lngGreen
    btnUndo.HoverColor = lngGreen
      
    'defaults
    lblWork.BackColor = lngCream
    tbxIcon.ForeColor = lngRed
    btnComment.Enabled = False
    btnSave.Enabled = False
    lblMsgIcon.Caption = ""
    lblMsg.Caption = ""

    'unbind the control from Access_ID to allow selections (cleans up prior uses)
    cbxUserRole.ControlSource = ""
    cbxUserRole.RowSource = GetTemplate("s_access")
    
    'ID default -> value used only for edits of existing table values
    tbxID.DefaultValue = 0
    
    'initialize values
    ClearForm Me
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Contact form])"
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
' Source/date:  Bonnie Campbell, June 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
'   BLC - 6/27/2016 - adjusted for ToggleForm
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
            "Error encountered (#" & Err.Number & " - Form_Load[Contact form])"
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
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
              
      'If Nz(tbxID, 0) > 0 Then btnComment.Enabled = True
    Debug.Print Me.Name & " Form_Current Me.Dirty IS " & Me.Dirty
    Debug.Print Me.Name & " Form_Current Me.list.Form.Dirty IS " & Me.list.Form.Dirty

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[Contact form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_BeforeUpdate
' Description:  form before update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 29, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/29/2016 - initial version
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    Debug.Print Me.Name & " Form_BeforeUpdate SaveOK = " & m_SaveOK

    Debug.Print Me.Name & " Form_BeforeUpdate Me.Dirty IS " & Me.Dirty
    Debug.Print Me.Name & " Form_BeforeUpdate Me.list.Form.Dirty IS " & Me.list.Form.Dirty

    'avoid actions when clearing form
    If RefSub = "ClearForm" Then GoTo Exit_Handler
'
    If m_SaveOK <> True Then
        Cancel = True
        'm_ContinueSave = False
    Else
        Cancel = True
        'm_ContinueSave = True
    End If

    Debug.Print Me.Name & " Form_BeforeUpdate after cnx Me.Dirty IS " & Me.Dirty
    Debug.Print Me.Name & " Form_BeforeUpdate after cnx Me.list.Form.Dirty IS " & Me.list.Form.Dirty


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[Contact form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_AfterUpdate
' Description:  form before update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Unknown, unknown
'   http://msaccesstipsandtricks.blogspot.com/2013/04/solution-for-property-not-found-error.html
' Source/date:  Bonnie Campbell, August 29, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/29/2016 - initial version
' ---------------------------------
Private Sub Form_AfterUpdate()
On Error GoTo Err_Handler

    Debug.Print Me.Name & " Form_AfterUpdate Me.Dirty IS " & Me.Dirty
    Debug.Print Me.Name & " Form_AfterUpdate Me.list.Form.Dirty IS " & Me.list.Form.Dirty

'    Debug.Print "ContinueSave = " & m_ContinueSave
'
'    If m_ContinueSave <> True Then
'        'undo the save only if the record is
'        Debug.Print Me.Name & " Form_AfterUpdate UNDOING..."
'        Me.list.Form.Undo
'    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_AfterUpdate[Contact form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxFirst_AfterUpdate
' Description:   AfterUpdate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 22, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/22/2016 - initial version
' ---------------------------------
Private Sub tbxFirst_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxFirst.text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxFirst_AfterUpdate[Contact form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxLast_AfterUpdate
' Description:   AfterUpdate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 22, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/22/2016 - initial version
' ---------------------------------
Private Sub tbxLast_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxLast.text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxLast_AfterUpdate[Contact form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxEmail_AfterUpdate
' Description:   AfterUpdate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 22, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/22/2016 - initial version
' ---------------------------------
Private Sub tbxEmail_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxEmail.text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxEmail_AfterUpdate[Contact form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxUsername_AfterUpdate
' Description:   AfterUpdate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 22, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/22/2016 - initial version
' ---------------------------------
Private Sub tbxUsername_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxUsername.text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxUsername_AfterUpdate[Contact form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxOrganization_AfterUpdate
' Description:   AfterUpdate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 22, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/22/2016 - initial version
' ---------------------------------
Private Sub tbxOrganization_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxOrganization.text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxOrganization_AfterUpdate[Contact form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPosition_AfterUpdate
' Description:   AfterUpdate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 22, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/22/2016 - initial version
' ---------------------------------
Private Sub tbxPosition_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPosition.text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPosition_AfterUpdate[Contact form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPhone_AfterUpdate
' Description:   AfterUpdate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 22, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/22/2016 - initial version
' ---------------------------------
Private Sub tbxPhone_AfterUpdate()
On Error GoTo Err_Handler

    'If Len(tbxPhone.Text) > 0 Then _

        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPhone_AfterUpdate[Contact form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxExtension_AfterUpdate
' Description:   AfterUpdate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 22, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/22/2016 - initial version
' ---------------------------------
Private Sub tbxExtension_AfterUpdate()
On Error GoTo Err_Handler

    'If Len(tbxExtension.Text) > 0 Then _

        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxExtension_AfterUpdate[Contact form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxUserRole_AfterUpdate
' Description:  AfterUpdate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 22, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/22/2016 - initial version
' ---------------------------------
Private Sub cbxUserRole_AfterUpdate()
On Error GoTo Err_Handler

    If Len(cbxUserRole.text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxUserRole_AfterUpdate[Contact form])"
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
' ---------------------------------
Private Sub btnUndo_Click()
On Error GoTo Err_Handler
    
    ClearForm Me
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUndo_Click[Contact form])"
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
' Source/date:  Bonnie Campbell, June 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
'   BLC - 8/23/2016 - use UpsertRecord()
'   BLC - 8/29/2016 - code cleanup
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    'set enable btnSave_Click save
    m_SaveOK = True
    
    UpsertRecord Me
    
    'revert to disable non-btnSave_Click save
    m_SaveOK = False
    
    'refresh the list to avoid Name #Type!
    'Me.list.Requery

    Me.Requery

    'clear the dirt
    'Me.Dirty = False
    
    'clear the form
    ClearForm Me
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[Contact form])"
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
    DoCmd.OpenForm "Comment", acNormal, , , , , "Contact|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[Contact form])"
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
' Source/date:  Bonnie Campbell, June 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
'   BLC - 10/17/2016 - revise to restore calling form vs. Main/DbAdmin only
'   BLC - 10/25/2016 - revised to remove strParent calls, Main calling form passed via ClickAction()
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
            "Error encountered (#" & Err.Number & " - Form_Close[Contact form])"
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
' Source/date:  Bonnie Campbell, June 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
'   BLC - 8/23/2016 - changed to public for mod_App_Data UpsertRecord
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean
    
    Debug.Print Me.Name & " Form_Current Me.Dirty IS " & Me.Dirty
    Debug.Print Me.Name & " Form_Current Me.list.Form.Dirty IS " & Me.list.Form.Dirty

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    'requires: first, last, email', username, org, pos, phone, ext
    If Len(Nz(tbxFirst.Value, "")) > 0 _
        And Len(Nz(tbxLast.Value, "")) > 0 _
        And IsEmail(Nz(tbxEmail.Value, "")) Then '_
        'And Len(Nz(tbxUsername.Value, "")) > 0 _
        'And Len(Nz(tbxOrganization.Value, "")) > 0 Then
        isOK = True
    End If
    
    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    btnSave.Enabled = isOK
    
    'refresh form
    Me.Requery
    
    'check for unsaved changes
    If Me.Dirty Then
        lblMsg.ForeColor = lngYellow
        lblMsgIcon.ForeColor = lngYellow
        lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
        lblMsg.Caption = "** UNSAVED CHANGES! **"
        
        Debug.Print "ReadyForSave " & Me.Name & " DIRTY"

    Else
        Debug.Print "ReadyForSave " & Me.Name & " CLEAN"
    
        lblMsgIcon.Caption = ""
        lblMsg.Caption = ""
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[Contact form])"
    End Select
    Resume Exit_Handler
End Sub
