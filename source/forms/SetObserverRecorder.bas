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
    ItemSuffix =42
    Left =3525
    Top =2490
    Right =12465
    Bottom =13875
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x747b49e35508e540
    End
    Caption ="Set Observer/Recorder"
    OnCurrent ="[Event Procedure]"
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
        Begin ListBox
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
            Height =1320
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =180
                    Top =60
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Set Observer/Recorder"
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
                    Caption ="Choose the observer && recorder for the record."
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =735
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
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =3360
                    Top =60
                    Width =4380
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblContext"
                    Caption ="context"
                    GridlineColor =10921638
                    LayoutCachedLeft =3360
                    LayoutCachedTop =60
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =1425
                    Top =960
                    Width =2115
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblObserver"
                    Caption ="Observer(s)"
                    GridlineColor =10921638
                    LayoutCachedLeft =1425
                    LayoutCachedTop =960
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1275
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3705
                    Top =960
                    Width =2055
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblRecorder"
                    Caption ="Recorder(s)"
                    GridlineColor =10921638
                    LayoutCachedLeft =3705
                    LayoutCachedTop =960
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =1275
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =108
                    Top =960
                    Width =1755
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =6750156
                    Name ="lblRecordRefID"
                    Caption ="refID"
                    GridlineColor =10921638
                    LayoutCachedLeft =108
                    LayoutCachedTop =960
                    LayoutCachedWidth =1863
                    LayoutCachedHeight =1275
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =8475
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6720
                    Top =60
                    Width =720
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

                    LayoutCachedLeft =6720
                    LayoutCachedTop =60
                    LayoutCachedWidth =7440
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
                    TabIndex =2
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
                    Left =5940
                    Top =60
                    Width =720
                    TabIndex =1
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

                    LayoutCachedLeft =5940
                    LayoutCachedTop =60
                    LayoutCachedWidth =6660
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
                    Top =3975
                    Width =7650
                    Height =4380
                    TabIndex =3
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.SetObserverRecorderList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =3975
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =8355
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =3855
                    Width =7860
                    Height =4620
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =3855
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =8475
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
                    TabIndex =4
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
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =3660
                    Width =7860
                    Height =315
                    FontSize =9
                    LeftMargin =360
                    TopMargin =36
                    RightMargin =360
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblMsg"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedTop =3660
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =3975
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =4320
                    Top =3480
                    Width =825
                    Height =600
                    FontSize =20
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =16772541
                    Name ="lblMsgIcon"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =4320
                    LayoutCachedTop =3480
                    LayoutCachedWidth =5145
                    LayoutCachedHeight =4080
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1380
                    Top =60
                    Width =2160
                    FontSize =8
                    TabIndex =5
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxObserver"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;2160"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="Select data observer(s)"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1380
                    LayoutCachedTop =60
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3660
                    Top =60
                    Width =2160
                    FontSize =8
                    TabIndex =6
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxRecorder"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;2160"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="Select data recorder(s)"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =3660
                    LayoutCachedTop =60
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1380
                    Top =1920
                    Width =2160
                    FontSize =8
                    TabIndex =7
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxObservers"
                    RowSourceType ="Value List"
                    ColumnWidths ="0;2160"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="Select data observer(s)"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1380
                    LayoutCachedTop =1920
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =3360
                    BackThemeColorIndex =-1
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3660
                    Top =1920
                    Width =2160
                    FontSize =8
                    TabIndex =8
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxRecorders"
                    RowSourceType ="Value List"
                    ColumnWidths ="0;2160"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="Select data recorder(s)"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =3660
                    LayoutCachedTop =1920
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =3360
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =1560
                    Width =2355
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblDir"
                    Caption ="Choose at least 1 of each"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =1560
                    LayoutCachedWidth =2475
                    LayoutCachedHeight =1860
                End
                Begin Label
                    OverlapFlags =85
                    Left =3480
                    Top =1560
                    Width =360
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblDownArrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =3480
                    LayoutCachedTop =1560
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =1875
                End
                Begin Label
                    OverlapFlags =85
                    Left =5940
                    Top =1380
                    Width =1500
                    Height =780
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblHintMoves"
                    Caption ="Double click to move up or down, SHIFT click to move multiple items"
                    GridlineColor =10921638
                    LayoutCachedLeft =5940
                    LayoutCachedTop =1380
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =2160
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
' Form:         SetObserverRecorder
' Level:        Application form
' Version:      1.02
' Basis:        Dropdown form
'
' Description:  Observer / recorder setting form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, September 8, 2016
' References:   -
' Revisions:    BLC - 9/1/2016  - 1.00 - initial version
'               BLC - 10/19/2017 - 1.01 - added comment length & replaced event w/ observerrecorder
'               BLC - 12/5/2017  - 1.02 - updated for multiple observers/recorders
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_ButtonCaption
Private m_SelectedID As Integer
Private m_SelectedValue As String
Private m_CallingForm As String

Private m_SaveOK As Boolean 'ok to save record (prevents bound form from immediately updating)

Private m_RefTable As String
Private m_RefID As Long

Private m_ObserverID As Long
Private m_RecorderID As Long

Private m_RAAction As String
Private m_RAContactID As Long

Private m_ActionDate As Date

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidLabel(Value As String)
Public Event InvalidCaption(Value As String)
Public Event InvalidCallingForm(Value As String)
Public Event InvalidRefTable(Value As String)
Public Event InvalidRefID(Value As Long)
Public Event InvalidRAAction(Value As String)
Public Event InvalidRAContactID(Value As Long)

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

Public Property Let ButtonCaption(Value As String)
    If Len(Value) > 0 Then
        m_ButtonCaption = Value

        'set the form button caption
        Me.btnSave.Caption = m_ButtonCaption
    Else
        RaiseEvent InvalidCaption(Value)
    End If
End Property

Public Property Get ButtonCaption() As String
    ButtonCaption = m_ButtonCaption
End Property

Public Property Let SelectedID(Value As Integer)
        m_SelectedID = Value
End Property

Public Property Get SelectedID() As Integer
    SelectedID = m_SelectedID
End Property

Public Property Let SelectedValue(Value As String)
        m_SelectedValue = Value
End Property

Public Property Get SelectedValue() As String
    SelectedValue = m_SelectedValue
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

Public Property Let RefTable(Value As String)
    If Len(Value) > 0 Then
        m_RefTable = Value
    Else
        RaiseEvent InvalidRefTable(Value)
    End If
End Property

Public Property Get RefTable() As String
    RefTable = m_RefTable
End Property

Public Property Get RefID() As Long
    RefID = m_RefID
End Property

Public Property Let RefID(Value As Long)
        m_RefID = Value
End Property

Public Property Get ObserverID() As Long
    ObserverID = m_ObserverID
End Property

Public Property Let ObserverID(Value As Long)
        m_ObserverID = Value
End Property

Public Property Get RecorderID() As Long
    RecorderID = m_RecorderID
End Property

Public Property Let RecorderID(Value As Long)
        m_RecorderID = Value
End Property

'Record Action properties
Public Property Let RAAction(Value As String)
    If Len(Value) > 0 Then
        If InStr(Value, "O") + InStr(Value, "R") > 0 Then
            m_RAAction = Value
        Else
            RaiseEvent InvalidRAAction(Value)
        End If
    Else
        RaiseEvent InvalidRAAction(Value)
    End If
End Property

Public Property Get RAAction() As String
    RAAction = m_RAAction
End Property

Public Property Get RAContactID() As Long
    RAContactID = m_RAContactID
End Property

Public Property Let RAContactID(Value As Long)
        m_RAContactID = Value
End Property

Public Property Get ActionDate() As Date
    ActionDate = m_ActionDate
End Property

Public Property Let ActionDate(Value As Date)
        m_ActionDate = Value
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
' Source/date:  Bonnie Campbell, September 8, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/8/2016 - initial version
'   BLC - 12/5/2017 - update for multiple observers/recorders
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "Main"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 And _
        Len(Nz(Me.OpenArgs, "")) <> Replace(Nz(Me.OpenArgs, ""), "|", "") Then
        
        'set the referencing table & record
        Me.CallingForm = Split(Me.OpenArgs, "|")(0)
        lblRecordRefID.Caption = Me.CallingForm _
                                & " ID # " & Split(Me.OpenArgs, "|")(1)
        
        'set properties
        Me.RefTable = Me.CallingForm
        Me.RefID = Split(Me.OpenArgs, "|")(1)
        Me.ActionDate = Split(Me.OpenArgs, "|")(2)
        
    End If

    'minimize Main
    ToggleForm Me.CallingForm, -1
    
    'set context - based on TempVars
    lblContext.ForeColor = lngLime
    lblContext.Caption = GetContext()
    'Nz(TempVars("ParkCode"), "") & Space(2) & ">" & Space(2) & _
    '             Nz(TempVars("River"), "")

    Title = "Set Observer/Recorder"
    Directions = "Choose the observer && recorder for the record."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
    btnComment.Caption = StringFromCodepoint(uComment)
    btnComment.ForeColor = lngBlue
        
    lblRecordRefID.ForeColor = lngLtLime
        
    'set hover
    btnComment.HoverColor = lngGreen
    btnSave.HoverColor = lngGreen
    btnUndo.HoverColor = lngGreen
      
    'defaults
    tbxIcon.ForeColor = lngRed
    btnComment.Enabled = False
    btnSave.Enabled = False
    lblMsgIcon.Caption = ""
    lblMsg.Caption = ""
    lblDir.Caption = "Choose at least 1 of each"
    lblDownArrow.Caption = StringFromCodepoint(uDArrow)
    lblHintMoves.Caption = "Double click to move up or down, SHIFT click to move multiple items"
    lbxObservers.BackColor = lngWhite
    lbxRecorders.BackColor = lngWhite
    lbxObservers.BackColor = lngYellow
    lbxRecorders.BackColor = lngYellow
    
    'set as multiselect listboxes
    '2 = Extended - SHFT click to move multiples
'    lbxObserver.MultiSelect = 2
'    lbxRecorder.MultiSelect = 2
'    lbxObservers.MultiSelect = 2
'    lbxRecorders.MultiSelect = 2
  
    'ID default -> value used only for edits of existing table values
    tbxID.DefaultValue = 0
    
    'hide unused controls
    btnUndo.Visible = False
    tbxID.Visible = False
    
    'clear form datasource in case it was saved (to keep unbound)
    Me.RecordSource = ""
    
    'set data sources
    Set lbxObserver.Recordset = GetRecords("s_contact_list")
    Set lbxRecorder.Recordset = GetRecords("s_contact_list")
    lbxObserver.BoundColumn = 1
    lbxObserver.ColumnHeads = True
    'lbxObserver.ColumnWidths = ""
    lbxRecorder.BoundColumn = 1
    lbxRecorder.ColumnHeads = True
    'lbxRecorder.ColumnHeads = ""
    lbxObservers.BoundColumn = 1
    lbxRecorders.BoundColumn = 1
    
    'set columns same for selected Observers/Recorders
    lbxObservers.ColumnWidths = lbxObserver.ColumnWidths
    lbxRecorders.ColumnWidths = lbxRecorder.ColumnWidths
    
    'clear selected Observers/Recorders
    lbxObservers.RowSource = ""
    lbxRecorders.RowSource = ""
    
    'set list data source
    Dim Params(0 To 2) As Variant
    
    Params(0) = Me.RefTable
    Params(1) = Me.RefID
    
    Set Me.list.Form.Recordset = GetRecords("s_record_action_by_refID", Params)
    Me.list.Form.Requery
    
    'initialize values
    ClearForm Me
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[SetObserverRecorder form])"
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
' Source/date:  Bonnie Campbell, September 8, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/8/2016 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Load[SetObserverRecorder form])"
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
              
'      If tbxID > 0 Then btnComment.Enabled = True
    'MsgBox tbxID, vbCritical, "Current"
    'If tbxID > 0 Then ReadyForSave

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[SetObserverRecorder form])"
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
Private Sub zForm_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler
              
    If Not m_SaveOK Then
        Cancel = True
    End If
    'Cancel = True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[SetObserverRecorder form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lbxObserver_AfterUpdate
' Description:  Dropdown after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 5, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/5/2017 - initial version
' ---------------------------------
Private Sub lbxObserver_AfterUpdate()
On Error GoTo Err_Handler

    'Me.ObserverID = lbxObserver.Value
    
    'ReadyForSave
    'lbxObservers.AddItem lbxObserver.Column(0)
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxObserver_AfterUpdate[SetObserverObserver form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lbxRecorder_AfterUpdate
' Description:  Listbox double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 5, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/5/2017 - initial version
' ---------------------------------
Private Sub lbxRecorder_AfterUpdate()
On Error GoTo Err_Handler

    'Me.RecorderID = lbxRecorder.Value
    
    'ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxRecorder_AfterUpdate[SetObserverRecorder form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lbxObserver_DblClick
' Description:  Listbox double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 5, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/5/2017 - initial version
' ---------------------------------
Private Sub lbxObserver_DblClick(Cancel As Integer)
On Error GoTo Err_Handler

    'Me.ObserverID = lbxObserver.Value
    
    'ReadyForSave
    Dim i As Integer
    
    With lbxObservers
        For i = 0 To lbxObserver.ListCount - 1
            If lbxObserver.Selected(i) Then
                '.ItemData(.NewIndex) = lbxObserver.Column(0)
                'check if duplicate
                If Not IsDupeItem(lbxObservers, lbxObserver.ItemData(i)) Then _
                    .AddItem Item:=lbxObserver.Column(0, i) & ";" & lbxObserver.Column(1, i)
                    '.AddItem item:=lbxObserver.ListIndex & ";" & lbxObserver.Column(1)
            End If
            
'            lbxObserver.RemoveItem lbxObserver.ListIndex
        Next
        
        If .ListCount > 0 Then
            .BackColor = lngWhite
        Else
            .BackColor = lngYellow
        End If
    End With
    
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxObserver_DblClick[SetObserverObserver form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lbxRecorder_DblClick
' Description:  Listbox double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 5, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/5/2017 - initial version
' ---------------------------------
Private Sub lbxRecorder_DblClick(Cancel As Integer)
On Error GoTo Err_Handler

    'Me.RecorderID = lbxRecorder.Value
    
    'ReadyForSave
    Dim i As Integer
    
    With lbxRecorders
        For i = 0 To lbxRecorder.ListCount - 1
            If lbxRecorder.Selected(i) Then
                '.ItemData(.NewIndex) = lbxRecorder.Column(0)
                
                'check if duplicate
                If Not IsDupeItem(lbxRecorders, lbxRecorder.ItemData(i)) Then _
                    .AddItem Item:=lbxRecorder.Column(0, i) & ";" & lbxRecorder.Column(1, i)
            End If
        Next
    
        If .ListCount > 0 Then
            .BackColor = lngWhite
        Else
            .BackColor = lngYellow
        End If
    
    End With
        
    ReadyForSave
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxRecorder_DblClick[SetObserverRecorder form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lbxObservers_DblClick
' Description:  Listbox double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 5, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/5/2017 - initial version
' ---------------------------------
Private Sub lbxObservers_DblClick(Cancel As Integer)
On Error GoTo Err_Handler

    Dim i As Integer
    
    With lbxObservers
        For i = 0 To lbxObservers.ListCount - 1
            If lbxObservers.Selected(i) Then
                .RemoveItem lbxObservers.ListIndex
            End If
        Next
    End With
    
    If ListHasItems(lbxObservers) = False Then lbxObservers.BackColor = lngYellow
    
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxObserver_DblClick[SetObserverObserver form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lbxRecorders_DblClick
' Description:  Listbox double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 5, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/5/2017 - initial version
' ---------------------------------
Private Sub lbxRecorders_DblClick(Cancel As Integer)
On Error GoTo Err_Handler

    Dim i As Integer
    
    With lbxRecorders
        For i = 0 To lbxRecorders.ListCount - 1
            If lbxRecorders.Selected(i) Then
                .RemoveItem lbxRecorders.ListIndex
            End If
        Next
    End With
    
    If ListHasItems(lbxRecorders) = False Then lbxRecorders.BackColor = lngYellow
    
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxRecorder_DblClick[SetObserverRecorder form])"
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
            "Error encountered (#" & Err.Number & " - btnUndo_Click[SetObserverRecorder form])"
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
' Source/date:  Bonnie Campbell, September 8, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/8/2016 - initial version
'   BLC - 9/1/2016  - cleanup commented code
'   BLC - 12/5/2017 - updated for multiple observers/recorders
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    'set enable btnSave_Click save
    m_SaveOK = True
    
'    'pre-save form
'    Me![list].Form.Dirty = False
    
'    'run for observer
'    Me.RAAction = "O"
'    Me.RAContactID = Me.ObserverID
'
'    UpsertRecord Me
'
'    'run for recorder
'    Me.RAAction = "R"
'    Me.RAContactID = Me.RecorderID
'
'    UpsertRecord Me

    'ensure RefTable & RefID are set
    If Len(Me.RefTable) > 0 And Me.RefID > 0 And IsDate(Me.ActionDate) Then

        Dim i As Integer
        
        For i = 0 To lbxObservers.ListCount - 1
            Debug.Print "O:" & lbxObservers.ItemData(i)
            SetObserverRecorder Me, Me.RefTable, "O", CLng(lbxObservers.ItemData(i))
        Next
        
        For i = 0 To lbxRecorders.ListCount - 1
            Debug.Print "R:" & lbxRecorders.ItemData(i)
            SetObserverRecorder Me, Me.RefTable, "R", CLng(lbxRecorders.ItemData(i))
        Next
        'SetObserverRecorder Me, Me.RefTable, action, ContactID
    
    End If
    
    Me![list].Form.Requery
    
    'revert to disable non-btnSave_Click save
    m_SaveOK = False
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[SetObserverRecorder form])"
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
'   BLC - 10/19/2017 - added comment length & replaced event w/ observerrecorder
' ---------------------------------
Private Sub btnComment_Click()
On Error GoTo Err_Handler
    
    'open comment form
    DoCmd.OpenForm "Comment", acNormal, , , , , "observerrecorder|" & tbxID & "|255"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[SetObserverRecorder form])"
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
' Source/date:  Bonnie Campbell, September 8, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/8/2016 - initial version
'   BLC - 8/2/2016 - use Me.CallingForm
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
            "Error encountered (#" & Err.Number & " - Form_Close[SetObserverRecorder form])"
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
' Source/date:  Bonnie Campbell, September 8, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/8/2016 - initial version
'   BLC - 8/23/2016 - changed ReadyForSave() to public for mod_App_Data Upsert/SetRecord()
'   BLC - 12/5/2017 - update for multiple observer/recorder
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    'requires: site ID, location ID, protocol ID, start date
    If lbxObservers.ListCount > 0 And lbxRecorders.ListCount Then
        isOK = True
    End If
    
    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    btnSave.Enabled = isOK
    
    'refresh form
'    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[SetObserverRecorder form])"
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
' Source/date:  Bonnie Campbell, September 8, 2016 - for NCPN tools
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
            "Error encountered (#" & Err.Number & " - RunReadyForSave[SetObserverRecorder form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     IsDupeItem
' Description:  Determine if a list item is already present
' Assumptions:  -
' Parameters:   lbx - Listbox to check (listbox)
'               item - item being added (variant)
' Returns:      DupeItem - whether the item is a duplicate or not (boolean)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 5, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/5/2017 - initial version
' ---------------------------------
Public Function IsDupeItem(lbx As ListBox, Item As Variant) As Boolean
On Error GoTo Err_Handler

    Dim DupeItem As Boolean
    Dim i As Integer
    
    'default
    DupeItem = False
    
    For i = 0 To lbx.ListCount - 1
    
'        If lbx.ListIndex = item Then DupeItem = True
    
        If lbx.ItemData(i) = Item Then
            DupeItem = True
            Exit For
        End If
    Next
    
Exit_Handler:
    IsDupeItem = DupeItem
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsDupeItem[SetObserverRecorder form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     ListHasItems
' Description:  Determine if a list has items
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 5, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/5/2017 - initial version
' ---------------------------------
Public Function ListHasItems(lbx As ListBox) As Boolean
On Error GoTo Err_Handler
    
    If lbx.ListCount > 0 Then
        ListHasItems = True
    Else
        ListHasItems = False
    End If
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ListHasItems[SetObserverRecorder form])"
    End Select
    Resume Exit_Handler
End Function
