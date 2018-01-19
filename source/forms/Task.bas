Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9000
    DatasheetFontHeight =11
    ItemSuffix =32
    Left =4035
    Top =3045
    Right =13980
    Bottom =14895
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x6370e280b30de540
    End
    Caption ="Task"
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
    SplitFormSplitterBar =0
    SaveSplitterBarPosition =0
    SplitFormSplitterBar =0
    SaveSplitterBarPosition =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
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
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
            Height =734
            BackColor =4538399
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    OverlapFlags =85
                    Top =720
                    Width =7859
                    Height =14
                    BorderColor =65280
                    Name ="lineIndicator"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedTop =720
                    LayoutCachedWidth =7859
                    LayoutCachedHeight =734
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =30
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =30
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =330
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =120
                    Top =360
                    Width =6840
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Enter task details."
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =360
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =675
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =4800
                    Top =60
                    Width =4140
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblContext"
                    Caption ="context"
                    GridlineColor =10921638
                    LayoutCachedLeft =4800
                    LayoutCachedTop =60
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =8340
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =240
                    Top =900
                    Width =8400
                    Height =1380
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTask"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff000100000000000000040000001f0000000100000000000000 ,
                        0xfff2000000000000000000000000000000000000000000000000000000000000 ,
                        0x220022000000000049004900660028004c0065006e0028005b00740062007800 ,
                        0x5400610073006b005d0029003d00220022002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =900
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =2280
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000fff200001a00000049004900660028004c0065006e0028005b ,
                        0x007400620078005400610073006b005d0029003d00220022002c0031002c0030 ,
                        0x002900000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1080
                    Left =1020
                    Top =60
                    Width =1680
                    Height =360
                    TabIndex =1
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"30\""
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff00010000000000000004000000210000000100000000000000 ,
                        0xfff2000000000000000000000000000000000000000000000000000000000000 ,
                        0x220022000000000049004900660028004c0065006e0028005b00630062007800 ,
                        0x5300740061007400750073005d0029003d00220022002c0031002c0030002900 ,
                        0x00000000
                    End
                    Name ="cbxStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ID, Status, Icon, Sequence FROM Status ORDER BY Sequence; "
                    ColumnWidths ="0;1080"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =420
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000fff200001c00000049004900660028004c0065006e0028005b ,
                        0x006300620078005300740061007400750073005d0029003d00220022002c0031 ,
                        0x002c0030002900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =360
                            Top =60
                            Width =984
                            Height =314
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblStatus"
                            Caption ="Status"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =60
                            LayoutCachedWidth =1344
                            LayoutCachedHeight =374
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1080
                    Left =6840
                    Top =60
                    Width =1260
                    Height =360
                    TabIndex =2
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"30\""
                    ConditionalFormat = Begin
                        0x01000000a8000000020000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff00010000000000000004000000230000000100000000000000 ,
                        0xfff2000000000000000000000000000000000000000000000000000000000000 ,
                        0x220022000000000049004900660028004c0065006e0028005b00630062007800 ,
                        0x5000720069006f0072006900740079005d0029003d00220022002c0031002c00 ,
                        0x3000290000000000
                    End
                    Name ="cbxPriority"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ID, Priority, Sequence FROM Priority; "
                    ColumnWidths ="0;1080"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6840
                    LayoutCachedTop =60
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =420
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000fff200001e00000049004900660028004c0065006e0028005b ,
                        0x006300620078005000720069006f0072006900740079005d0029003d00220022 ,
                        0x002c0031002c0030002900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6060
                            Top =60
                            Width =984
                            Height =314
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPriority"
                            Caption ="Priority"
                            GridlineColor =10921638
                            LayoutCachedLeft =6060
                            LayoutCachedTop =60
                            LayoutCachedWidth =7044
                            LayoutCachedHeight =374
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7200
                    Top =2340
                    Width =720
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnSave"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Save Record"
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

                    LayoutCachedLeft =7200
                    LayoutCachedTop =2340
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =2700
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
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    TextFontFamily =2
                    Left =8100
                    Top =2340
                    Width =720
                    FontSize =14
                    TabIndex =4
                    ForeColor =255
                    Name ="btnCancel"
                    Caption ="í ½í·´"
                    OnClick ="[Event Procedure]"
                    FontName ="Academy Engraved LET"
                    GridlineColor =10921638

                    LayoutCachedLeft =8100
                    LayoutCachedTop =2340
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =2700
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =-1
                    BackColor =14136213
                    BorderColor =14136213
                    ThemeFontIndex =-1
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
                    Top =3840
                    Width =8775
                    Height =4380
                    TabIndex =5
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.TaskList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =3840
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =8220
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =3720
                    Width =9000
                    Height =4620
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =3720
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =8340
                    BackThemeColorIndex =-1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =3480
                    Width =9000
                    Height =315
                    FontSize =9
                    LeftMargin =360
                    TopMargin =36
                    RightMargin =360
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblMsg"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedTop =3480
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =3795
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =4320
                    Top =3300
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
                    LayoutCachedTop =3300
                    LayoutCachedWidth =5145
                    LayoutCachedHeight =3900
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =75
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =75
                    LayoutCachedWidth =780
                    LayoutCachedHeight =375
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8430
                    Top =105
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxID"
                    ControlSource ="ID"
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =8430
                    LayoutCachedTop =105
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =405
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =5340
                    Top =600
                    Width =1980
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCharacterCount"
                    Caption ="Character Count:"
                    GridlineColor =10921638
                    LayoutCachedLeft =5340
                    LayoutCachedTop =600
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =840
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =223
                    TextAlign =2
                    Left =6720
                    Top =600
                    Width =660
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =255
                    Name ="lblCount"
                    GridlineColor =10921638
                    LayoutCachedLeft =6720
                    LayoutCachedTop =600
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =840
                    ForeThemeColorIndex =-1
                End
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =7320
                    Top =540
                    Width =1500
                    Height =324
                    BorderColor =10921638
                    Name ="rctAlert"
                    GridlineColor =10921638
                    LayoutCachedLeft =7320
                    LayoutCachedTop =540
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =864
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =7380
                    Top =600
                    Width =1380
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    Name ="lblMaxCount"
                    Caption ="255"
                    GridlineColor =10921638
                    LayoutCachedLeft =7380
                    LayoutCachedTop =600
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =840
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1920
                    Top =2460
                    Width =2274
                    Height =315
                    TabIndex =8
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x01000000ac000000020000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff00010000000000000004000000250000000100000000000000 ,
                        0xfff2000000000000000000000000000000000000000000000000000000000000 ,
                        0x220022000000000049004900660028004c0065006e0028005b00630062007800 ,
                        0x520065007100750065007300740065006400420079005d0029003d0030002c00 ,
                        0x31002c003000290000000000
                    End
                    Name ="cbxRequestedBy"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    ControlTipText ="Person who collected the plant"
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =2460
                    LayoutCachedWidth =4194
                    LayoutCachedHeight =2775
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000fff200002000000049004900660028004c0065006e0028005b ,
                        0x00630062007800520065007100750065007300740065006400420079005d0029 ,
                        0x003d0030002c0031002c00300029000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Top =2460
                    Width =1620
                    Height =315
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblRequestedBy"
                    Caption ="â¯ˆ Requested By"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =2460
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =2775
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Top =2895
                    Width =1620
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblRequestDate"
                    Caption ="Request Date"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =2895
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =3210
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =2895
                    Width =1320
                    Height =300
                    TabIndex =9
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxRequestDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000ec000000020000000100000000000000000000002200000001000000 ,
                        0x00000000ffffff00010000000000000023000000450000000100000000000000 ,
                        0xfff2000000000000000000000000000000000000000000000000000000000000 ,
                        0x490049006600280049007300440061007400650028005b007400620078005200 ,
                        0x65007100750065007300740044006100740065005d0029002c0031002c003000 ,
                        0x290000000000490049006600280049007300440061007400650028005b007400 ,
                        0x62007800520065007100750065007300740044006100740065005d0029002c00 ,
                        0x30002c003100290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =2895
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =3195
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000ffffff00210000004900 ,
                        0x49006600280049007300440061007400650028005b0074006200780052006500 ,
                        0x7100750065007300740044006100740065005d0029002c0031002c0030002900 ,
                        0x0000000000000000000000000000000000000000000100000000000000010000 ,
                        0x0000000000fff200002100000049004900660028004900730044006100740065 ,
                        0x0028005b00740062007800520065007100750065007300740044006100740065 ,
                        0x005d0029002c0030002c00310029000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =540
                    Width =840
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label28"
                    Caption ="Context:"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =540
                    LayoutCachedWidth =960
                    LayoutCachedHeight =855
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =1080
                    Top =600
                    Width =2520
                    Height =300
                    FontSize =9
                    BackColor =65535
                    BorderColor =10921638
                    Name ="lblTaskContext"
                    Caption ="task context"
                    ControlTipText ="Table && record ID referenced by task"
                    GridlineColor =10921638
                    LayoutCachedLeft =1080
                    LayoutCachedTop =600
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =900
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeTint =100.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2880
                    Top =60
                    Width =780
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxType"
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3720
                    Top =60
                    Width =780
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxTypeID"
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =3720
                    LayoutCachedTop =60
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
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
' Form:         Task
' Level:        Framework form
' Version:      1.03
'
' Description:  Task form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
'               BLC - 10/25/2016 - 1.01 - revised to clear header title, use GetContext(), code cleanup
'               BLC - 2/13/2017 - 1.02 - revised to use callingform, code cleanup
'               BLC - 9/29/2017 - 1.03 - revised to handle NULL OpenArgs
' =================================

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

Private m_ContextType As String
Private m_ContextID As Long

Private m_CountLabel As String
Private m_CurrentCount As String
Private m_MaxCount As String
Private m_AlertCount As Integer
Private m_RemainingCount As String

Private m_CountLabelFontColor As Long
Private m_CurrentCountFontColor As Long
Private m_MaxCountFontColor As Long
Private m_RemainingCountFontColor As Long
Private m_AlertBoxBackgroundColor As Long

Private m_CountLabelVisible As Byte
Private m_CurrentCountVisible As Byte
Private m_MaxCountVisible As Byte
Private m_RemainingCountVisible As Byte
Private m_AlertCountVisible As Byte
Private m_AlertBoxVisible As Byte

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

'----- task references (type & ID) -----
Public Property Get ContextType() As String
    ContextType = m_ContextType
End Property

Public Property Let ContextType(Value As String)
    m_ContextType = Value
End Property

Public Property Get ContextID() As Long
    ContextID = m_ContextID
End Property

Public Property Let ContextID(Value As Long)
    m_ContextID = Value
End Property

Public Property Get CurrentCount() As String
    CurrentCount = m_CurrentCount
End Property

Public Property Let CurrentCount(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "1"
    If ValidateString(Value, "numeric") Then
        m_CurrentCount = Value
    End If
    lblCount.Caption = m_CurrentCount
End Property

Public Property Get MaxCount() As String
    MaxCount = m_MaxCount
End Property

Public Property Let MaxCount(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "/ XX characters"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_MaxCount = Value
    End If
    lblMaxCount.Caption = m_MaxCount
End Property

'set the value at which the count display changes color
Public Property Get AlertCount() As Integer
    AlertCount = m_AlertCount
End Property

Public Property Let AlertCount(Value As Integer)
    If Len(Trim(Value)) = 0 Then Value = 10
    m_AlertCount = Value
End Property

Public Property Get RemainingCount() As String
    RemainingCount = m_RemainingCount
End Property

Public Property Let RemainingCount(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "XX characters remain"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_RemainingCount = Value
    End If
    lblMaxCount.Caption = m_RemainingCount
End Property

' ==== Color ====
Public Property Get CountLabelFontColor() As Long
    CountLabelFontColor = m_CountLabelFontColor
End Property

Public Property Let CountLabelFontColor(Value As Long)
    m_CountLabelFontColor = Value
    lblCount.ForeColor = m_CountLabelFontColor
End Property

Public Property Get CurrentCountFontColor() As Long
    CurrentCountFontColor = m_CurrentCountFontColor
End Property

Public Property Let CurrentCountFontColor(Value As Long)
    m_CurrentCountFontColor = Value
    lblCount.ForeColor = m_CurrentCountFontColor
End Property

Public Property Get MaxCountFontColor() As Long
    MaxCountFontColor = m_MaxCountFontColor
End Property

Public Property Let MaxCountFontColor(Value As Long)
    m_MaxCountFontColor = Value
    lblMaxCount.ForeColor = m_MaxCountFontColor
End Property

Public Property Get RemainingCountFontColor() As Long
    RemainingCountFontColor = m_RemainingCountFontColor
End Property

Public Property Let RemainingCountFontColor(Value As Long)
    m_RemainingCountFontColor = Value
    lblMaxCount.ForeColor = m_RemainingCountFontColor
End Property

Public Property Get AlertBoxBackgroundColor() As Long
    AlertBoxBackgroundColor = m_AlertBoxBackgroundColor
End Property

Public Property Let AlertBoxBackgroundColor(Value As Long)
    rctAlert.BackStyle = 1 '1 = Normal, 0 = Transparent
    m_AlertBoxBackgroundColor = Value
    rctAlert.BackColor = m_AlertBoxBackgroundColor
End Property

' ==== Visibility ====
Public Property Get CountLabelVisible() As Byte
    CountLabelVisible = m_CountLabelVisible
End Property

Public Property Let CountLabelVisible(Value As Byte)
    m_CountLabelVisible = Value
    lblCount.Visible = m_CountLabelVisible
End Property

Public Property Get CurrentCountVisible() As Byte
    CurrentCountVisible = m_CurrentCountVisible
End Property

Public Property Let CurrentCountVisible(Value As Byte)
    m_CurrentCountVisible = Value
    lblCount.Visible = m_CurrentCountVisible
    lblCharacterCount.Visible = m_CurrentCountVisible
End Property

Public Property Get MaxCountVisible() As Byte
    MaxCountVisible = m_MaxCountVisible
End Property

Public Property Let MaxCountVisible(Value As Byte)
    m_MaxCountVisible = Value
    lblMaxCount.Visible = m_MaxCountVisible
End Property

Public Property Get RemainingCountVisible() As Byte
    RemainingCountVisible = m_RemainingCountVisible
End Property

Public Property Let RemainingCountVisible(Value As Byte)
    m_RemainingCountVisible = Value
End Property

Public Property Get AlertBoxVisible() As Byte
    AlertBoxVisible = m_AlertBoxVisible
End Property

Public Property Let AlertBoxVisible(Value As Byte)
    m_AlertBoxVisible = Value
    Me.rctAlert.Visible = m_AlertBoxVisible
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
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 10/25/2016 - revised to clear header title, use GetContext(), CallingForm property
'   BLC - 9/29/2017 - revised to handle NULL OpenArgs
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'defaults
    Me.CallingForm = "Main"
    Me.ContextType = ""
    Me.ContextID = 0

    'handle NULL OpenArgs
    If Len(Nz(Me.OpenArgs, "")) > 0 Then
        If CountInString(Me.OpenArgs, "|") = 1 Then '2 Then
            Dim aryContext() As String
            
            aryContext() = Split(Me.OpenArgs, "|")
            Me.CallingForm = aryContext(0)
            
            'set task context
            Me.ContextType = aryContext(0) '(1)
            Me.ContextID = aryContext(1) '(2)
        End If
    End If
    
    'minimize calling form
    ToggleForm Me.CallingForm, -1
        
    'set context - based on TempVars
    lblContext.ForeColor = lngLime
    lblContext.Caption = GetContext()

    'set the task context
    lblTaskContext.Caption = Me.ContextType & " (" & Me.ContextID & ")"
   ' tbxType.Value = Me.ContextType
    'tbxTypeID.Value = Me.ContextID
    tbxType = Me.ContextType
    tbxTypeID = Me.ContextID


    Title = "Task"
    lblTitle.Caption = "" 'clear header title
    Directions = "Enter task details."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
    
    'set hover
    btnCancel.HoverColor = lngGreen
    btnSave.HoverColor = lngGreen
      
    'defaults
    tbxIcon.ForeColor = lngRed
    btnCancel.Enabled = False
    btnSave.Enabled = False
    cbxStatus.BackColor = lngYellow
    cbxPriority.BackColor = lngYellow
    lblMsgIcon.Caption = ""
    lblMsg.Caption = ""
  
    cbxStatus.DefaultValue = 1 'open
    'populate form
    PopulateCombobox cbxPriority, "priority"
    PopulateCombobox cbxStatus, "status"
      
    'populate dropdowns
    Set cbxRequestedBy.Recordset = GetRecords("s_contact_list")
    cbxRequestedBy.BoundColumn = 1
    cbxRequestedBy.ColumnCount = 2
    cbxRequestedBy.ColumnWidths = "0;1"
  
    'set default --> current user
    cbxRequestedBy.Value = TempVars("AppUserID")
    'highlight the requested by label to bring attention
    'in case user isn't the requestor
    lblRequestedBy.ForeColor = lngBlue
    lblRequestedBy.Caption = StringFromCodepoint(uRTriangle) & " Requested By"
  
    'counts
    Me.CountLabelVisible = False
    Me.CurrentCount = "Characters Remaining:"
    Me.lblCharacterCount.Visible = True
    Me.MaxCount = 255
    Me.AlertCount = 10
  
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
            "Error encountered (#" & Err.Number & " - Form_Open[Task form])"
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
'   BLC - 1/13/2017 - remove extraneous code
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Site form])"
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
              
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[Site form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxTask_KeyPress
' Description:  Textbox keypress actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 28, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/28/2016 - initial version
' ---------------------------------
Private Sub tbxTask_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler

    LimitKeyPress Me.tbxTask, 255, KeyAscii
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxTask_KeyPress[Task form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxRequestedBy_GotFocus
' Description:  Combobox got focus actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 14, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/14/2017 - initial version
' ---------------------------------
Private Sub cbxRequestedBy_GotFocus()
On Error GoTo Err_Handler

    lblRequestedBy.ForeColor = lngGray50 'RGB(127,127,127)
    lblRequestedBy.Caption = "Requested By"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxRequested_GotFocus[Site form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxTask_Change
' Description:  Textbox change actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 28, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/28/2016 - initial version
' ---------------------------------
Private Sub tbxTask_Change()
On Error GoTo Err_Handler

'    LimitChange Me.tbxTask, 255
    
    Dim CurrentCount As Integer
    
    CurrentCount = CInt(Me.MaxCount) - Len(tbxTask.Text)

    Me.lblMaxCount.Caption = CurrentCount & " remaining"
    
    Me.CurrentCountFontColor = vbBlack
    Me.AlertBoxVisible = False
    Me.MaxCountFontColor = vbBlack
    
    Select Case CurrentCount
        Case Is < Me.AlertCount
            Me.AlertBoxVisible = True
            Me.AlertBoxBackgroundColor = lngYellow
            Me.MaxCountFontColor = vbRed
        Case Is = 0
            Me.CurrentCountFontColor = vbRed
        Case Else
    End Select


'    If Me.MaxCount - Len(tbxTask.Text) < 10 Then
'        Me.MaxCountFontColor = vbRed
'    Else
'        Me.MaxCountFontColor = vbBlack
'    End If

    If Len(tbxTask.Text) > Me.MaxCount Then
        Me.lblMaxCount.Caption = -(Me.MaxCount - Len(tbxTask.Text)) & " over"
    End If
    
        If CurrentCount < 1 Then 'CInt(Me.MaxCount) Then
        Me.MaxCountFontColor = vbRed
    End If
    
    If Len(tbxTask.Text) > CInt(Me.MaxCount) Then
        Me.lblMaxCount.Caption = -CurrentCount & " over"
        'disable add comment button until count is < or = MaxCount
        Me.btnSave.Enabled = False
    ElseIf Len(tbxTask.Text) = 0 Then
        'disable add comment button if count = 0
        Me.btnSave.Enabled = False
    Else
        're-enable add comment button
        Me.btnSave.Enabled = True
    End If

    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxTask_Change[Task form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxPriority_AfterUpdate
' Description:  Combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 28, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/28/2016 - initial version
' ---------------------------------
Private Sub cbxPriority_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPriority_AfterUpdate[Site form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxStatus_AfterUpdate
' Description:  Combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 28, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/28/2016 - initial version
' ---------------------------------
Private Sub cbxStatus_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxStatus_AfterUpdate[Site form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxTask_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 28, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/28/2016 - initial version
' ---------------------------------
Private Sub tbxTask_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxTask_AfterUpdate[Task form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxRequestedBy_AfterUpdate
' Description:  Combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 14, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/14/2017 - initial version
' ---------------------------------
Private Sub cbxRequestedBy_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxRequested_AfterUpdate[Site form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnCancel_Click
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
Private Sub btnCancel_Click()
On Error GoTo Err_Handler
    
    ClearForm Me
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCancel_Click[Task form])"
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
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    UpsertRecord Me
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[Task form])"
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
'   BLC - 2/13/2017 - revised to use calling form
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore CallingForm
    ToggleForm Me.CallingForm, 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[Task form])"
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
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    'requires: status, task & priority
    If Len(Nz(cbxStatus.Value, "")) > 0 _
        And Len(Nz(cbxPriority.Value, "")) > 0 _
        And Len(Nz(tbxTask.Value, "")) > 0 Then
        isOK = True
    End If
    
    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    btnSave.Enabled = isOK
    
    'refresh form
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[Task form])"
    End Select
    Resume Exit_Handler
End Sub
