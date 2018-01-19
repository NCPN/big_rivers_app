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
    Width =9360
    DatasheetFontHeight =11
    ItemSuffix =88
    Left =4545
    Top =1425
    Right =13905
    Bottom =13410
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x8a05f9ebf1d4e440
    End
    Caption ="Map Import Fields"
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            CanGrow = NotDefault
            Height =900
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =180
                    Top =60
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Map Import Fields"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    Left =180
                    Top =60
                    Width =6840
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Select the table to import to && map the CSV fields at right.\015\012Then import"
                        " the CSV data by clicking the button at right."
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =240
                    Top =480
                    Width =1440
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTable"
                    Caption ="Database Table"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =480
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =795
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =5160
                    Top =60
                    Width =4140
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblContext"
                    Caption ="context"
                    GridlineColor =10921638
                    LayoutCachedLeft =5160
                    LayoutCachedTop =60
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListRows =10
                    Left =1740
                    Top =480
                    Width =2964
                    Height =315
                    ColumnOrder =1
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxTable"
                    ColumnWidths ="1440"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Table to import into"
                    GridlineColor =10921638
                    SeparatorCharacters =2
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =1740
                    LayoutCachedTop =480
                    LayoutCachedWidth =4704
                    LayoutCachedHeight =795
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =480
                    Width =180
                    Height =300
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =4
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =240
                    LayoutCachedHeight =780
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7740
                    Top =420
                    Width =720
                    TabIndex =6
                    ForeColor =4210752
                    Name ="btnSave"
                    Caption ="Edit"
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

                    LayoutCachedLeft =7740
                    LayoutCachedTop =420
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =780
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
                    Left =4800
                    Top =480
                    Width =240
                    Height =300
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =5
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxID"
                    DefaultValue ="0"
                    ConditionalFormat = Begin
                        0x01000000ae000000020000000100000000000000000000001200000001000000 ,
                        0x92929200ffffff0001000000000000001300000026000000010000003f3f3f00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00540072007500 ,
                        0x6500000000005b007400620078004400650076004d006f00640065005d003d00 ,
                        0x460061006c007300650000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4800
                    LayoutCachedTop =480
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =780
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000092929200ffffff00110000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d005400720075006500 ,
                        0x0000000000000000000000000000000000000000000100000000000000010000 ,
                        0x003f3f3f00ffffff00120000005b007400620078004400650076004d006f0064 ,
                        0x0065005d003d00460061006c0073006500000000000000000000000000000000 ,
                        0x000000000000
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8520
                    Top =420
                    Width =720
                    TabIndex =3
                    ForeColor =16711680
                    Name ="btnComment"
                    Caption ="í ½í·©"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =8520
                    LayoutCachedTop =420
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =780
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
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5220
                    Top =420
                    Width =720
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnNewTable"
                    Caption ="Make new table"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Create a new table to import data to"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b0a090ff604830ff705040ff605040ff605040ff604830ff ,
                        0x604830ff604830ff705040ff0000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a090fffff0f0fffff0f0fffff0e0ffffe8e0ffffe8d0ff ,
                        0xffe0d0fff0e0d0ff705040ff00000000000000000000000020182080000000ff ,
                        0x0000007000000000c0a8a0fffff8f0ffd0c0b0ffd0b8a0ffffe8e0ffd0b0a0ff ,
                        0xd0b0a0ffffe0d0ff705040ff000000000000000000000000808080ff606060ff ,
                        0x000000ff00000000c0b0a0fffff8f0fffff8f0fffff0f0fffff0e0ffffe8e0ff ,
                        0xffe8e0ffffe0d0ff705040ff00000000000000000000000070707080807880ff ,
                        0x1010105000000000c0b8a0fffff8ffffe0d0d0ffe0c8c0fffff0f0ffd0b0a0ff ,
                        0xd0b0a0ffffe8e0ff705040ff0000000000000000000000000000000010081070 ,
                        0x0000000000000000d0b8b0fffffffffffff8f0fffff8f0fffff0f0fffff0f0ff ,
                        0xfff0e0ffffe8e0ff705040ff00000000000000000000000000000000202020ff ,
                        0x0000000000000000d0b8b0ffffffffffe0d8d0ffe0d0d0fffff8f0ffd0c0b0ff ,
                        0xd0c0b0ffffe8e0ff705040ff30b8e00030b8e0000000000000000070404040ff ,
                        0x0000007000000000d0b8b0fffffffffffffffffffff8ffff80d0e0ff90e0f0ff ,
                        0x90e8ffff40c0e0ff90e0e0ffa0e0f0ffc0d8d0ff00000000202820ff404040ff ,
                        0x000000f000000000f0a890fff0a880fff0a080ffe09870ffa0b0a0ff30b8e0ff ,
                        0x80e8ffff60c8e0ff90f0ffff30b8e0ff90c0c0ff40404070606060ff505850ff ,
                        0x202020ff00000030f0a890ffffc0a0ffffc0a0ffffb890ffa0d8e0ff90f0ffff ,
                        0xc0f8ffffb0e8f0ffc0f8ffff90f0ffffa0e0e0ff404040c0707070ff606060ff ,
                        0x504850ff00000080f0b090fff0a890fff0a080fff09870ff20a8d0ff50c0e0ff ,
                        0xb0e8f0fff0ffffffb0e8f0ff50c0e0ff30b8e0ff605860ff909890ff606060ff ,
                        0x505050ff000000f00000000000000000000000000000000080e8ffc090f0ffff ,
                        0xc0f8ffffb0e8f0ffc0f8ffff90f0ffff80e8ffc0707070ffa0a8a0ff707070ff ,
                        0x606060ff101010f00000000000000000000000000000000050d8ff8030b8e0ff ,
                        0x90f0ffff60c0e0ff90f0ffff30b8e0ff50d0f080807880ffc0b8c0ffb0b0b0ff ,
                        0x908890ff100810e00000000000000000000000000000000030b0e0a040c8f090 ,
                        0x80e8ffc020b0e0ff70e8ffc050d8f08030b0e08070707030807880ff606860ff ,
                        0x505050ff50505020000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5220
                    LayoutCachedTop =420
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =780
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
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6060
                    Top =420
                    Width =720
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnCopyTable"
                    Caption ="Make new table"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Copy the currently selected table for import"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a890ff908070ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff000000000000000000000000 ,
                        0x0000000000000000c0a8a0fffffffffff0e0d0fff0d8d0fff0d8d0fff0d8d0ff ,
                        0xf0d8d0fff0d8d0fff0d0c0fff0d0c0ff604830ff000000000000000000000000 ,
                        0x0000000000000000c0b0a0ffffffffffd0b8b0ffc0b8a0ffffffffffffffffff ,
                        0xffffffffc0a8a0ffc0a890fff0d0c0ff604830ff000000000000000000000000 ,
                        0x0000000000000000c0b0a0fffffffffffffffffffff8f0fffff8f0fffff8f0ff ,
                        0xfff8f0fffff8f0fffff8f0fff0d8c0ff604830ff000000000000000000000000 ,
                        0x0000000000000000d0b8a0ffffffffffd0c0b0ffd0b8b0fffff8f0ffc0a890ff ,
                        0x908070ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ffd0b8b0ffffffffffffffffffffffffffffffffffc0a8a0ff ,
                        0xfffffffffffffffffffffffff0d8d0ff604830fff0d8d0fff0d8d0fff0d0c0ff ,
                        0xf0d0c0ff604830ffd0c0b0ffffffffffd0c0b0ffd0c0b0ffffffffffc0b0a0ff ,
                        0xffffffffc0b0a0ffc0b0a0fff0d8d0ff705040ffffffffffc0a8a0ffc0a890ff ,
                        0xf0d0c0ff604830ffd0c0b0ffffffffffffffffffffffffffffffffffc0b0a0ff ,
                        0xfffffffffffffffffffffffffff8f0ff806050fffff8f0fffff8f0fffff8f0ff ,
                        0xf0d8c0ff604830ffa09860ffa09050ffa09050ff908850ff908840ffd0b8a0ff ,
                        0x908030ff807830ff807830ff807020ff807020ffffffffffc0b0a0ffc0a8a0ff ,
                        0xf0d8d0ff604830ffa09860ffd0d0a0ffd0c890ffd0c090ffc0b880ffd0b8b0ff ,
                        0xb0a870ffb0a060ffa09860ffa09050ff807020ffffffffffffffffffffffffff ,
                        0xf0d8d0ff604830ffa0a070ffa09860ffa09860ffa09860ffa09050ffd0c0b0ff ,
                        0x908850ff908850ff908840ff908040ff908040ffffffffffc0b0a0ffc0b0a0ff ,
                        0xf0d0c0ff705040ff0000000000000000000000000000000000000000d0c0b0ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfff8f0ff806050ff0000000000000000000000000000000000000000a09860ff ,
                        0xa09050ffa09050ff908850ff908840ff908040ff908030ff807830ff807830ff ,
                        0x807020ff807020ff0000000000000000000000000000000000000000a09860ff ,
                        0xd0d0a0ffd0c890ffd0c090ffc0b880ffc0b070ffb0a870ffb0a060ffa09860ff ,
                        0xa09050ff807020ff0000000000000000000000000000000000000000a0a070ff ,
                        0xa09860ffa09860ffa09860ffa09050ffa09050ff908850ff908850ff908840ff ,
                        0x908040ff908040ff
                    End

                    LayoutCachedLeft =6060
                    LayoutCachedTop =420
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =780
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
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6960
                    Top =480
                    Width =540
                    Height =315
                    ColumnOrder =0
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxDevMode"
                    ConditionalFormat = Begin
                        0x0100000084000000010000000100000000000000000000001100000001000000 ,
                        0xececec00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004400450056005f004d004f00440045005d003d00460061006c0073006500 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =6960
                    LayoutCachedTop =480
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =795
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ffffff00100000005b00 ,
                        0x4400450056005f004d004f00440045005d003d00460061006c00730065000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =11100
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =45
                    Top =5640
                    Width =9255
                    Height =5340
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.CSVDataList"
                    GridlineColor =10921638

                    LayoutCachedLeft =45
                    LayoutCachedTop =5640
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =10980
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =5520
                    Width =9360
                    Height =5580
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =5520
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =11100
                    BackThemeColorIndex =-1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =5280
                    Width =9360
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
                    LayoutCachedTop =5280
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =5595
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =4320
                    Top =5100
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
                    LayoutCachedTop =5100
                    LayoutCachedWidth =5145
                    LayoutCachedHeight =5700
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Subform
                    CanShrink = NotDefault
                    OverlapFlags =95
                    Left =300
                    Top =510
                    Width =3600
                    Height =4590
                    TabIndex =1
                    BorderColor =10921638
                    Name ="listTableFields"
                    SourceObject ="Form.TableFieldList"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =510
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =5100
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =300
                    Width =1260
                    Height =180
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =255
                    Name ="lblHintReqd"
                    Caption ="* = Required Field"
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =300
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =2640
                    Top =330
                    Width =1260
                    Height =180
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblHintZLS"
                    Caption ="Blue = Allows ZLS"
                    GridlineColor =10921638
                    LayoutCachedLeft =2640
                    LayoutCachedTop =330
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Subform
                    CanShrink = NotDefault
                    OverlapFlags =85
                    Left =5580
                    Top =540
                    Width =3600
                    Height =4590
                    TabIndex =2
                    BorderColor =10921638
                    Name ="listCSV"
                    SourceObject ="Form.ImportColumnList"
                    GridlineColor =10921638

                    LayoutCachedLeft =5580
                    LayoutCachedTop =540
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =5130
                End
                Begin Label
                    OverlapFlags =85
                    Left =5400
                    Top =60
                    Width =3719
                    Height =180
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="lblHintSelected"
                    Caption ="Green = Import column to the field at left"
                    GridlineColor =10921638
                    LayoutCachedLeft =5400
                    LayoutCachedTop =60
                    LayoutCachedWidth =9119
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =5400
                    Top =300
                    Width =3720
                    Height =180
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =255
                    Name ="lblHintNone"
                    Caption ="None = Set table column values to NULL on import"
                    GridlineColor =10921638
                    LayoutCachedLeft =5400
                    LayoutCachedTop =300
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =4380
                    Top =1980
                    Width =720
                    FontSize =20
                    TabIndex =3
                    ForeColor =255
                    Name ="btnImport"
                    Caption ="â—€"
                    StatusBarText ="Import CSV data to table"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Import CSV data to table"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b0482000b0482050000000000000000000000000 ,
                        0x0000000000000000000000004068ff0000000000000000000000000000000000 ,
                        0x000000000000000000000000b0502050904820ff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000a0482040d06830ff905030ff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000b0502040d06030fff06820ffa05030ff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xb0502050d06830fff07030fff06820ffa05830ff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b0502050 ,
                        0xe07040ffffa060fff08850fff07030ffb06040ff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000d0704040e08850ff ,
                        0xffc0a0ffffb090ffffa070ffff8040ffb06840ff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000e0906000e0906040 ,
                        0xe08850ffffc0a0ffffb080ffff8850ffc07040ff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e0906000 ,
                        0xe0906040e08860ffffc0a0ffff9870ffd07850ff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xe0906000e0906040e08860ffffc0a0ffd07850ff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000e0906000f0906030e08860ffd08050ff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000e0906000f0906020e08850ff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000e0906000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4380
                    LayoutCachedTop =1980
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =2340
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
                    Left =360
                    Top =60
                    Width =4080
                    Height =180
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =255
                    Name ="lblHintIDField"
                    Caption ="í ½í»‡ = Autogenerated ID field. CSV field should be 'None'"
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    Left =4080
                    Top =2520
                    Width =1380
                    Height =840
                    FontSize =8
                    BackColor =62207
                    BorderColor =8355711
                    ForeColor =255
                    Name ="lblImportAlert"
                    Caption ="CSV table is empty, please import new CSV by clicking the button below."
                    ControlTipText ="CSV table is empty, please import new CSV."
                    GridlineColor =10921638
                    LayoutCachedLeft =4080
                    LayoutCachedTop =2520
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =3360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =2
                    Left =4380
                    Top =3480
                    Width =720
                    FontSize =14
                    TabIndex =4
                    ForeColor =255
                    Name ="btnImportCSVData"
                    Caption ="import"
                    StatusBarText ="Import CSV data to usys_temp_csv"
                    OnClick ="[Event Procedure]"
                    FontName ="Academy Engraved LET"
                    ControlTipText ="Import CSV data to usys_temp_csv"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000707880ff608090ff607880ff507080ff506070ff405860ff ,
                        0x405060ff404850ff404850ff404040ff303030ff202020ff1018208000000000 ,
                        0x0000000000000000708090ff80a0a0ff70b0d0ff0090d0ff0090d0ff1090d0ff ,
                        0x30a0d0ff50a8d0ff80b8d0ff80b8d0ff70a8c0ff5090b0ff305860ff30384050 ,
                        0x0000000000000000708090ff80c0d0ff80a0b0ff80e0ffff60d0ffff60d0ffff ,
                        0x70d0ffff509860ff308040ff206030ff90b8a0ff80c0e0ff5088a0ff303840c0 ,
                        0xfff8f00000000000708890ff80d0f0ff80a0b0ff80c0d0ff70d8ffff70d8ffff ,
                        0x80d8ffffb0e0ffff308040ff60a870ff206830ff80a890ff70b0e0ff406070ff ,
                        0x2038403000000000708890ff80d8f0ff80c8e0ff80a0b0ff80e0ffff70d8ffff ,
                        0x80d8ffffa0e0ffffd0f0ffff308040ff60a870ff206030ffa0d8f0ff5088a0ff ,
                        0x30587090fff8f000808890ff90e0f0ff90e0ffff90a0b0ff90b8c0ff80d8ffff ,
                        0x80d8ffffb0e8ffffe0f0ffff308040ff80d8a0ff206030ffd0e8f0ff80c8e0ff ,
                        0x707880f0705040608090a0ff90e0f0ffa0e8ffff80c0e0ff90a0b0ff90e0ffff ,
                        0xb0e8ffff308050ff308040ff60a870ff80d8a0ff308040ff206830ff307040ff ,
                        0x90c0e0ff706860d08090a0ffa0e8f0ffa0e8ffffa0e8ffff80a8b0ff90a8b0ff ,
                        0xa0b8c0ffb0c0b0ff308050ff70c080ff80d8a0ff50a060ff408050ffb0c0b0ff ,
                        0xa0a8b0ff8090a0ff8098a0ffa0e8f0ffa0f0ffffa0e8ffffa0e8ffff80d8ffff ,
                        0xc0b0a0fffff8f0ffd0e0d0ff408050ff60a870ff408050ffc0d0c0fffff8f0ff ,
                        0xffe8e0ff705040ff8098a0ffa0f0f0ffb0f0f0ffa0f0ffffa0e8ffffa0e8ffff ,
                        0xc0a8a0ffd0c0b0ffe0d0c0ffc0c8c0ff408050ffc0c8c0ffe0c8c0ffd0b8b0ff ,
                        0xc0b0a0ff604830ff8098a0ffa0d0e0ffb0f0f0ffb0f0f0ffa0f0ffffa0e8ffff ,
                        0xb0a8a0fffffffffffff8ffffd0c0c0fffff8f0fffff0e0ffd0b8b0fffff8f0ff ,
                        0xffe8e0ff604830ff8098a0508098a0ff8098a0ff8098a0ff8098a0ff8098a0ff ,
                        0xb0a8a0ffc0b0a0ffc0b8a0ffc0b0a0ffc0b0a0ffc0b0a0ffc0b0a0ffc0b0a0ff ,
                        0xc0b0a0ff604830ff000000000000000000000000000000000000000000000000 ,
                        0xb0a8a0ffffffffffffffffffc0b0a0fffff8fffffff0f0ffc0b0a0fffff8f0ff ,
                        0xfff0f0ff604830ff000000000000000000000000000000000000000000000000 ,
                        0xb09080ffb08060ffb08060ffb08060ffb07860ffb07860ffb07860ffb07860ff ,
                        0xb08060ffb08060ff000000000000000000000000000000000000000000000000 ,
                        0xb08870ffe0c8b0ffe0c0b0ffb08060ffe0c0b0ffe0c0b0ffb07860ffe0b8b0ff ,
                        0xe0b8b0ffb08060ff000000000000000000000000000000000000000000000000 ,
                        0xb08870ffc09080ffc09070ffb08870ffb08070ffb08060ffb08060ffb08060ff ,
                        0xb08060ffb08060ff
                    End

                    LayoutCachedLeft =4380
                    LayoutCachedTop =3480
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =3840
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
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4440
                    Top =420
                    Width =480
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCSVRecord"
                    ControlSource ="=[Forms]![ImportColumnList].[CurrentRecord]"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4440
                    LayoutCachedTop =420
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =735
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =247
                    Top =765
                    Width =720
                    Height =4320
                    TabIndex =6
                    BorderColor =10921638
                    Name ="overlay"
                    SourceObject ="Form.TableFieldListOverlay"
                    GridlineColor =10921638

                    LayoutCachedTop =765
                    LayoutCachedWidth =720
                    LayoutCachedHeight =5085
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
' Form:         ImportMap
' Level:        Application form
' Version:      1.10
' Basis:        Dropdown form
'
' Description:  ImportMap form object related properties, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 10/18/2016 - 1.00 - initial version
'               BLC - 10/19/2016 - 1.01 - code cleanup, added callingform property
'               BLC - 10/20/2016 - 1.02 - adjusted to use GetContext(), remove btnSave, ReadyForSave,
'                                         code cleanup
'               BLC - 10/27/2016 - 1.03 - revised to refresh data list after import
'               BLC - 12/1/2016 - 1.04 - added btnComment & click event
'               BLC - 12/5/2016 - 1.05 - revised comment click event to pass imported data ID & max length
'               BLC - 12/8/2016 - 1.06 - revise to make comment button invisible, require CSV import to start
'               BLC - 12/13/2016 - 1.07 - added row highlighting (current CSV record drives highlighting of table fields)
'               BLC - 1/3/2017  - 1.08  - btnImportCSV_Click code cleanup, enabled XLS export button
'                                         when table is specified (CSV data list is populated)
'               BLC - 1/9/2018  - 1.09  - added new table creation button (btnNewTable)
'               BLC - 1/11/2018 - 1.10  - added copy table button (btnCopyTable)
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
Private m_SelectedTable As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)
Public Event InvalidSelectedTable(Value As String)

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

Public Property Let SelectedTable(Value As String)
    If Len(Value) > 0 Then
        m_SelectedTable = Value
    Else
        RaiseEvent InvalidSelectedTable(Value)
    End If
End Property

Public Property Get SelectedTable() As String
    SelectedTable = m_SelectedTable
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
' Source/date:  Bonnie Campbell, October 6, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/6/2016 - initial version
'   BLC - 10/19/2016 - adjusted to use callingform property
'   BLC - 10/20/2016 - adjusted to use GetContext(), revised ListTables()
'   BLC - 12/8/2016 - revised to make comment invisible, require CSV import to start
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "Main"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'restore calling form
    ToggleForm Me.CallingForm, -1

    'set dev mode
    Me.tbxDevMode.Value = DEV_MODE

    'set context - based on TempVars
    lblContext.ForeColor = lngLime
    'lblContext.Caption = GetContext()
                 
    Title = "Map Import Fields"
    Me.lblTitle.Visible = False
    Me.lblContext.Visible = False
    Directions = "Select the table to import to && map the CSV fields at right." _
                & vbCrLf & "Import the CSV data by clicking the left arrow button."
    lblDirections.ForeColor = lngLtBlue
    lblDirections.Caption = Directions
    tbxIcon.Value = StringFromCodepoint(uBullet)
    btnImport.Caption = StringFromCodepoint(uTriangleBlkL)
    btnImport.ForeColor = lngRed
    
    'new CSV import
    lblImportAlert.Visible = False
    lblImportAlert.TextAlign = taCenter
    lblImportAlert.Caption = "Importing a new CSV?" & vbCrLf & "Use the button below."
    lblImportAlert.ForeColor = lngRed
    lblImportAlert.BackColor = lngYellow
    btnImportCSVData.Visible = False
    
    'comment, save --> not used
    btnSave.Visible = False
    btnSave.Enabled = False
    btnComment.Visible = False
    btnComment.Enabled = False
    
    'disable import until comment complete
    btnImport.Enabled = False
    
    'set data sources
    PopulateTables
'    Dim strTables As String
'    Dim showsys As Boolean
'
'    'default --> show linked; exclude msys, tsys, usys tables
'    showsys = False
'
'    'include all except msys for administrators
'    If TempVars("UserAccessLevel") = "admin" Then showsys = True
'
'    strTables = ListTables(False, showsys, showsys, True)
'Debug.Print "ImportMap form_open strTables = " & strTables
'
'    cbxTable.SeparatorCharacters = acSeparatorCharactersSemiColon
'    cbxTable.RowSourceType = "Value List"
'    cbxTable.RowSource = Replace(strTables, "|", ";")
    
    'hints
    lblHintReqd.Caption = "* = Required Field"
    lblHintReqd.ForeColor = lngRed
    lblHintReqd.Visible = False
    lblHintZLS.Caption = "Blue = Allows ZLS"
    lblHintZLS.ForeColor = lngBlue
    lblHintZLS.Visible = False
    lblHintSelected.Caption = "Green = Import column to the field at left"
    lblHintSelected.ForeColor = lngDkGreen
    lblHintSelected.Visible = False
    lblHintNone.Caption = "None = Set table column values to NULL on import"
    lblHintNone.ForeColor = lngRed
    lblHintNone.Visible = False
    lblHintIDField.ForeColor = lngRed
    lblHintIDField.Visible = False
    lblHintIDField.Caption = StringFromCodepoint(uProhibited) & " = Autogenerated ID field. CSV field should be 'None'"
    
    'set fore colors
    btnNewTable.ForeColor = lngBlue
    btnCopyTable.ForeColor = lngBlue
    btnComment.ForeColor = lngBlue
    btnSave.ForeColor = lngBlue
    
    'set hover
    btnSave.HoverColor = lngGreen
    btnImport.HoverColor = lngGreen
    btnNewTable.HoverColor = lngGreen
    btnCopyTable.HoverColor = lngGreen
      
    'defaults
    tbxIcon.ForeColor = lngRed
    btnSave.Enabled = False
    btnSave.Visible = False
    btnImport.Enabled = False
    btnNewTable.Enabled = True
    btnCopyTable.Enabled = False
    cbxTable.BackColor = lngYellow
    
    'ID default -> value used only for edits of existing table values
    tbxID.Value = 0
  
    'defaults --> turn off items
    btnImport.Visible = False
    listTableFields.Visible = False
    listCSV.Visible = False
     
    'ID default -> value used only for edits of existing table values
    tbxID.DefaultValue = 0
    
    'hide control tracker
    tbxCSVRecord.Visible = False
    
    'initialize values
    ClearForm Me
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[ImportMap form])"
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
            "Error encountered (#" & Err.Number & " - Form_Load[ImportMap form])"
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
'   BLC - 1/11/2018 - clear & update msg & msg icon
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    'pseudo current record for CSV columns (record highlighting)
'    tbxCSVRecord.Value = Replace(Me.listCSV.Form.ActiveControl.Name, "cbxColumnName", "")
    'Forms![ImportColumnList].CurrentRecord '[listCSV].[Form].[CurrentRecord]

    'clear msg & icon
'    lblMsg.ForeColor = lngRobinEgg
'    lblMsgIcon.ForeColor = lngRobinEgg
'    lblMsg.Caption = ""
'    lblMsgIcon.Caption = ""
    ClearMsgIcon Me


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[ImportMap form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxTable_AfterUpdate
' Description:  Combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 6, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/6/2016 - initial version
'   BLC - 12/5/2016 - enabled comment button
'   BLC - 12/8/2016 - disabled import button until CSV fields selected
'   BLC - 1/3/2017  - enable XLS export button when table selected
' ---------------------------------
Private Sub cbxTable_AfterUpdate()
On Error GoTo Err_Handler
    
    If Len(cbxTable.Text) > 0 Then
        
        'set selected table property
        Me.SelectedTable = cbxTable.Text
        
        'unhide & populate controls
        lblHintReqd.Visible = True
        lblHintZLS.Visible = True
        lblHintIDField.Visible = True
        listTableFields.Visible = True
        lblHintSelected.Visible = True
        lblHintNone.Visible = True
        listCSV.Visible = True
        btnImport.Visible = True
        btnNewTable.Visible = True
        btnCopyTable.Visible = True
        
        'enable table copy
        btnCopyTable.Enabled = True
        
        'disable import until fields are selected
        btnImport.Enabled = False
        
        'new CSV import
        lblImportAlert.Visible = True
        btnImportCSVData.Visible = True
        
        listTableFields.Form.Table = cbxTable.Text
        
'        'hide CSV form controls to initialize
'        listCSV.Form.HideControls
'
'        'set recordset for # of dropdowns
'        listCSV.Form.NumColumns = Me.listTableFields.Form.Recordset.RecordCount
'        listCSV.Form.Table = cbxTable.Text
'
'        'disable import on any table ID field columns
'        Debug.Print listTableFields.Form.Controls("tbxFieldName")
'
'        If listTableFields.Form.Controls("tbxFieldName") = "ID" Then
'
'            With listCSV.Form.Controls("cbxColumnName1")
'                .Value = "None"
'                .Enabled = False
'            End With
'
'        End If
        
        SetCSVFieldsDisplay
        
        'display table data - IF view is set to table
        'Me.list.Form.DataList.Form.RecordSource = "SELECT * FROM " & cbxTable.Text & ";" 'SourceObject
        If Me!list.Form!optgView = 1 Then
            Me!list.Form!DataList.SourceObject = "Table." & cbxTable.Text
        End If
        
        'ReadyForSave
    
        'ready to comment
        btnComment.Enabled = True
        
        'ready for XLS export
        Me!list.Form!btnExportXLS.Enabled = True
        
        'clear msg & msgIcon
        ClearMsgIcon Me
        
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTable_AfterUpdate[ImportMap form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxCSVRecord_Change
' Description:  textbox value change actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/13/2016 - initial version
'   BLC - 1/17/2017  - highlight overlay control
' ---------------------------------
Public Sub tbxCSVRecord_Change()
On Error GoTo Err_Handler
    
'MsgBox Me.Parent.Form.Controls("TableFieldList").Form.Controls("tbxHighlight").Name
    Me.listTableFields.Form.Controls("tbxHighlight") = CStr(Me.tbxCSVRecord)
    
    'clear/hide all
    Dim ctrl As Control
    
    For Each ctrl In Me.overlay.Form.Controls
    
        If InStr(ctrl.Name, "lblColumnName") Then
        
            ctrl.Visible = False
            ctrl.BackStyle = 0 'transparent
        
        End If
    
    Next
    
    'highlight
    Dim strControl As String
    
    strControl = "lblColumnName" & CStr(Me.tbxCSVRecord)
    
    Debug.Print "tbxCSVRecord_Change ctrl = " & strControl
    
    Me.overlay.Form.Controls(strControl).BackColor = lngYelLime
    Me.overlay.Form.Controls(strControl).ForeColor = lngBlue
    Me.overlay.Form.Controls(strControl).Caption = StringFromCodepoint(uRArrow)
    'Me.overlay.Form.Requery
    Me.overlay.Form.Controls(strControl).Visible = True
    Me.overlay.Form.Controls(strControl).BackStyle = 1
    
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxCSVRecord_Change[ImportMap form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnNewTable_Click
' Description:  New table creation button click actions
' Assumptions:  Assumes individual creating new tables understands
'               database schema and table construction
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   BlackHawk, January 10, 2018
'   https://stackoverflow.com/questions/48179277/how-to-trigger-the-access-ribbons-create-table-design-button-via-vba/48195078#48195078
' Source/date:  Bonnie Campbell, January 9, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/9/2018 - initial version
' ---------------------------------
Private Sub btnNewTable_Click()
On Error GoTo Err_Handler
    
    'minimize form
    ToggleForm Me.Name, -1
    
    'open new table design view
    DoCmd.RunCommand acCmdNewObjectDesignTable
    
    PopulateTables
    cbxTable.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNewTable_Click[ImportMap form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnCopyTable_Click
' Description:  Copies table button click actions
' Assumptions:  Assumes individual creating new tables understands
'               database schema and table construction
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 11, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/11/2018 - initial version
' ---------------------------------
Private Sub btnCopyTable_Click()
On Error GoTo Err_Handler
        
    'copy currently selected table if it exists (it should)
    'new table name:  OldTableName_YYYYMMDD_HHMM
    'copy structure only: don't use "True" as this still copies data
    If TableExists(Me.cbxTable) Then
    
        Dim tblName As String
        tblName = Me.cbxTable & "_" & Format(Now, "YYYYMMDD_HHMM")
        DoCmd.TransferDatabase acExport, "Microsoft Access", CurrDb().Name, _
            acTable, Me.cbxTable, tblName, 1

        'add new table to tabledefs
        CurrDb.TableDefs.Refresh
        
        lblMsg.ForeColor = lngYellow
        lblMsgIcon.ForeColor = lngYellow
        lblMsg.Caption = tblName & " created!"
        lblMsgIcon.Caption = StringFromCodepoint(uRTriangle) & _
                                        StringFromCodepoint(uRTriangle)
 Debug.Print tblName
 
        'update cbx
        PopulateTables
        cbxTable.Requery
 
 Debug.Print tblName
        
        'add new table to IMPORTS custom group
        SetNavGroup "IMPORTS", tblName, "table"
    
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCopyTable_Click[ImportMap form])"
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
' Source/date:  Bonnie Campbell, December 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/1/2016 - initial version
'   BLC - 12/5/2016 - add imported data ID & max length (255)
' ---------------------------------
Private Sub btnComment_Click()
On Error GoTo Err_Handler
    
    'determine next import ID --> DMax("ID","ImportedData")
    
    'open comment form
    DoCmd.OpenForm "Comment", acNormal, , , , , "ImportedData|" & Nz(DMax("ID", "ImportedData"), 0) + 1 & "|255"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[ImportMap form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnImport_Click
' Description:  Import button click actions
' Assumptions:
'               Assumes that the first ID imported (StartImportID) is the current max record ID + 1
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/18/2016 - initial version
'   BLC - 10/27/2016 - revised to update datalist & complete import
'   BLC - 12/1/2016 - add import logging via ImportedData
' ---------------------------------
Private Sub btnImport_Click()
On Error GoTo Err_Handler

    'determine next import ID --> DMax("ID","ImportedData")
    
    'open comment form
    DoCmd.OpenForm "Comment", acNormal, , , , , "ImportedData|" & Nz(DMax("ID", "ImportedData"), 0) + 1 & "|255"

    'determine initial imported record ID
    Dim StartImportID As Long, EndImportID As Long
    Dim StartCount As Integer, EndCount As Integer, ImportCount As Integer
    
    StartImportID = DMax("ID", Me.listTableFields.Form.Table) + 1
    StartCount = DCount("ID", Me.listTableFields.Form.Table)

    Dim strTableFields As String, strImportColumns As String
    
    strTableFields = Me.listTableFields.Form.TableColumns
    strImportColumns = Me.listCSV.Form.ImportColumns

    'compare the table vs. CSV field lists
    If CountInString(strTableFields, ",") <> _
            CountInString(strImportColumns, ",") Then GoTo Exit_Handler
    
    'remove starting ID & starting NULL
    If Left(strTableFields, 3) = "ID," Then
    
        strTableFields = Trim(Right(strTableFields, Len(strTableFields) - 3))
        strImportColumns = Trim(Right(strImportColumns, Len(strImportColumns) - 5))
    End If
    
    Dim strSQL As String
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    'prepare the SQL
'    strSQL = "INSERT INTO " & Me.listTableFields.Form.Table & "(" & _
'                strTableFields & _
'                ") VALUES (" & _
'                strImportColumns & _
'                ");"
    strSQL = "INSERT INTO " & Me.listTableFields.Form.Table & "(" & _
            strTableFields & _
            ") SELECT " & _
            strImportColumns & _
            " FROM usys_temp_csv;"
    
    Debug.Print strSQL

    'refresh data display
    'Form_CSVDataList.Requery
    'Me.list.Requery
    'switch to CSV view to avoid old data display
'    Form_CSVDataList.optgView.Value = 2
'    Call Form_CSVDataList.RefreshDataList

    'import!
'    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
'    DoCmd.SetWarnings True
    
    'retrieve end & import counts
    EndImportID = DMax("ID", Me.listTableFields.Form.Table)
    EndCount = DCount("ID", Me.listTableFields.Form.Table)
    
    ImportCount = EndCount - StartCount
    
    'record import info
    Dim Params(0 To 4) As Variant
    Dim sfile As String
    
    Params(0) = sfile
    Params(1) = Me.listTableFields.Form.Table
    Params(2) = ImportCount
    Params(3) = StartImportID
    Params(4) = EndImportID
    
    SetRecord "i_imported_data", Params
    
    'reset CSV comboboxes to remove user's list additions <--- FIX!!
    Call Form_ImportColumnList.RefreshColumnList
    
    'switch to CSV view to avoid old data display
    Form_CSVDataList.optgView.Value = 1
    Call Form_CSVDataList.RefreshDataList
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnImport_Click[ImportMap form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnImportCSVData_Click
' Description:  Enter button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   EraserveAP, July 21, 2008
'   Galaxiom, November 21, 2012
'   http://www.access-programmers.co.uk/forums/showthread.php?t=153447
' Source/date:  Bonnie Campbell, December 8, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2016 - initial version
'   BLC - 1/3/2017  - code cleanup
' ---------------------------------
Private Sub btnImportCSVData_Click()
On Error GoTo Err_Handler

    'call click event --> assumes CSVDataList is OPEN (it is as a subform)
    Form_CSVDataList.btnImportCSVData_Click
    
    SetCSVFieldsDisplay

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnImportCSVData_Click[ImportMap form])"
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
'   BLC - 6/27/2016 - revised to use ToggleForm()
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
            "Error encountered (#" & Err.Number & " - Form_Close[ImportMap form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          PopulateTables
' Description:  Populates the table combobox (cbxTables)
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 11, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/11/2018 - initial version
' ---------------------------------
Private Sub PopulateTables()
On Error GoTo Err_Handler

    'set data sources
    Dim strTables As String
    Dim showsys As Boolean
    
    'default --> show linked; exclude msys, tsys, usys tables
    showsys = False

    'include all except msys for administrators
    If TempVars("UserAccessLevel") = "admin" Then showsys = True
    
    strTables = ListTables(False, showsys, showsys, True)
Debug.Print "ImportMap form_open strTables = " & strTables

    cbxTable.SeparatorCharacters = acSeparatorCharactersSemiColon
    cbxTable.RowSourceType = "Value List"
    cbxTable.RowSource = Replace(strTables, "|", ";")
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateTables[ImportMap form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          SetCSVFieldsDisplay
' Description:  CSV field list display actions
' Assumptions:  Public to allow
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 6, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2016 - initial version
' ---------------------------------
Public Sub SetCSVFieldsDisplay()
On Error GoTo Err_Handler
    
    'hide CSV form controls to initialize
    listCSV.Form.HideControls
            
    'set recordset for # of dropdowns
    listCSV.Form.NumColumns = Me.listTableFields.Form.Recordset.RecordCount
    listCSV.Form.Table = Me.cbxTable ' cbxTable.Text --> error #2185: can't reference a property or
                                     ' method for a control unless the control has the focus
    
    'disable import on any table ID field columns
    Debug.Print "ImportMap form SetCSVFieldsDisplay = " & listTableFields.Form.Controls("tbxFieldName")
    
    If listTableFields.Form.Controls("tbxFieldName") = "ID" Then
    
        With listCSV.Form.Controls("cbxColumnName1")
            .Value = "None"
            .Enabled = False
        End With
        
    End If
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetCSVFieldsDisplay[ImportMap form])"
    End Select
    Resume Exit_Handler
End Sub
