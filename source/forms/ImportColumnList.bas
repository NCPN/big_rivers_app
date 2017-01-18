Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3480
    DatasheetFontHeight =11
    ItemSuffix =105
    Right =24690
    Bottom =12585
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xeecc3f14b0d0e440
    End
    Caption ="CSV Columns"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyDown ="[Event Procedure]"
    OnGotFocus ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
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
        Begin FormHeader
            CanGrow = NotDefault
            Height =300
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="CSV fields"
                    GridlineColor =10921638
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1980
                    Top =60
                    Width =780
                    Height =180
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSetFocus"
                    GridlineColor =10921638

                    LayoutCachedLeft =1980
                    LayoutCachedTop =60
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =240
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =18935
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ComboBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =15
                    Width =3240
                    Height =314
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000a6010000030000000100000000000000000000005700000000000000 ,
                        0x00000000ffffff000000000002000000580000006000000001000000ff000000 ,
                        0xffffff00010000000000000061000000a20000000100000000800000ffffff00 ,
                        0x49004900660028004d0065002e0050006100720065006e0074002e0046006f00 ,
                        0x72006d002e0043006f006e00740072006f006c007300280022006c0069007300 ,
                        0x74005400610062006c0065004600690065006c0064007300220029002e004600 ,
                        0x6f0072006d002e0043006f006e00740072006f006c0073002800220074006200 ,
                        0x78004600690065006c0064004e0061006d006500220029003d00220049004400 ,
                        0x22002c0031002c00300029000000000022004e006f006e006500780022000000 ,
                        0x000049004900660028004c0065006e0028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d00650031005d0029003e0030002c00490049006600 ,
                        0x28005b0063006200780043006f006c0075006d006e004e0061006d0065003100 ,
                        0x5d003c003e0022004e006f006e00650022002c0031002c00300029002c003000 ,
                        0x290000000000
                    End
                    Name ="cbxColumnName1"
                    RowSourceType ="Table/Query"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =15
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =329
                    ConditionalFormat14 = Begin
                        0x01000500000001000000000000000000000000000000ffffff00560000004900 ,
                        0x4900660028004d0065002e0050006100720065006e0074002e0046006f007200 ,
                        0x6d002e0043006f006e00740072006f006c007300280022006c00690073007400 ,
                        0x5400610062006c0065004600690065006c0064007300220029002e0046006f00 ,
                        0x72006d002e0043006f006e00740072006f006c00730028002200740062007800 ,
                        0x4600690065006c0064004e0061006d006500220029003d002200490044002200 ,
                        0x2c0031002c003000290000000000000000000000000000000000000000000000 ,
                        0x0000000200000001000000ff000000ffffff000700000022004e006f006e0065 ,
                        0x0078002200000000000000000000000000000000000000000000010000000000 ,
                        0x00000100000000800000ffffff004000000049004900660028004c0065006e00 ,
                        0x28005b0063006200780043006f006c0075006d006e004e0061006d0065003100 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d00650031005d003c003e0022004e006f006e006500 ,
                        0x22002c0031002c00300029002c00300029000000000000000000000000000000 ,
                        0x0000000000000001000000000000000100000000000000ccff99005300000052 ,
                        0x00650070006c0061006300650028005b004d0065005d002e005b004e0061006d ,
                        0x0065005d002c00220063006200780043006f006c0075006d006e004e0061006d ,
                        0x00650022002c002200220029003d0046006f0072006d00730021005b0049006d ,
                        0x0070006f00720074004d00610070005d002e005b0043006f006e00740072006f ,
                        0x006c0073005d0021005b007400620078004300530056005200650063006f0072 ,
                        0x0064005d00000000000000000000000000000000000000000000000000000200 ,
                        0x000001000000ed1c2400ffffff000600000022004e006f006e00650022000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =329
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f4000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff00010000000000000008000000490000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650032005d00 ,
                        0x29003e0030002c0049004900660028005b0063006200780043006f006c007500 ,
                        0x6d006e004e0061006d00650032005d003c003e0022004e006f006e0065002200 ,
                        0x2c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName2"
                    RowSourceType ="Table/Query"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =329
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =643
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004000000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x00650032005d0029003e0030002c0049004900660028005b0063006200780043 ,
                        0x006f006c0075006d006e004e0061006d00650032005d003c003e0022004e006f ,
                        0x006e00650022002c0031002c00300029002c0030002900000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =643
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f4000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff00010000000000000008000000490000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650033005d00 ,
                        0x29003e0030002c0049004900660028005b0063006200780043006f006c007500 ,
                        0x6d006e004e0061006d00650033005d003c003e0022004e006f006e0065002200 ,
                        0x2c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName3"
                    RowSourceType ="Table/Query"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =643
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =957
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004000000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x00650033005d0029003e0030002c0049004900660028005b0063006200780043 ,
                        0x006f006c0075006d006e004e0061006d00650033005d003c003e0022004e006f ,
                        0x006e00650022002c0031002c00300029002c0030002900000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =957
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f4000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff00010000000000000008000000490000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650034005d00 ,
                        0x29003e0030002c0049004900660028005b0063006200780043006f006c007500 ,
                        0x6d006e004e0061006d00650034005d003c003e0022004e006f006e0065002200 ,
                        0x2c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName4"
                    RowSourceType ="Table/Query"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =957
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =1271
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004000000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x00650034005d0029003e0030002c0049004900660028005b0063006200780043 ,
                        0x006f006c0075006d006e004e0061006d00650034005d003c003e0022004e006f ,
                        0x006e00650022002c0031002c00300029002c0030002900000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =1271
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f4000000020000000000000002000000000000000700000001000000 ,
                        0xed1c2400ffffff00010000000000000008000000490000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650035005d00 ,
                        0x29003e0030002c0049004900660028005b0063006200780043006f006c007500 ,
                        0x6d006e004e0061006d00650035005d003c003e0022004e006f006e0065002200 ,
                        0x2c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName5"
                    RowSourceType ="Table/Query"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =1271
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =1585
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ed1c2400ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004000000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x00650035005d0029003e0030002c0049004900660028005b0063006200780043 ,
                        0x006f006c0075006d006e004e0061006d00650035005d003c003e0022004e006f ,
                        0x006e00650022002c0031002c00300029002c0030002900000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =1585
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f4000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff00010000000000000008000000490000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650036005d00 ,
                        0x29003e0030002c0049004900660028005b0063006200780043006f006c007500 ,
                        0x6d006e004e0061006d00650036005d003c003e0022004e006f006e0065002200 ,
                        0x2c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName6"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =1585
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =1899
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004000000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x00650036005d0029003e0030002c0049004900660028005b0063006200780043 ,
                        0x006f006c0075006d006e004e0061006d00650036005d003c003e0022004e006f ,
                        0x006e00650022002c0031002c00300029002c0030002900000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =1899
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f4000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff00010000000000000008000000490000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650037005d00 ,
                        0x29003e0030002c0049004900660028005b0063006200780043006f006c007500 ,
                        0x6d006e004e0061006d00650037005d003c003e0022004e006f006e0065002200 ,
                        0x2c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName7"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =1899
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =2213
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004000000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x00650037005d0029003e0030002c0049004900660028005b0063006200780043 ,
                        0x006f006c0075006d006e004e0061006d00650037005d003c003e0022004e006f ,
                        0x006e00650022002c0031002c00300029002c0030002900000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =2213
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f4000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff00010000000000000008000000490000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650038005d00 ,
                        0x29003e0030002c0049004900660028005b0063006200780043006f006c007500 ,
                        0x6d006e004e0061006d00650038005d003c003e0022004e006f006e0065002200 ,
                        0x2c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName8"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =2213
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =2527
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004000000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x00650038005d0029003e0030002c0049004900660028005b0063006200780043 ,
                        0x006f006c0075006d006e004e0061006d00650038005d003c003e0022004e006f ,
                        0x006e00650022002c0031002c00300029002c0030002900000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =2527
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f4000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff00010000000000000008000000490000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650039005d00 ,
                        0x29003e0030002c0049004900660028005b0063006200780043006f006c007500 ,
                        0x6d006e004e0061006d00650039005d003c003e0022004e006f006e0065002200 ,
                        0x2c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName9"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =2527
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =2841
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004000000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x00650039005d0029003e0030002c0049004900660028005b0063006200780043 ,
                        0x006f006c0075006d006e004e0061006d00650039005d003c003e0022004e006f ,
                        0x006e00650022002c0031002c00300029002c0030002900000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =2841
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650031003000 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500310030005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName10"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =2841
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =3155
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500310030005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500310030005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =3155
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650031003100 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500310031005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName11"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =3155
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =3469
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500310031005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500310031005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =3469
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650031003200 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500310032005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName12"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =3469
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =3783
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500310032005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500310032005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =3783
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650031003300 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500310033005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName13"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =3783
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =4097
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500310033005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500310033005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =4097
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650031003400 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500310034005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName14"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =4097
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =4411
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500310034005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500310034005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =4411
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650031003500 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500310035005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName15"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =4411
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =4725
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500310035005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500310035005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =4725
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650031003600 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500310036005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName16"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =4725
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =5039
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500310036005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500310036005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =5039
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650031003700 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500310037005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName17"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =5039
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =5353
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500310037005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500310037005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =5353
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650031003800 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500310038005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName18"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =5353
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =5667
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500310038005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500310038005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =5667
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650031003900 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500310039005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName19"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =5667
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =5981
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500310039005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500310039005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =5981
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =19
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650032003000 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500320030005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName20"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =5981
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =6295
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500320030005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500320030005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =6295
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =20
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650032003100 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500320031005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName21"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =6295
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =6609
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500320031005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500320031005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =6609
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =21
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650032003200 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500320032005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName22"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =6609
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =6923
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500320032005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500320032005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =6923
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =22
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650032003300 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500320033005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName23"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =6923
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =7237
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500320033005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500320033005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =7237
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =23
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650032003400 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500320034005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName24"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =7237
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =7551
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500320034005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500320034005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =7551
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =24
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650032003500 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500320035005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName25"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =7551
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =7865
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500320035005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500320035005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =7865
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =25
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650032003600 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500320036005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName26"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =7865
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =8179
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500320036005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500320036005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =8179
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =26
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650032003700 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500320037005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName27"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =8179
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =8493
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500320037005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500320037005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =8493
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =27
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650032003800 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500320038005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName28"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =8493
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =8807
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500320038005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500320038005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =8807
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =28
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650032003900 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500320039005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName29"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =8807
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =9121
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500320039005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500320039005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =9121
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =29
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650033003000 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500330030005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName30"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =9121
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =9435
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500330030005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500330030005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =9435
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =30
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650033003100 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500330031005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName31"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =9435
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =9749
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500330031005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500330031005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =9749
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =31
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650033003200 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500330032005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName32"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =9749
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =10063
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500330032005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500330032005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =10063
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =32
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650033003300 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500330033005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName33"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =10063
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =10377
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500330033005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500330033005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =10377
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =33
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650033003400 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500330034005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName34"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =10377
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =10691
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500330034005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500330034005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =10691
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =34
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650033003500 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500330035005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName35"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =10691
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =11005
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500330035005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500330035005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =11005
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =35
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650033003600 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500330036005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName36"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =11005
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =11319
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500330036005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500330036005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =11319
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =36
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650033003700 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500330037005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName37"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =11319
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =11633
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500330037005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500330037005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =11633
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =37
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650033003800 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500330038005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName38"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =11633
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =11947
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500330038005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500330038005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =11947
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =38
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650033003900 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500330039005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName39"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =11947
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =12261
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500330039005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500330039005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =12261
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =39
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650034003000 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500340030005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName40"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =12261
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =12575
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500340030005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500340030005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =12575
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =40
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650034003100 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500340031005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName41"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =12575
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =12889
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500340031005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500340031005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =12889
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =41
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650034003200 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500340032005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName42"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =12889
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =13203
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500340032005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500340032005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =13203
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =42
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650034003300 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500340033005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName43"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =13203
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =13517
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500340033005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500340033005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =13517
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =43
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650034003400 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500340034005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName44"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =13517
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =13831
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500340034005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500340034005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =13831
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =44
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650034003500 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500340035005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName45"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =13831
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =14145
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500340035005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500340035005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =14145
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =45
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650034003600 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500340036005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName46"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =14145
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =14459
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500340036005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500340036005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =14459
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =46
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650034003700 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500340037005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName47"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =14459
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =14773
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500340037005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500340037005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =14773
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =47
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650034003800 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500340038005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName48"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =14773
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =15087
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500340038005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500340038005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =15087
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =48
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650034003900 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500340039005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName49"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =15087
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =15401
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500340039005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500340039005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =15401
                    Width =3240
                    Height =344
                    FontSize =9
                    TabIndex =49
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650035003000 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500350030005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName50"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =15401
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =15745
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500350030005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500350030005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =15765
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =50
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650035003100 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500350031005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName51"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =15765
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =16079
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500350031005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500350031005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =16079
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =51
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650035003200 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500350032005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName52"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =16079
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =16393
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500350032005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500350032005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =16393
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =52
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650035003300 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500350033005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName53"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =16393
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =16707
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500350033005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500350033005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =16707
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =53
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650035003400 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500350034005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName54"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =16707
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =17021
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500350034005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500350034005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =17021
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =54
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650035003500 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500350035005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName55"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =17021
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =17335
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500350035005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500350035005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =17335
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =55
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650035003600 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500350036005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName56"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =17335
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =17649
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500350036005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500350036005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =17649
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =56
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650035003700 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500350037005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName57"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =17649
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =17963
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500350037005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500350037005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =17963
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =57
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650035003800 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500350038005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName58"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =17963
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =18277
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500350038005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500350038005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =18277
                    Width =3240
                    Height =314
                    FontSize =9
                    TabIndex =58
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650035003900 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500350039005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName59"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =18277
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =18591
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500350039005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500350039005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =18591
                    Width =3240
                    Height =344
                    FontSize =9
                    TabIndex =59
                    BorderColor =10921638
                    ForeColor =4138256
                    ConditionalFormat = Begin
                        0x01000000f8000000020000000000000002000000000000000700000001000000 ,
                        0xff000000ffffff000100000000000000080000004b0000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e006f006e00650022000000000049004900660028004c0065006e002800 ,
                        0x5b0063006200780043006f006c0075006d006e004e0061006d00650036003000 ,
                        0x5d0029003e0030002c0049004900660028005b0063006200780043006f006c00 ,
                        0x75006d006e004e0061006d006500360030005d003c003e0022004e006f006e00 ,
                        0x650022002c0031002c00300029002c003000290000000000
                    End
                    Name ="cbxColumnName60"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="=[PrepareImportColumns]"
                    OnGotFocus ="=SetCurrentPseudoRecord([screen].[activecontrol])"
                    OnClick ="=ChangeBackColor([Screen].[ActiveControl],[lngYelLime])"
                    OnChange ="=PopulateCSVFields([Screen].[ActiveControl])"
                    GridlineColor =10921638
                    CanGrow =255
                    CanShrink =255
                    InheritValueList =0

                    LayoutCachedTop =18591
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =18935
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ff000000ffffff00060000002200 ,
                        0x4e006f006e006500220000000000000000000000000000000000000000000001 ,
                        0x000000000000000100000000800000ffffff004200000049004900660028004c ,
                        0x0065006e0028005b0063006200780043006f006c0075006d006e004e0061006d ,
                        0x006500360030005d0029003e0030002c0049004900660028005b006300620078 ,
                        0x0043006f006c0075006d006e004e0061006d006500360030005d003c003e0022 ,
                        0x004e006f006e00650022002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
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
' Form:         ImportColumnList
' Level:        Application form
' Version:      1.04
' Basis:        Dropdown form
'
' Description:  List form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, October 18, 2016
' References:   -
' Revisions:    BLC - 10/18/2016 - 1.00 - initial version
'               BLC - 10/20/2016 - 1.01 - removed ButtonCaption, SelectedID, SelectedValue properties,
'                                         added cbxColumnName 51-60
'               BLC - 10/25/2016 - 1.02 - added RefreshCSVColumnList() to handle newly imported CSV
'               BLC - 12/8/2016  - 1.03 - revised PrepareImportColumns() to enable import button
'                                         when all import columns set, fixed Hide to include 51-60
'               BLC - 12/13/2016 - 1.04 - added current row highlighting, key up|down navigation
'               BLC - 1/17/2017  - 1.05 - adjusted to set combobox events w/in form load,
'                                         switched PrepareImportColumns() to function from sub
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
'Private m_Directions As String

Private m_Table As String
Private m_Fields As String

Private m_Records As DAO.Recordset
Private m_NumColumns As Integer
Private m_ImportColumns As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidRecords(Value As DAO.Recordset)
Public Event InvalidNumColumns(Value As Integer)

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

Public Property Let Table(Value As String)
        m_Table = Value

        'populate form
        PopulateForm
End Property

Public Property Get Table() As String
    Table = m_Table
End Property

Public Property Let Records(Value As DAO.Recordset)
    If IsRecordset(Value) Then
        Set m_Records = Value
        
        'set the form's # of records based on this rs
        Set Me.Recordset = m_Records
    Else
        RaiseEvent InvalidRecords(Value)
    End If
End Property

Public Property Get Records() As DAO.Recordset
    Set Records = m_Records
End Property

'number of columns to import (from table vs. CSV)
Public Property Let NumColumns(Value As Integer)
        m_NumColumns = Value
End Property

Public Property Get NumColumns() As Integer
    NumColumns = m_NumColumns
End Property

Public Property Get ImportColumns() As String
    ImportColumns = m_ImportColumns
End Property

Public Property Let ImportColumns(Value As String)
        m_ImportColumns = Value
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
' Source/date:  Bonnie Campbell, October 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/18/2016 - initial version
'   BLC - 10/20/2016 - code cleanup
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    Me.Title = "CSV Columns"
    Me.Table = "CSV Columns"
    
    lblTitle.Caption = Me.Title
    
    Dim i As Integer
    Dim strControl As String
    
    'hide dropdowns (1-50
    For i = 1 To 50
        strControl = "cbxColumnName" & i
        
        Me.Controls(strControl).Visible = False
    
    Next

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[ImportColumnList form])"
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
' Source/date:  Bonnie Campbell, October 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/18/2016 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'eliminate NULLs
'    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[ImportColumnList form])"
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
' Source/date:  Bonnie Campbell, October 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/18/2016 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
              
    'set psuedo current record <- set the # of the cbx
    'MsgBox ActiveControl.Name
'    MsgBox Screen.ActiveForm.Name & " is the active form."
'    MsgBox Screen.ActiveControl.Name & "is the active control."
'    If InStr(Me.ActiveControl, "cbxColumnName") Then
'        Me.Parent.Controls("tbxCSVRecord").Value = Replace(Me.ActiveControl.Name, "cbxColumnName", "")
'    End If
              
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub


'   BLC - 1/17/2017  - adjusted to set comboxbox events vs in IDE
Private Sub Form_GotFocus()
    
    Dim ctrl As Control

    'set events
    For Each ctrl In Me.Controls
        
        With ctrl
            If Len(.Name) > Len(Replace(.Name, "cbxColumnName", "")) Then
 Debug.Print "ImportColumnList - " & .Name
            'set events
            'onclick        =ChangeBackColor([Screen].[ActiveControl],[lngYelLime])
            'afterupdate    =PrepImportCols()
            'onchange       =PopulateCSVFields([Screen].[ActiveControl])
            'ongotfocus     =SetCurrentPseudoRecord([screen].[activecontrol])
            
            'using [Screen].[ActiveControl] / Me.ActiveControl --> error 2474
            
             .OnClick = ChangeBackColor(ctrl, lngYelLime)
             .AfterUpdate = PrepImportCols
             .OnChange = PopulateCSVFields(ctrl)
             '.OnGotFocus = SetCurrentPseudoRecord(ctrl)
             
            End If
        End With
            
    Next

End Sub

' ---------------------------------
' Sub:          Form_KeyDown
' Description:  form actions for arrow (up|down) navigation
' Assumptions:  Form > Events > Key Preview Property = Yes
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Missinglinq, June 12, 2011
'   http://www.utteraccess.com/forum/Arrow-Keys-Navigate-Co-t1968556.html
' Source/date:  Bonnie Campbell, December 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/13/2016 - initial version
' ---------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo Err_Handler

    Dim strControl As String
    Dim intRecord As Integer

    'intRecord = CInt(Replace(Screen.ActiveControl.Name, "cbxColumnNumber", ""))
    intRecord = Replace(Me.ActiveControl.Name, "cbxColumnName", "")
    strControl = Me.ActiveControl.Name

    Select Case KeyCode
    
        Case vbKeyDown
        
            If intRecord <> 60 Then _
                strControl = "cbxColumnName" & (intRecord + 1)
                            
'            If Me.CurrentRecord <> RecordsetClone.RecordCount Then
'                DoCmd.GoToRecord , , acNext
'            Else
'                'Do nothing
'            End If
'
'            KeyCode = 0
        
        Case vbKeyUp
            If intRecord <> 1 Then _
                strControl = "cbxColumnName" & (intRecord - 1)
            
'            If Me.CurrentRecord <> 1 Then
'                DoCmd.GoToRecord , , acPrevious
'            Else
'                'Do nothing
'            End If
'
'            KeyCode = 0
    
    End Select
    
    'KeyCode = 0
    If Me.Controls(strControl).Visible = True Then _
        Me.Controls(strControl).SetFocus

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_KeyDown[ImportColumnList form])"
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
' Source/date:  Bonnie Campbell, October 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/18/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          PopulateForm
' Description:  form populating actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
'   mbizup, 5/29/2008
'   https://www.experts-exchange.com/questions/23441990/moving-data-from-array-to-a-table-in-Vba.html
'   missinglinq, 1/31/2009
'   http://www.access-programmers.co.uk/forums/showthread.php?t=164897
' Source/date:  Bonnie Campbell, October 6, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/6/2016 - initial version
' ---------------------------------
Private Sub PopulateForm()
On Error GoTo Err_Handler
    
    'set displayed title
    lblTitle.Caption = "CSV fields"
    
'    'retrieve field info
'    Dim aryFieldInfo() As Variant 'string
'
'    aryFieldInfo = FetchDbTableFieldInfo("usys_temp_csv")
'
'    'clear table
'    ClearTable "usys_temp_rs2"
'
'    'populate w/ table data
'    Dim rs2 As DAO.Recordset
'    Dim aryRecord() As String
'    Dim i As Integer
'
'    Set rs2 = CurrentDb.OpenRecordset("usys_temp_rs2", dbOpenDynaset)
'
'    'add the "None" value
'    rs2.AddNew
'    rs2.Fields(0) = "None"
'    rs2.Update
'
'    For i = 0 To UBound(aryFieldInfo)
'
'        'create new record
'        rs2.AddNew
'
'        aryRecord = Split(aryFieldInfo(i), "|")
'
'        'rs!Column = aryRecord(0)
'        rs2.Fields(0) = aryRecord(0)
'
'        'add the new record
'        rs2.Update
'
'    Next
    
    Dim strControl As String
    Dim i As Integer

Debug.Print "PopulateForm - NumColumns = " & Me.NumColumns
    
    'expose & populate the proper # of dropdowns
    For i = 1 To Me.NumColumns 'CInt(Me.Records.RecordCount)
        strControl = "cbxColumnName" & i
Debug.Print "PopulateForm control:" & strControl

'FIX HERE!
        If i = 30 Then
            Debug.Print "30"
        End If
        
        Me.Controls(strControl).Visible = True
'        Set Me.Controls(strControl).Recordset = rs2 '<--ERROR #5302
        'Me.Controls(strControl).AddItem item:="None", index:=0
        
        'set "None" to red --> Conditional formmating = "None"
    
        'requery to refresh displayed controls
        Me.Controls(strControl).Requery
Debug.Print "PopulateForm - ListRows = " & Me.Controls(strControl).ListRows
    Next

    If Me.NumColumns > 0 Then
        'set detail to proper height
        Me.Detail.Height = Me.Controls(strControl).Height * Me.NumColumns 'Me.Records.RecordCount
    End If
    
'    Set Me.Recordset = rs
    
'    Set cbxColumnName.Recordset = rs2
    
    'set the # of repeats of the cbx
'    Set Me.Recordset = rs

Exit_Handler:
'    'cleanup
'    Set rs2 = Nothing
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateForm[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          HideControls
' Description:  Hides form controls
' Assumptions:  -
' Parameters:   WhichControls - Controls that should be hidden (string)
'                               default is "ALL" which includes all column comboboxes
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/18/2016 - initial version
'   BLC - 12/8/2016 - increased # from 50 to 60
' ---------------------------------
Public Sub HideControls(Optional WhichControls As String = "ALL")
On Error GoTo Err_Handler

    Select Case WhichControls
    
        Case "ALL"
    
            Dim strControl As String
            Dim i As Integer
            
            'set focus elsewhere to avoid error
            Me.tbxSetFocus.Enabled = True
            Me.tbxSetFocus.SetFocus
            Me.tbxSetFocus.Enabled = False
            
            'hide dropdowns
            For i = 1 To 60 'Me.NumColumns
                strControl = "cbxColumnName" & i
                
                Me.Controls(strControl).Value = ""
                Me.Controls(strControl).Visible = False
                
            Next
    
    End Select
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HideControls[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          PrepareImportColumns
' Description:  prepare the string of import columns (in order)
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/18/2016 - initial version
'   BLC - 12/8/2016 - added check for enabling import button when all columns are set
' ---------------------------------
Private Sub PrepareImportColumns()
On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim strControl As String, strImportColumns As String
    
    'default
    strImportColumns = ""
    
    For i = 1 To Me.NumColumns
    
        strControl = "cbxColumnName" & i
        
        strImportColumns = strImportColumns & Me.Controls(strControl) & ", "
    Next
    
    'avoid errors where no columns are defined
    If Len(strImportColumns) = 0 Then GoTo Exit_Handler
    
    Me.ImportColumns = Replace(Left(Trim(strImportColumns), Len(Trim(strImportColumns)) - 1), "None", "NULL")
    
    Debug.Print "PrepareImportColumns importcols = " & Me.ImportColumns
    
    'disable import on any table ID field columns
    Dim ctrl As Control
    Set ctrl = Me.Parent.Form.Controls("listTableFields").Form.Controls("tbxFieldName")
    
    If ctrl = "ID" Then
    
 '       Debug.Print "ID is here"
    
        'ensure cbxColumnName1 doesn't have focus
'        Me.tbxSetFocus.SetFocus
        
        With Me.cbxColumnName1
            .Value = "None"
            .Enabled = False
            .Locked = True
'            .Visible = False
        End With
    
'        Me.Requery
'        Debug.Print Me.cbxColumnName1.Enabled

    End If
    
    'check if all fields are represented
    Dim c As Integer
    Dim aryCols As Variant 'import columns
    Dim pcols As Integer 'populated columns
    
    aryCols = Split(Me.ImportColumns, ",")
    
    For c = 0 To UBound(aryCols)
        If Not IsZLS(Trim(aryCols(c))) Then pcols = pcols + 1
    Next
    
'    If CountInString(Me.ImportColumns, ",") + 1 = Me.NumColumns Then
    If pcols = Me.NumColumns Then
    
        Form_ImportMap.btnImport.Enabled = True
    
    End If
    
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PrepareImportColumns[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Public Function PrepImportCols()

Debug.Print "PrepImportCols Me Name - " & Me.Name

PrepareImportColumns

End Function

' ---------------------------------
' Sub:          RefreshColumnList
' Description:  refreshes dropdown data source to reflect newly imported CSV (new tsys_temp_csv data)
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 25, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/25/2016 - initial version
' ---------------------------------
Public Sub RefreshColumnList()
On Error GoTo Err_Handler

'    Me.Dirty = False
    
    PopulateForm
    
'    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshColumnList[CSVColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     SetCurrentRecord
' Description:  sets a pseudo current record # based on the combobox w/ current focus
' Assumptions:  -
' Parameters:   -
' Returns:      current # for combobox (integer)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/13/2016 - initial version
' ---------------------------------
Public Function xSetCurrentRecord() As Integer
On Error GoTo Err_Handler
              
    'set psuedo current record <- set the # of the cbx
    'MsgBox ActiveControl.Name
    MsgBox Screen.ActiveForm.Name & " is the active form."
    MsgBox Screen.ActiveControl.Name & "is the active control."
    If InStr(Me.ActiveControl, "cbxColumnName") Then
        Me.Parent.Controls("tbxCSVRecord").Value = Replace(Me.ActiveControl.Name, "cbxColumnName", "")
    End If
              
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetCurrentRecord[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Function


' ------------------------------------
'   Combobox 1-50 After Update Events
' ------------------------------------

' ---------------------------------
' Sub:          cbxColumnNameX_AfterUpdate
' Description:  combobox after update actions
' Assumptions:
'               ALL cbxColumnNameX comboboxes take the SAME actions
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/18/2016 - initial version
' ---------------------------------

' ------- Columns 1-10 ---------------
Private Sub cbxColumnName1_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName1_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName2_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName2_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName3_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName3_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName4_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName4_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName5_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName5_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName6_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName6_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName7_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName7_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName8_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName8_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName9_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName9_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName10_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName10_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

' ------- Columns 11-20 --------------
Private Sub cbxColumnName11_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName11_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName12_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName12_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName13_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName13_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName14_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName14_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName15_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName15_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName16_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName16_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName17_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName17_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName18_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName18_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName19_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName19_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName20_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName20_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

' ------- Columns 21-30 --------------
Private Sub cbxColumnName21_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName21_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName22_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName22_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName23_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName23_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName24_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName24_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName25_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName25_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName26_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName26_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName27_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName27_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName28_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName28_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName29_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName29_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName30_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName30_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

' ------- Columns 31-40 --------------
Private Sub cbxColumnName31_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName31_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName32_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName32_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName33_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName33_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName34_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName34_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName35_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName35_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName36_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName36_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName37_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName37_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName38_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName38_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName39_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName39_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName40_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName40_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

' ------- Columns 41-50 --------------
Private Sub cbxColumnName41_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName41_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName42_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName42_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName43_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName43_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName44_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName44_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName45_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName45_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName46_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName46_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName47_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName47_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName48_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName48_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName49_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName49_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName50_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName50_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

' ------- Columns 51-60 --------------
Private Sub cbxColumnName51_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName51_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName52_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName52_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName53_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName53_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName54_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName55_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName55_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName55_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName56_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName56_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName57_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName57_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName58_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName58_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName59_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName59_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cbxColumnName60_AfterUpdate()
On Error GoTo Err_Handler
    
    PrepareImportColumns
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxColumnName60_AfterUpdate[ImportColumnList form])"
    End Select
    Resume Exit_Handler
End Sub
