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
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7560
    DatasheetFontHeight =11
    ItemSuffix =38
    Right =9240
    Bottom =10995
    DatasheetGridlinesColor =14806254
    OrderBy ="EffectiveDate DESC"
    RecSrcDt = Begin
        0x0680db994fd0e440
    End
    RecordSource ="tsys_Db_Templates"
    Caption ="_List"
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
        Begin FormHeader
            Height =1380
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
                    GridlineColor =10921638
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    Width =7260
                    Height =840
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Edit or Delete Records using the buttons for the record at right.\015\012Icon co"
                        "des at left identify if record may be edited/deleted."
                    GridlineColor =10921638
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =840
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =3600
                    Top =1020
                    Width =900
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblRemarks"
                    Caption ="Remarks"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =3600
                    LayoutCachedTop =1020
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =960
                    Top =1020
                    Width =270
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblHdrID"
                    Caption ="ID"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =960
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1230
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =2040
                    Top =1020
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTemplate"
                    Caption ="Template"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =2040
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3285
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1260
                    Top =1020
                    Width =720
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblVersion"
                    Caption ="Version"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =1260
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =4860
                    Top =780
                    Width =1020
                    Height =555
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblEffectiveDate"
                    Caption ="Effective Date"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =4860
                    LayoutCachedTop =780
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =1860
            Name ="Detail"
            OnMouseMove ="[Event Procedure]"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =6780
                    Top =840
                    Width =720
                    ForeColor =4210752
                    Name ="btnEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000303840ff404040ff505050ff504850f080686020 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000606060ff909890ffd0d0d0ffa0a8b0ff304850ff ,
                        0xa090905000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000a0a0a0fff0f0f0fff0f8ffffc0e0f0ff5090b0ff ,
                        0x204850ff80686020000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000080787080e0e0e0ffd0f0f0ff90e0f0ff50c0d0ff ,
                        0x4098b0ff204850ff806860200000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000006090a080c0e8f0ffa0f0f0ff70e0f0ff ,
                        0x50c0d0ff4098b0ff204850ff8068602000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000006090a090b0e8f0ffa0f0f0ff ,
                        0x70e0f0ff50c0d0ff4098b0ff204850ff80686020000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000006090a090b0e8f0ff ,
                        0xa0f0f0ff70e0f0ff50c0d0ff4098b0ff204850ff806860200000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006090a0a0 ,
                        0xb0e8f0ffa0f0f0ff70e0f0ff50c0d0ff4098b0ff204850ff8068602000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x6090a0a0b0e8f0ffa0f0f0ff70e0f0ff50c0d0ff4098b0ff204850ff80686020 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xd08060006090a0a0b0e8f0ffa0f0f0ff70e0f0ff50b8d0ff4098b0ff204850ff ,
                        0x8068602000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000d0d8e0006090a0b0b0e8f0ffa0f0f0ff70d0e0ff50a0b0ff808890ff ,
                        0x303870ff80686020000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000d0d8e0006090a0b0c0f0f0ffa0e0e0ffb0b0a0ff5058b0ff ,
                        0x303090ff505880ff000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000d0d8e0006090a0b0a0b8d0ff8088d0ff6070d0ff ,
                        0x303090ff202860ff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000d0d8e0006070b0b09098d0ff7078d0ff ,
                        0x4050a0ff9098b0ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000d0d8e000606090d05060a0ff ,
                        0x9090b0ff00000000
                    End

                    LayoutCachedLeft =6780
                    LayoutCachedTop =840
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =1200
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
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
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =45
                    Width =720
                    Height =315
                    FontSize =9
                    TabIndex =1
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =45
                    LayoutCachedWidth =840
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =2
                    Left =6780
                    Width =720
                    FontSize =14
                    TabIndex =2
                    ForeColor =255
                    Name ="btnDelete"
                    Caption ="í ½í·´"
                    OnClick ="[Event Procedure]"
                    FontName ="Academy Engraved LET"
                    GridlineColor =10921638

                    LayoutCachedLeft =6780
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =360
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =-1
                    BackColor =14136213
                    BorderColor =14136213
                    ThemeFontIndex =-1
                    HoverColor =15060409
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
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1875
                    Top =45
                    Width =2925
                    Height =315
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxTemplate"
                    ControlSource ="TemplateName"
                    OnMouseMove ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1875
                    LayoutCachedTop =45
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =900
                    Top =45
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =900
                    LayoutCachedTop =45
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5760
                    Top =660
                    Width =780
                    Height =1200
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxControlTip"
                    OnMouseMove ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5760
                    LayoutCachedTop =660
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1860
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1350
                    Top =45
                    Width =420
                    Height =315
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxVersion"
                    ControlSource ="Version"
                    GridlineColor =10921638

                    LayoutCachedLeft =1350
                    LayoutCachedTop =45
                    LayoutCachedWidth =1770
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4860
                    Top =45
                    Width =1020
                    Height =315
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxEffectiveDate"
                    ControlSource ="EffectiveDate"
                    ConditionalFormat = Begin
                        0x0100000056010000030000000100000000000000000000001e00000001000000 ,
                        0x22b14c00ffffff0001000000000000001f0000005c00000001000000ed1c2400 ,
                        0xffffff0001000000000000005d0000007a000000010000000000ff00ffffff00 ,
                        0x4900490066002800490073004e0075006c006c0028005b005200650074006900 ,
                        0x7200650044006100740065005d0029002c0031002c0030002900000000004900 ,
                        0x4900660028004e006f0074002000490073004e0075006c006c0028005b005200 ,
                        0x6500740069007200650044006100740065005d0029002c004900490066002800 ,
                        0x5b0052006500740069007200650044006100740065005d003c00440061007400 ,
                        0x6500280029002c0031002c00300029002c003000290000000000490049006600 ,
                        0x28005b0052006500740069007200650044006100740065005d003d0044006100 ,
                        0x74006500280029002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4860
                    LayoutCachedTop =45
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000022b14c00ffffff001d0000004900 ,
                        0x490066002800490073004e0075006c006c0028005b0052006500740069007200 ,
                        0x650044006100740065005d0029002c0031002c00300029000000000000000000 ,
                        0x00000000000000000000000000010000000000000001000000ed1c2400ffffff ,
                        0x003c00000049004900660028004e006f0074002000490073004e0075006c006c ,
                        0x0028005b0052006500740069007200650044006100740065005d0029002c0049 ,
                        0x004900660028005b0052006500740069007200650044006100740065005d003c ,
                        0x004400610074006500280029002c0031002c00300029002c0030002900000000 ,
                        0x0000000000000000000000000000000000000100000000000000010000000000 ,
                        0xff00ffffff001c00000049004900660028005b00520065007400690072006500 ,
                        0x44006100740065005d003d004400610074006500280029002c0031002c003000 ,
                        0x2900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =420
                    Height =315
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="lblType"
                    ControlSource ="=Left([TemplateName],1)"
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =420
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1500
                    Top =960
                    Width =4020
                    Height =480
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="Text35"
                    ControlSource ="Remarks"
                    GridlineColor =10921638

                    LayoutCachedLeft =1500
                    LayoutCachedTop =960
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =1440
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6000
                    Top =15
                    Width =720
                    TabIndex =10
                    ForeColor =4210752
                    Name ="btnViewSQL"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="View template SQL"
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
                        0x000000000000000000000060000000a0000000d0000000600000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000060000000f00000007000000010000000000000000000000000 ,
                        0x000000000000000000000090000000c000000080000000b00000004000000000 ,
                        0x000000a0000000e0000000e0000000b00000001000000070000000ff000000d0 ,
                        0x000000f0000000d0000000e00000003000000020000000a0000000e000000080 ,
                        0x000000f00000003000000000000000c00000009000000000000000ff00000090 ,
                        0x00000010000000b00000005000000000000000a0000000ff000000e0000000e0 ,
                        0x000000d0000000000000000000000090000000b000000000000000ff00000070 ,
                        0x000000000000002000000020000000b0000000ff000000e000000040000000f0 ,
                        0x000000b0000000000000000000000070000000ff00000000000000ff00000070 ,
                        0x000000000000000000000090000000ff000000b00000002000000020000000d0 ,
                        0x000000c0000000000000000000000090000000c000000000000000ff00000070 ,
                        0x0000000000000000000000e0000000900000000000000020000000b000000060 ,
                        0x000000f00000004000000020000000e0000000a000000000000000ff00000070 ,
                        0x000000000000000000000050000000b000000070000000d00000009000000000 ,
                        0x00000070000000a0000000c0000000800000001000000070000000ff000000c0 ,
                        0x0000002000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6000
                    LayoutCachedTop =15
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =375
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
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
                    Left =1920
                    Top =420
                    Width =420
                    Height =315
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxSyntax"
                    ControlSource ="Syntax"
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =420
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
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
' Form:         TemplateList
' Level:        Application form
' Version:      1.00
' Basis:        Dropdown form
'
' Description:  List form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
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

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidLabel(Value As String)
Public Event InvalidCaption(Value As String)

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
        Me.btnEdit.Caption = m_ButtonCaption
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

'---------------------
' Methods
'---------------------

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
            "Error encountered (#" & Err.Number & " - Form_Load[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

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
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    lblTitle.Caption = ""
    lblDirections.Caption = "Edit or Delete Records using the buttons for the record at right." _
                            & vbCrLf & "Icon codes at left identify if record may be edited/deleted."
    tbxIcon.Value = StringFromCodepoint(uLocked)
    tbxIcon.ForeColor = lngDkGreen
    lblDirections.ForeColor = lngLtBlue
    'set hover
    btnEdit.HoverColor = lngGreen
    btnDelete.HoverColor = lngGreen

    btnDelete.Caption = StringFromCodepoint(uDelete)
    btnDelete.ForeColor = lngRed

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[TemplateList form])"
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
            "Error encountered (#" & Err.Number & " - Form_Current[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnViewSQL_Click
' Description:  Delete button click actions
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
Private Sub btnViewSQL_Click()
On Error GoTo Err_Handler
    
    Dim strOA As String
    
    strOA = Me.ID.Value & "|" _
            & Me.Version.Value & "|" _
            & Me.Template.Value & "|" _
            & Me.EffectiveDate.Value & "|" _
            & Me.Syntax.Value
    
    DoCmd.OpenForm "TemplateSQL", acNormal, , , , , strOA

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnViewSQL_Click[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnEdit_Click
' Description:  Enter button click actions
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
Private Sub btnEdit_Click()
On Error GoTo Err_Handler
    
    'populate the parent form
    PopulateForm Me.Parent, ID

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEdit_Click[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnDelete_Click
' Description:  Delete button click actions
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
Private Sub btnDelete_Click()
On Error GoTo Err_Handler
    
    Dim result As Integer
    
    'identify the record ID
     result = MsgBox("Delete Record this record: #" & tbxID & " ?" _
                        & vbCrLf & "This action cannot be undone.", vbYesNo, "Delete Record?")

    If result = vbYes Then DeleteRecord "Event", ID
    
    'clear the deleted record
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[TemplateList form])"
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
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblHdrID_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblHdrID_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblHdrID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblHdrID_Click[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblVersion_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblVersion_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblVersion

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblVersion_Click[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblTemplate_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblTemplate_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblTemplate

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblTemplate_Click[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblRemarks_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblRemarks_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblRemarks

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblRemarks_Click[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblEffectiveDate_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblEffectiveDate_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblEffectiveDate

'    If InStr(Me.OrderBy, "EffectiveDate") = 0 Then
'        Me.OrderBy = "EffectiveDate"
'    ElseIf Right(Me.OrderBy, 4) = "Desc" Then
'        Me.OrderBy = "EffectiveDate"
'    Else
'        Me.OrderBy = "EffectiveDate Desc"
'    End If
'
'    Me.OrderByOn = True
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblEffectiveDate_Click[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxTemplate_MouseMove
' Description:  mouse move (hover) actions
' Assumptions:  -
'               Template Name textbox is disabled, so control tips won't display
'               Otherwise this would be tbxTemplateName_MouseMove instead & tbxTemplate would
'               not be necessary
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   hnaser, March 17, 2013
'   https://www.experts-exchange.com/questions/28067200/MS-Access-tooltip-on-a-disabled-control.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub tbxTemplate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Me.tbxTemplate.ControlTipText = Nz(FetchAddlData("tsys_Db_Templates", "Remarks", Me.tbxID)(0), "")
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxTemplate_MouseMove[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxControlTip_MouseMove
' Description:  mouse move (hover) actions
' Assumptions:  -
'               Template Name textbox is disabled, so control tips won't display
'               Otherwise this would be tbxTemplateName_MouseMove instead & tbxControlTip would
'               not be necessary
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   hnaser, March 17, 2013
'   https://www.experts-exchange.com/questions/28067200/MS-Access-tooltip-on-a-disabled-control.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub tbxControlTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Me.tbxControlTip.ControlTipText = Nz(FetchAddlData("tsys_Db_Templates", "Remarks", Me.tbxID)(0), "")
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxControlTip_MouseMove[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Detail_MouseMove
' Description:  mouse move (hover) actions
' Assumptions:  -
'               Template Name textbox is disabled, so control tips won't display
'               Otherwise this would be tbxTemplateName_MouseMove instead & tbxControlTip would
'               not be necessary
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   hnaser, March 17, 2013
'   https://www.experts-exchange.com/questions/28067200/MS-Access-tooltip-on-a-disabled-control.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Me.tbxControlTip.ControlTipText = ""
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_MouseMove[TemplateList form])"
    End Select
    Resume Exit_Handler
End Sub
