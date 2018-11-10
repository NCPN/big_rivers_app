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
    ItemSuffix =103
    Left =3225
    Top =2640
    Right =22995
    Bottom =16815
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x95f0359a100be540
    End
    Caption ="VegPlot"
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
            CanGrow = NotDefault
            Height =2280
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
                    Top =120
                    Width =7500
                    Height =615
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Enter the plot information and click save.\015\012Add substrates, cover species,"
                        " taglines via buttons at right."
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3300
                    Top =1920
                    Width =2520
                    Height =315
                    FontSize =10
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblModalSedSize"
                    Caption ="Modal Sediment Size (Overall)"
                    GridlineColor =10921638
                    LayoutCachedLeft =3300
                    LayoutCachedTop =1920
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =2235
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6960
                    Top =1860
                    Width =720
                    TabIndex =5
                    ForeColor =16711680
                    Name ="btnComment"
                    Caption ="í ½í·©"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6960
                    LayoutCachedTop =1860
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =2220
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
                    Left =1800
                    Top =1920
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDistance"
                    Caption ="Distance (m)"
                    GridlineColor =10921638
                    LayoutCachedLeft =1800
                    LayoutCachedTop =1920
                    LayoutCachedWidth =3045
                    LayoutCachedHeight =2235
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =1080
                    Top =1920
                    Width =600
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblNumber"
                    Caption ="Plot #"
                    GridlineColor =10921638
                    LayoutCachedLeft =1080
                    LayoutCachedTop =1920
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =2235
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =3600
                    Top =60
                    Width =4140
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblContext"
                    Caption ="context"
                    GridlineColor =10921638
                    LayoutCachedLeft =3600
                    LayoutCachedTop =60
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6120
                    Top =1860
                    Width =720
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnSetObserverRecorder"
                    Caption ="í ½í±¥"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set the selected veg plot's observer & recorder"
                    GridlineColor =10921638

                    LayoutCachedLeft =6120
                    LayoutCachedTop =1860
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =2220
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
                    Left =180
                    Top =900
                    Width =600
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblEvent"
                    Caption ="Event"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =900
                    LayoutCachedWidth =780
                    LayoutCachedHeight =1215
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =8
                    Left =1080
                    Top =900
                    Width =3414
                    Height =315
                    ColumnOrder =4
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
                    Name ="cbxEvent"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Event (sample visit)"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1080
                    LayoutCachedTop =900
                    LayoutCachedWidth =4494
                    LayoutCachedHeight =1215
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =4620
                    Top =900
                    TabIndex =2
                    ForeColor =16711680
                    Name ="btnAddEvent"
                    Caption ="í ½í·“  Add Event"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add a new event/sampling visit"
                    GridlineColor =10921638

                    LayoutCachedLeft =4620
                    LayoutCachedTop =900
                    LayoutCachedWidth =6060
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
                    OverlapFlags =85
                    Left =180
                    Top =1380
                    Width =855
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTransect"
                    Caption ="Transect"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =1380
                    LayoutCachedWidth =1035
                    LayoutCachedHeight =1695
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =9
                    Left =1080
                    Top =1380
                    Width =3414
                    Height =315
                    ColumnOrder =3
                    TabIndex =1
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
                    Name ="cbxTransect"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;0;0;0;1"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Veg transect"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1080
                    LayoutCachedTop =1380
                    LayoutCachedWidth =4494
                    LayoutCachedHeight =1695
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =4620
                    Top =1380
                    TabIndex =3
                    ForeColor =16711680
                    Name ="btnAddTransect"
                    Caption ="  Add Transect"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add a new veg transect"
                    GridlineColor =10921638

                    LayoutCachedLeft =4620
                    LayoutCachedTop =1380
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =1740
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
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6240
                    Top =960
                    Width =960
                    Height =300
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =6
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="tbxEventID"
                    DefaultValue ="0"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =6240
                    LayoutCachedTop =960
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ffffff00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6240
                    Top =1380
                    Width =960
                    Height =300
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =7
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="tbxTransectID"
                    DefaultValue ="0"
                    ConditionalFormat = Begin
                        0x0100000084000000010000000100000000000000000000001100000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004400450056005f004d004f00440045005d003d00460061006c0073006500 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =6240
                    LayoutCachedTop =1380
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =1680
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ffffff00100000005b00 ,
                        0x4400450056005f004d004f00440045005d003d00460061006c00730065000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7320
                    Top =960
                    Width =420
                    Height =300
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =8
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="tbxDevMode"
                    DefaultValue ="0"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000000000002000000000000000600000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x460061006c007300650000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =7320
                    LayoutCachedTop =960
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x01000100000000000000020000000100000000000000ffffff00050000004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =9105
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =180
                    Top =930
                    Width =5520
                    Height =2760
                    FontSize =14
                    LeftMargin =216
                    TopMargin =288
                    BackColor =16051931
                    Name ="lblPct"
                    Caption ="%"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =930
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =3690
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =223
                    Left =2580
                    Top =2940
                    Width =5100
                    Height =750
                    FontSize =14
                    LeftMargin =72
                    TopMargin =144
                    BackColor =8576482
                    Name ="lblSizeClass"
                    GridlineColor =10921638
                    LayoutCachedLeft =2580
                    LayoutCachedTop =2940
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =3690
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =6300
                    Top =2520
                    Width =1140
                    TabIndex =29
                    ForeColor =4210752
                    Name ="btnSubstrateCover"
                    Caption ="Substrates"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit Substrate Cover"
                    GridlineColor =10921638

                    LayoutCachedLeft =6300
                    LayoutCachedTop =2520
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =2880
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
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =180
                    Top =450
                    Width =5520
                    Height =420
                    FontSize =14
                    LeftMargin =29
                    TopMargin =14
                    BackColor =11397670
                    Name ="lblSpecialPlot"
                    Caption ="Special Plot?"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =450
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =870
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =5820
                    Top =4230
                    Width =1860
                    Height =720
                    FontSize =20
                    LeftMargin =72
                    TopMargin =72
                    BackColor =15855852
                    Name ="lblPlotDensityBgd"
                    Caption ="â¡¢"
                    FontName ="Arial"
                    ControlTipText ="Plot Density"
                    GridlineColor =10921638
                    LayoutCachedLeft =5820
                    LayoutCachedTop =4230
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =4950
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =5820
                    Top =480
                    Width =1860
                    Height =1560
                    FontSize =14
                    LeftMargin =72
                    TopMargin =144
                    BackColor =12444887
                    Name ="lblCover"
                    Caption ="Cover\015\012Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =5820
                    LayoutCachedTop =480
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =2040
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =180
                    Top =3810
                    Width =5580
                    Height =1140
                    FontSize =14
                    LeftMargin =216
                    TopMargin =288
                    BackColor =11916796
                    Name ="lblChkboxes"
                    Caption ="âœ”"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =3810
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =4950
                    BackThemeColorIndex =9
                    BackTint =40.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6660
                    Top =60
                    Width =720
                    TabIndex =21
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
                Begin Subform
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =105
                    Top =5475
                    Width =7650
                    Height =3510
                    TabIndex =26
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.VegPlotList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =5475
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =8985
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =5820
                    Top =60
                    Width =720
                    TabIndex =30
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

                    LayoutCachedLeft =5820
                    LayoutCachedTop =60
                    LayoutCachedWidth =6540
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
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =5355
                    Width =7860
                    Height =3750
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =5355
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =9105
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
                    TabIndex =27
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
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1980
                    Top =60
                    Width =720
                    Height =315
                    TabIndex =1
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDistance"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1980
                    LayoutCachedTop =60
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Width =720
                    Height =315
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxNumber"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =1020
                    Top =1890
                    Height =315
                    TabIndex =7
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctFines"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Percent of plot covered by fines"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =1890
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =2205
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =2580
                    Top =1890
                    Height =315
                    TabIndex =8
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctWater"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Percent of plot covered by water (inundated)"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =1890
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =2205
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =2580
                    Top =1260
                    Height =315
                    TabIndex =5
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctURC"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Total percent understory rooted cover (URC) for the plot"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =1260
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =1575
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =6840
                    Top =4530
                    Width =780
                    Height =315
                    TabIndex =20
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPlotDensity"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Set plot density 1/x where X = 1, 2, 3, 4, or 8"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =6840
                    LayoutCachedTop =4530
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =4845
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =1020
                    Top =3240
                    Height =315
                    TabIndex =13
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctFA"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Percent of plot covered by filamentous algae"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =3240
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =3555
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =1020
                    Top =3000
                    Width =1020
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFA"
                    Caption ="FA"
                    ControlTipText ="Percent of plot covered by filamentous algae"
                    GridlineColor =10921638
                    LayoutCachedLeft =1020
                    LayoutCachedTop =3000
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =3240
                End
                Begin Label
                    OverlapFlags =255
                    Left =6420
                    Top =4230
                    Width =1200
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPlotDensity"
                    Caption ="Plot Density"
                    GridlineColor =10921638
                    LayoutCachedLeft =6420
                    LayoutCachedTop =4230
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =4545
                End
                Begin Label
                    OverlapFlags =215
                    Left =1020
                    Top =1650
                    Width =1020
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFines"
                    Caption ="Fines"
                    ControlTipText ="Percent of plot covered by fines"
                    GridlineColor =10921638
                    LayoutCachedLeft =1020
                    LayoutCachedTop =1650
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =1890
                End
                Begin Label
                    OverlapFlags =215
                    Left =2580
                    Top =1650
                    Width =1020
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblWater"
                    Caption ="Water"
                    ControlTipText ="Percent of plot covered by water (inundated)"
                    GridlineColor =10921638
                    LayoutCachedLeft =2580
                    LayoutCachedTop =1650
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =1890
                End
                Begin Label
                    OverlapFlags =215
                    Left =2580
                    Top =1020
                    Width =1020
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblURC"
                    Caption ="Total URC"
                    ControlTipText ="Total percent understory rooted cover (URC) for the plot"
                    GridlineColor =10921638
                    LayoutCachedLeft =2580
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =1260
                End
                Begin Label
                    OverlapFlags =247
                    Left =6480
                    Top =4530
                    Width =315
                    Height =315
                    BorderColor =8355711
                    Name ="lblFraction"
                    Caption ="1 /"
                    GridlineColor =10921638
                    LayoutCachedLeft =6480
                    LayoutCachedTop =4530
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =4845
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =28
                    BorderColor =8355711
                    ForeColor =255
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =840
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =6840
                    Top =1080
                    Width =720
                    TabIndex =23
                    ForeColor =4210752
                    Name ="btnURC"
                    Caption ="URC"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit Understory Rooted Cover Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =6840
                    LayoutCachedTop =1080
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =1440
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
                    OverlapFlags =215
                    Left =6840
                    Top =600
                    Width =720
                    TabIndex =22
                    ForeColor =4210752
                    Name ="btnWCC"
                    Caption ="WCC"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit Woody Canopy Cover Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =6840
                    LayoutCachedTop =600
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =960
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
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =6840
                    Top =1560
                    Width =720
                    TabIndex =24
                    ForeColor =4210752
                    Name ="btnARC"
                    Caption ="ARC"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit All Rooted Cover Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =6840
                    LayoutCachedTop =1560
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =1920
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
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =4380
                    Top =3240
                    Width =2694
                    Height =315
                    TabIndex =15
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxModalSedSize"
                    RowSourceType ="Table/Query"
                    RowSource ="PARAMETERS etype Text ( 255 ); SELECT DISTINCT id, label, summary, label & ' - '"
                        " & summary AS display, Sequence FROM AppEnum WHERE EnumType = ModWentworthClassS"
                        "ize ORDER BY Sequence; "
                    ColumnWidths ="0;1872;576;0;0;0"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Modal Sediment Size Class"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =4380
                    LayoutCachedTop =3240
                    LayoutCachedWidth =7074
                    LayoutCachedHeight =3555
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
                Begin Label
                    OverlapFlags =215
                    Left =300
                    Top =1830
                    Width =600
                    Height =420
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblHintPct"
                    Caption ="Nearest 1% or T"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =1830
                    LayoutCachedWidth =900
                    LayoutCachedHeight =2250
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =5115
                    Width =7860
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
                    LayoutCachedTop =5115
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =5430
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =3960
                    Top =4980
                    Width =825
                    Height =600
                    FontSize =20
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =16772541
                    Name ="lblMsgIcon"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =3960
                    LayoutCachedTop =4980
                    LayoutCachedWidth =4785
                    LayoutCachedHeight =5580
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ToggleButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =3180
                    Top =4410
                    Width =270
                    Height =300
                    TabIndex =19
                    Name ="tglBeaverBrowse"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Plot shows beaver browse activity"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =3180
                    LayoutCachedTop =4410
                    LayoutCachedWidth =3450
                    LayoutCachedHeight =4710
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    Gradient =12
                    BackColor =12419407
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderColor =12419407
                    BorderTint =100.0
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3540
                            Top =4410
                            Width =1530
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblBeaverBrowse"
                            Caption ="Beaver Browse"
                            ControlTipText ="Plot shows beaver browse activity"
                            GridlineColor =10921638
                            LayoutCachedLeft =3540
                            LayoutCachedTop =4410
                            LayoutCachedWidth =5070
                            LayoutCachedHeight =4725
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =3180
                    Top =3990
                    Width =270
                    Height =299
                    TabIndex =18
                    Name ="tglNoIndicatorSpecies"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Plot has no indicator species"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =3180
                    LayoutCachedTop =3990
                    LayoutCachedWidth =3450
                    LayoutCachedHeight =4289
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    Gradient =12
                    BackColor =12419407
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderColor =12419407
                    BorderTint =100.0
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3540
                            Top =3990
                            Width =1965
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblNoIndicatorSpecies"
                            Caption ="No Indicator Species"
                            ControlTipText ="Plot has no indicator species"
                            GridlineColor =10921638
                            LayoutCachedLeft =3540
                            LayoutCachedTop =3990
                            LayoutCachedWidth =5505
                            LayoutCachedHeight =4305
                        End
                    End
                End
                Begin ToggleButton
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =1020
                    Top =4410
                    Width =270
                    Height =299
                    TabIndex =17
                    Name ="tglNoRootedVeg"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Plot has no rooted vegetation"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =4410
                    LayoutCachedWidth =1290
                    LayoutCachedHeight =4709
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    Gradient =12
                    BackColor =12419407
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderColor =12419407
                    BorderTint =100.0
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =215
                            Left =1380
                            Top =4410
                            Width =1470
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblNoRootedVeg"
                            Caption ="No Rooted Veg"
                            ControlTipText ="Plot has no rooted vegetation"
                            GridlineColor =10921638
                            LayoutCachedLeft =1380
                            LayoutCachedTop =4410
                            LayoutCachedWidth =2850
                            LayoutCachedHeight =4725
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =1020
                    Top =3990
                    Width =270
                    Height =299
                    TabIndex =16
                    Name ="tglNoCanopyVeg"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Plot has no canopy vegetation"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =3990
                    LayoutCachedWidth =1290
                    LayoutCachedHeight =4289
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    Gradient =12
                    BackColor =12419407
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderColor =12419407
                    BorderTint =100.0
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =1380
                            Top =3990
                            Width =1485
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblNoCanopyVeg"
                            Caption ="No Canopy Veg"
                            ControlTipText ="Plot has no canopy vegetation"
                            GridlineColor =10921638
                            LayoutCachedLeft =1380
                            LayoutCachedTop =3990
                            LayoutCachedWidth =2865
                            LayoutCachedHeight =4305
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =1020
                    Top =2565
                    Height =315
                    TabIndex =10
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctLitter"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Percent of plot covered by litter"
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =2565
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =2880
                    DatasheetCaption ="Litter"
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =2580
                    Top =2565
                    Height =315
                    TabIndex =11
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctWoodyDebris"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Percent of plot covered by woody debris"
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =2565
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =2880
                    DatasheetCaption ="Woody Debris"
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =2580
                    Top =2325
                    Width =1380
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblWoodyDebris"
                    Caption ="Woody Debris"
                    ControlTipText ="Percent of plot covered by woody debris"
                    GridlineColor =10921638
                    LayoutCachedLeft =2580
                    LayoutCachedTop =2325
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =2565
                End
                Begin Label
                    OverlapFlags =215
                    Left =1020
                    Top =2325
                    Width =1020
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLitter"
                    Caption ="Litter"
                    ControlTipText ="Percent of plot covered by litter"
                    GridlineColor =10921638
                    LayoutCachedLeft =1020
                    LayoutCachedTop =2325
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =2565
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =4200
                    Top =2565
                    Height =315
                    TabIndex =12
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctStandingDead"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Percent of plot woody canopy covered by standing dead (rooted/non-rotted), all s"
                        "pecies."
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4200
                    LayoutCachedTop =2565
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =2880
                    DatasheetCaption ="Standing Dead"
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =1020
                    Top =1260
                    Height =315
                    TabIndex =4
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctWCC"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Total percent woody canopy cover (WCC) for the plot"
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =1260
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =1575
                    DatasheetCaption ="Total WCC"
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =1020
                    Top =1020
                    Width =1380
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblWCC"
                    Caption ="Total WCC"
                    ControlTipText ="Total percent woody canopy cover (WCC) for the plot"
                    GridlineColor =10921638
                    LayoutCachedLeft =1020
                    LayoutCachedTop =1020
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =1260
                End
                Begin Label
                    OverlapFlags =215
                    Left =4200
                    Top =2325
                    Width =1380
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblStandingDead"
                    Caption ="Standing Dead"
                    ControlTipText ="Percent of plot woody canopy covered by standing dead (rooted/non-rotted), all s"
                        "pecies."
                    GridlineColor =10921638
                    LayoutCachedLeft =4200
                    LayoutCachedTop =2325
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =2565
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =3240
                    Top =570
                    Width =300
                    Height =300
                    TabIndex =2
                    BorderColor =10921638
                    Name ="chkCalibrationPlot"
                    DefaultValue ="0"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Check if plot is a calibration plot"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =570
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =870
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2160
                            Top =510
                            Width =1080
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblCalibrate"
                            Caption ="Calibration"
                            GridlineColor =10921638
                            LayoutCachedLeft =2160
                            LayoutCachedTop =510
                            LayoutCachedWidth =3240
                            LayoutCachedHeight =825
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =4680
                    Top =570
                    Width =300
                    Height =300
                    TabIndex =3
                    BorderColor =10921638
                    Name ="chkReplicatePlot"
                    DefaultValue ="0"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Check if plot is a replicate plot"
                    GridlineColor =10921638

                    LayoutCachedLeft =4680
                    LayoutCachedTop =570
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =870
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3780
                            Top =510
                            Width =900
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblReplicate"
                            Caption ="Replicate"
                            GridlineColor =10921638
                            LayoutCachedLeft =3780
                            LayoutCachedTop =510
                            LayoutCachedWidth =4680
                            LayoutCachedHeight =825
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =4200
                    Top =1260
                    Height =315
                    TabIndex =6
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctARC"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    ControlTipText ="Total percent all rooted cover (ARC) for the plot"
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4200
                    LayoutCachedTop =1260
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =1575
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =4200
                    Top =1020
                    Width =1020
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblARC"
                    Caption ="Total ARC"
                    ControlTipText ="Total percent all rooted cover (ARC) for the plot"
                    GridlineColor =10921638
                    LayoutCachedLeft =4200
                    LayoutCachedTop =1020
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1260
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =2640
                    Top =3240
                    Height =315
                    TabIndex =14
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctMSS"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    ControlTipText ="Total percent modal sediment cover (MSS) for the plot"
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =3240
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =3555
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =2640
                    Top =3000
                    Width =1020
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblMSS"
                    Caption ="Total MSS"
                    ControlTipText ="Total percent modal sediment cover (MSS) for the plot"
                    GridlineColor =10921638
                    LayoutCachedLeft =2640
                    LayoutCachedTop =3000
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =3240
                End
                Begin Label
                    OverlapFlags =223
                    Left =4200
                    Top =1650
                    Width =1155
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPctSocialTrails"
                    Caption ="Social Trails"
                    ControlTipText ="Total percent social trail cover for the plot"
                    GridlineColor =10921638
                    LayoutCachedLeft =4200
                    LayoutCachedTop =1650
                    LayoutCachedWidth =5355
                    LayoutCachedHeight =1965
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =223
                    Left =5820
                    Top =2100
                    Width =1860
                    Height =660
                    FontSize =20
                    LeftMargin =72
                    TopMargin =72
                    BackColor =12835293
                    Name ="lblTagline"
                    Caption ="í ½í³"
                    ControlTipText ="Add/Edit Tagline Measurements"
                    GridlineColor =10921638
                    LayoutCachedLeft =5820
                    LayoutCachedTop =2100
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =2760
                    BackThemeColorIndex =3
                    BackShade =90.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =6840
                    Top =2220
                    Width =720
                    Height =480
                    TabIndex =25
                    ForeColor =4210752
                    Name ="btnTaglines"
                    Caption ="Tagline"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit Tagline Measurements"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a890ff604830ff604830ff604830ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a890fffff0e0ffffe0d0ffffe0c0ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0fffff8f0ff000000ff000000ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0fffff8f0fffff8f0ffffe8d0ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0fffff8fffffff8f0ff000000ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0fffffffffffff8fffffff0e0ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0ffffffffff000000ff000000ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0fffffffffffffffffff0f0f0ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0ffffffffffffffffff000000ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0b0a0ffffffffffffffffffffffffff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0b0a0ffa08870ff806040ff705040ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff00000000c0b0a0fffffffffff0f8f0fff0f8f0ff705040ffffffffff ,
                        0x000000fff0f0f0ff000000ffffe8d0ff000000ffffe0c0ff000000ffffd8c0ff ,
                        0x604830ff00000000c0b0a0fffffffffffffffffff0f8f0ff805840ffffffffff ,
                        0xffffffffffffffff000000fffff8f0fffff0e0ffffe8e0ff000000ffffd8c0ff ,
                        0x604830ff00000000c0b0a0ffffffffffffffffffffffffffa08070ffffffffff ,
                        0xfffffffffffffffffffffffffff8f0fffff8f0fffff0e0ffffe8e0ffffe8d0ff ,
                        0x604830ff00000000c0b0a0ffc0b0a0ffc0b0a0ffc0b0a0ffc0b0a0ffc0b0a0ff ,
                        0xc0a8a0ffc0a8a0ffc0a8a0ffc0a8a0ffc0a8a0ffc0a8a0ffc0a8a0ffc0a890ff ,
                        0xc0a890ff00000000
                    End

                    LayoutCachedLeft =6840
                    LayoutCachedTop =2220
                    LayoutCachedWidth =7560
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
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =4200
                    Top =1890
                    Height =315
                    TabIndex =9
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctSocialTrails"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Total percent social trail cover for the plot"
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4200
                    LayoutCachedTop =1890
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =2205
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =255
                    Left =4380
                    Top =3000
                    Width =2895
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblModalSedimentSize"
                    Caption ="Modal Sediment Size (Overall)"
                    ControlTipText ="Modal Sediment Size Class"
                    GridlineColor =10921638
                    LayoutCachedLeft =4380
                    LayoutCachedTop =3000
                    LayoutCachedWidth =7275
                    LayoutCachedHeight =3315
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =2880
                    Width =2040
                    Height =420
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblHintPlotNumDist"
                    Caption ="Calibration/Replicate Plot? Set plot #/distance > 900"
                    ControlTipText ="Set plot number or distance > 900 if this is a calibration or replicate plot"
                    GridlineColor =10921638
                    LayoutCachedLeft =2880
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7200
                    Top =3240
                    Width =420
                    Height =300
                    FontSize =9
                    TabIndex =31
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxMSSID"
                    DefaultValue ="0"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xe2dd8200ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =7200
                    LayoutCachedTop =3240
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =3540
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000e2dd8200ffffff00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
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
' Form:         VegPlot
' Level:        Application form
' Version:      1.17
' Basis:        Dropdown form
'
' Description:  Vegplot form object related properties, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
'               BLC - 8/23/2016 - 1.01 - changed ReadyForSave() to public for
'                                        mod_App_Data Upsert/SetRecord()
'               BLC - 9/8/2016  - 1.02 - added SetObserverRecorder button
'               BLC - 10/3/2016 - 1.03 - disable taglines for CANY & DINO
'               BLC - 10/25/2016 - 1.04 - add CallingForm property & remove ButtonCaption,
'                                         SelectedID, SelectedValue properties
'               BLC - 1/9/2017 - 1.05 - added cbxEvent, observer/recorder, substrate cover %
'
'               BLC - 1/11/2017 - 1.06 - changed checkboxes (chk) to toggles (tgl)
'                                        & converted -1/0 values to 1/0 for SQL clarity,
'                                        changed event/transect display based on site/feature set
'               BLC - 1/12/2017 - 1.07 - revised to VegTransect vs. Transect form,
'                                        added % litter, % woody debris (all parks),
'                                        Total WCC %, standing dead
'               BLC - 9/25/2017 - 1.08 - revise for NCPN_framework.XX classes
'               BLC - 9/27/2017 - 1.09 - update to use Factory.NewClassXX() vs GetClass()
'               BLC - 11/1/2017 - 1.10 - added chkCalibration, chkReplicate
'               BLC - 11/7/2017 - 1.11 - fix cbxEvent column fofunctionalityr date (cbxEvent_AfterUpdate())
'               BLC - 11/10/2017 - 1.12 - updated park specific controls, updated comment click,
'                                         update to handle unset values using Ne()
'               BLC - 11/11/2017 - 1.13 - update percent values to use SetTrace()
'               BLC - 11/26/2017 - 1.14 - updated PctFines validation, added PctSocialTrails
'               BLC - 12/5/2017  - 1.15 - add VegPlot BeaverBrowse
'               BLC - 12/27/2017 - 1.16 - updated checkbox click events to properly check values
'               BLC - 8/25/2018  - 1.17 - adjusted DINO to include WCC & % Standing Dead (2018 & beyond)
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
        m_CallingForm = Value
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
' Assumptions:
'   Common fields across all parks (BLCA, CANY, DINO):
'       % Water (inundation)        Modal Sediment Size Class   % Modal Sediment Size
'       % Sand & Finer              % Modal Sediment Size       % Filamentous Algae
'       % Litter                    % Woody Debris
'       Calibration Plot            Replicate Plot
'
'   Certain fields are not applicable across all parks:
'       % Total ARC - DINO & CANY
'       % Total URC - BLCA
'       % Total WCC - BLCA & CANY, (DINO 2018 & beyond only)
'       % Standing Dead - BLCA & CANY, DINO 2018 & beyond
'       Transect - BLCA & CANY
'       Plot # - DINO only
'       Plot Distance - BLCA & CANY
'       Plot Density - BLCA & CANY << BLCA only
'       No Canopy Veg - BLCA & CANY
'       No Indicator Species - BLCA only
'       No Rooted Veg - DINO & CANY
'       Has Social Trail - BLCA only << replaced by PctSocialTrails
'       Taglines - BLCA
'
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 7/13/2016 - added validation, hints
'   BLC - 9/8/2016  - added SetObserverRecorder button
'   BLC - 10/24/2016 - revised to use CallingForm property, GetContext()
'   BLC - 1/9/2017 - added cbxEvent, observer/recorder, substrate cover % functionality
'   BLC - 1/11/2017 - changed event & transect to display based on site/feature set
'   BLC - 11/1/2017 - added chkCalibration, chkReplicate, tbxPctMSS, & updated park settings/documentation
'   BLC - 11/10/2017 - updated park specific controls
'   BLC - 11/26/2017 - updated PctFines validation, added PctSocialTrails
'   BLC - 12/5/2017  - added BeaverBrowse
'   BLC - 8/25/2018 - added DINO WCC, % Standing Dead (2018 forward only)
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "Main"
    
    If Len(Me.OpenArgs) > 0 Then Me.CallingForm = Me.OpenArgs
    
    'minimize calling form
    ToggleForm Me.CallingForm, -1
    
    'dev mode
    tbxDevMode = DEV_MODE

    'set context - based on TempVars
    lblContext.ForeColor = lngLime
    lblContext.Caption = GetContext()
                 
    Title = "VegPlot"
    lblTitle.Caption = "" 'clear header title
    Directions = "Enter the plot information and click save." _
                & vbCrLf & "Add substrates, cover species, taglines via buttons at right."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
    btnComment.Caption = StringFromCodepoint(uComment)
    btnComment.ForeColor = lngBlue
    lblChkboxes.Caption = StringFromCodepoint(uCheck)
    lblCover.Caption = "Cover" & vbCrLf & "Species"
    lblTagline.Caption = StringFromCodepoint(uRuler)
    lblPlotDensityBgd.Caption = StringFromCodepoint(uBrailleDots267)
    btnSetObserverRecorder.Caption = StringFromCodepoint(uUsers)
    btnSetObserverRecorder.ControlTipText = "Set the selected veg plot's observer & recorder"
    lblSizeClass.Caption = ""
    
    'disable Sediment Cover section << MSS % cover handled in VegPlot (for now)
    'lblSubstrates.Visible = False
    btnSubstrateCover.Visible = False
    'hide original size class label
    lblModalSedSize.Visible = False
    
    'set checkbox defaults to avoid black boxes
    'chkCalibrationPlot.DefaultValue = False
    'chkReplicatePlot.DefaultValue = False
    
    'hints
    lblHintPlotNumDist.Caption = "Calibration/Replicate Plot? Set plot #/distance > 900"
    lblHintPlotNumDist.ForeColor = lngBlue
    lblHintPlotNumDist.ControlTipText = "Set plot number or distance > 900 if this is a calibration or replicate plot"
    lblHintPlotNumDist.Visible = False
    lblHintPct.Caption = "Nearest 1% or T"
    lblHintPct.ForeColor = lngBlue
    
    'validation
    Dim strPctValidation As String, strValidationText As String
    strPctValidation = "Is Null Or ""T"" Or Between 0 and 101"
    tbxPctURC.ValidationRule = strPctValidation
    tbxPctFines.ValidationRule = strPctValidation '"Not Like ""[0-9]*.[0-9]*"""
    tbxPctWater.ValidationRule = strPctValidation
    tbxPctFA.ValidationRule = strPctValidation
    tbxPctLitter.ValidationRule = strPctValidation
    tbxPctWoodyDebris.ValidationRule = strPctValidation
    tbxPctWCC.ValidationRule = strPctValidation
    tbxPctStandingDead.ValidationRule = strPctValidation
    tbxPctARC.ValidationRule = strPctValidation
    tbxPctMSS.ValidationRule = strPctValidation
    tbxPctSocialTrails.ValidationRule = strPctValidation
    
    strValidationText = "Values may be whole percentages (0-100), 0.5, or T"
    tbxPctURC.ValidationText = strValidationText
    tbxPctFines.ValidationText = strValidationText
    tbxPctWater.ValidationText = strValidationText
    tbxPctFA.ValidationText = strValidationText
    tbxPctLitter.ValidationText = strValidationText
    tbxPctWoodyDebris.ValidationText = strValidationText
    tbxPctWCC.ValidationText = strValidationText
    tbxPctStandingDead.ValidationText = strValidationText
    tbxPctARC.ValidationText = strValidationText
    tbxPctMSS.ValidationText = strValidationText
    tbxPctSocialTrails.ValidationText = strValidationText
    
    'set hover
    btnSetObserverRecorder.HoverColor = lngGreen
    btnComment.HoverColor = lngGreen
    btnSubstrateCover.HoverColor = lngGreen
    btnTaglines.HoverColor = lngGreen
    btnWCC.HoverColor = lngGreen
    btnURC.HoverColor = lngGreen
    btnARC.HoverColor = lngGreen
    btnSave.HoverColor = lngGreen
    btnUndo.HoverColor = lngGreen
    tglNoCanopyVeg.HoverColor = lngGreen
    tglNoRootedVeg.HoverColor = lngGreen
    tglNoIndicatorSpecies.HoverColor = lngGreen
    'tglHasSocialTrails.HoverColor = lngGreen 'replaced by tbxPctSocialTrails
    tglBeaverBrowse.HoverColor = lngGreen
    
    'defaults
    tbxIcon.ForeColor = lngRed

    tbxNumber.BackColor = lngYellow
    tbxDistance.BackColor = lngYellow
    cbxModalSedSize.BackColor = lngYellow
    tbxPctURC.BackColor = lngYellow
    tbxPctFines.BackColor = lngYellow
    tbxPctWater.BackColor = lngYellow
    tbxPctFA.BackColor = lngYellow
    tbxPctLitter.BackColor = lngYellow
    tbxPctWoodyDebris.BackColor = lngYellow
    tbxPctWCC.BackColor = lngYellow
    tbxPctStandingDead.BackColor = lngYellow
    tbxPctARC.BackColor = lngYellow
    tbxPctMSS.BackColor = lngYellow
    tbxPlotDensity.BackColor = lngYellow
    
    btnComment.Enabled = False
    btnSave.Enabled = False
    btnSetObserverRecorder.Enabled = False
    btnSubstrateCover.Enabled = False
    btnWCC.Enabled = False
    btnURC.Enabled = False
    btnARC.Enabled = False
    btnTaglines.Enabled = False
    
    'disable until Event selected
    Me.cbxModalSedSize.Enabled = False
    
    'determine what level to populate
    Dim efilter As String, tfilter As String
    
    'site is default <-- cannot reach VegPlot if site isn't set
    efilter = "s_events_by_site"
    cbxEvent.ColumnCount = 6
    'cbxEvent.ColumnWidths = "0;.2in;.2in;.2in;0;0"
    cbxEvent.ColumnHeads = True
    tfilter = "s_vegtransect_by_site"
    cbxTransect.ColumnCount = 8
    'cbxTransect.ColumnWidths = "0;0;0;0;2in;0;0;0"

    Select Case TempVars("ParkCode")
        Case "BLCA" 'feature level if set
            If Not TempVars("Feature") Is Nothing Then _
                efilter = "s_events_by_feature"
                cbxEvent.ColumnCount = 8
 '               cbxEvent.ColumnWidths = "0;0;0;0;2in;0;0;0;0"
'                cbxTransect.ColumnCount = 8
'                cbxTransect.ColumnWidths = "0;0;0;0;2in;0;0;0;0"
                
                tfilter = "s_vegtransect_by_feature"
                cbxTransect.ColumnCount = 9
 '               cbxTransect.ColumnWidths = "0;0;0;0;2in;0;0;0;0"
                
        Case "CANY" 'site level
        Case "DINO" 'no transects/plots
    End Select
    
    'populate events
    'BLCA (7): ID,StartDate,FeatureEventDate,FeatureNameEventDate,SiteCode,Feature,ParkCode
    'CANY,DINO (5): ID,StartDate,SiteEventDate,SiteCode,ParkCode
    'ID, SiteCode, SiteName, Segment, ddEvent
    Set cbxEvent.Recordset = GetRecords(efilter) '"s_events_by_park_river")
    cbxEvent.BoundColumn = 1
    cbxEvent.ColumnCount = 5
    cbxEvent.ColumnWidths = "0;0;.5in;0;0"
    
    'populate veg transects
    Set cbxTransect.Recordset = GetRecords(tfilter) 's_transect_by_park_river")
    cbxTransect.BoundColumn = 1
    'cbxTransect.ColumnCount = 5
    'cbxTransect.ColumnWidths = "0;0;0;0;2in;0;0;0;0"
    
    'populate modal sediment size
    ' -------------------------------------------------------------------------------------
    ' NOTE: s_enums_for_type *MUST* include "DISTINCT" for the combobox autoexpand to work!(Access bug)
    '       Dan Some, August 7, 2011
    '       http://answers.microsoft.com/en-us/office/forum/office_2007-access/combo-box-property-auto-expand-yes-doesnt-seem-to/05fa61af-853e-4c9d-a3e3-2f51aa094668
    ' -------------------------------------------------------------------------------------
    'cbxModalSedSize.RowSource = GetTemplate("s_enums_for_type", "etype" & PARAM_SEPARATOR & "ModWentworthClassSize")
    'use default year for scale (set w/in GetRecords)
    Set cbxModalSedSize.Recordset = GetRecords("s_mod_wentworth_for_eventyr")
    cbxModalSedSize.BoundColumn = 1 'bind to label (not ID)
    cbxModalSedSize.ColumnCount = 6
    cbxModalSedSize.ColumnWidths = "0;1.3in;.4in;0;0;0" 'display the display column (combines label - summary)
    
    'ID default -> value used only for edits of existing table values
    tbxID.Value = 0
  
    'defaults --> always on items
    '% sand & finer, water, litter, woody debris, modal sediment cover,
    '% standing dead, filamentous algae, beaver browse
    tglBeaverBrowse.Visible = True
    lblBeaverBrowse.Visible = True
  
    'defaults --> turn off items
    lblTransect.Visible = False
    cbxTransect.Visible = False
    lblNumber.Visible = False
    tbxNumber.Visible = False
    lblDistance.Visible = False
    tbxDistance.Visible = False
    lblWCC.Visible = False
    lblURC.Visible = False
    lblARC.Visible = False
    tbxPctWCC.Visible = False
    tbxPctURC.Visible = False
    tbxPctARC.Visible = False
    tglNoCanopyVeg.Visible = False
    lblNoCanopyVeg.Visible = False
    tglNoIndicatorSpecies.Visible = False
    lblNoIndicatorSpecies.Visible = False
    tglNoRootedVeg.Visible = False
    lblNoRootedVeg.Visible = False
    'tglHasSocialTrails.Visible = False  'replaced by tbxPctSocialTrails
    'lblHasSocialTrails.Visible = False  'replaced by lblPctSocialTrails
    tbxPctSocialTrails.Visible = False
    lblPctSocialTrails.Visible = False
    btnAddTransect.Visible = False
    btnWCC.Visible = False
    btnURC.Visible = False
    btnARC.Visible = False
    lblTagline.Visible = False
    btnTaglines.Visible = False
    lblPlotDensity.Visible = False
    lblFraction.Visible = False
    tbxPlotDensity.Visible = False
    lblPlotDensityBgd.Visible = False
    
    'default plot density = 3 starting in 2015 (i.e. 1/3 density)
    tbxPlotDensity.Value = 3
    
    'adjust UI based on park
    Select Case TempVars("ParkCode")
        
        Case "BLCA"     'WCC & URC, NoCanopyVeg, NoIndicatorSpecies
            lblTransect.Visible = True
            cbxTransect.Visible = True
            lblWCC.Visible = True
            tbxPctWCC.Visible = True
            lblURC.Visible = True
            tbxPctURC.Visible = True
            lblDistance.Visible = True
            tbxDistance.Visible = True
            lblPlotDensity.Visible = True
            lblFraction.Visible = True
            tbxPlotDensity.Visible = True
            lblPlotDensityBgd.Visible = True
            tglNoCanopyVeg.Visible = True
            lblNoCanopyVeg.Visible = True
            tglNoIndicatorSpecies.Visible = True
            lblNoIndicatorSpecies.Visible = True
'            tglHasSocialTrails.Visible = True
'            lblHasSocialTrails.Visible = True
            tbxPctSocialTrails.Visible = True
            lblPctSocialTrails.Visible = True
            btnAddTransect.Visible = True
            btnWCC.Visible = True
            btnURC.Visible = True
            lblTagline.Visible = True
            btnTaglines.Visible = True
            
        Case "CANY"     'WCC & ARC, NoCanopyVeg, NoRootedVeg
            lblTransect.Visible = True
            cbxTransect.Visible = True
            lblWCC.Visible = True
            tbxPctWCC.Visible = True
            lblARC.Visible = True
            tbxPctARC.Visible = True
            lblDistance.Visible = True
            tbxDistance.Visible = True
            'lblPlotDensity.Visible = True
            'lblFraction.Visible = True
            'tbxPlotDensity.Visible = True
            'lblPlotDensityBgd.Visible = True
            tglNoCanopyVeg.Visible = True
            lblNoCanopyVeg.Visible = True
            tglNoRootedVeg.Visible = True
            lblNoRootedVeg.Visible = True
            btnAddTransect.Visible = True
            btnWCC.Visible = True
            btnARC.Visible = True
            
        Case "DINO"     'ARC, NoRootedVeg, No Transects
                        'WCC, % Standing Dead (2018 & after)
            lblARC.Visible = True
            tbxPctARC.Visible = True
            lblNumber.Visible = True
            tbxNumber.Visible = True
            tglNoRootedVeg.Visible = True
            lblNoRootedVeg.Visible = True
            btnARC.Visible = True
            
            '2018 & after only
            btnWCC.Visible = IIf(EventYear > 2017, True, False)
    
    End Select
    
    'ID default -> value used only for edits of existing table values
    tbxID.DefaultValue = 0
    
    'initialize values
    ClearForm Me
  
    'filter & order subform
'    With list.Form
'        .Filter = "Site_ID = " & TempVars("SiteID")
'        .FilterOn = True
'        .FilterOnLoad = True
'        .OrderBy = "ID DESC"
'        .OrderByOn = True
'        .OrderByOnLoad = True
'    End With
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[VegPlot form])"
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
            "Error encountered (#" & Err.Number & " - Form_Load[VegPlot form])"
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
              
    If tbxID > 0 Then btnComment.Enabled = True
    btnSetObserverRecorder.Enabled = IIf(tbxID.Value > 0, True, False)

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxEvent_AfterUpdate
' Description:  Combobox after update actions
' Assumptions:  Event combobox contains the following columns:
'                   column(0)= event ID                 column(3)= event date - site name (sitecode)
'                   column(1)= event date               column(4)= site code
'                   column(2)= event date - site code   column(5)= park code
'               Column 1 (event date) will be used to determine the proper MSS year
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 9, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/9/2017 - initial version
'   BLC - 11/7/2017 - fix cbxEvent column for date
'   BLC - 12/27/2017 - update related ID value
' ---------------------------------
Private Sub cbxEvent_AfterUpdate()
On Error GoTo Err_Handler

    'enable modal sediment size based on event year
    'column(1)= event date
    SetTempVar "EventYear", Year(cbxEvent.Column(1))
    Me.cbxModalSedSize.Enabled = True
    
    'update modal sed size classes
    Set cbxModalSedSize.Recordset = GetRecords("s_mod_wentworth_for_eventyr")
    cbxModalSedSize.Requery
    
    'updated related ID
    tbxEventID = cbxEvent.Value
    
    'check if ready
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxEvent_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxTransect_AfterUpdate
' Description:  Combobox after update actions
' Assumptions:  Transect combobox contains the following columns:
'                   column(0)= transect ID              column(3)= event date - site name (sitecode)
'                   column(1)= event date               column(4)= site code
'                   column(2)= event date - site code   column(5)= park code
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 27, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/27/2017 - initial version
' ---------------------------------
Private Sub cbxTransect_AfterUpdate()
On Error GoTo Err_Handler
    
    'updated related ID
    tbxTransectID = cbxTransect.Value
    
    'check if ready
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTransect_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxNumber_AfterUpdate
' Description:  Textbox after update actions
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
Private Sub tbxNumber_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxNumber.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNumber_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxDistance_AfterUpdate
' Description:  Textbox after update actions
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
Private Sub tbxDistance_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxDistance.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDistance_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxModalSedSize_AfterUpdate
' Description:  Combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
'   BLC - 7/13/2016 - revised to combobox
'   BLC - 12/27/2017 - update related ID value
' ---------------------------------
Private Sub cbxModalSedSize_AfterUpdate()
On Error GoTo Err_Handler

    If Len(cbxModalSedSize.Text) > 0 Then
        'update ID
        tbxMSSID = cbxModalSedSize.Value
        
        ReadyForSave
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxModalSedSize_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctURC_AfterUpdate
' Description:  Textbox after update actions
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
Private Sub tbxPctURC_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctURC.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctURC_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctFines_AfterUpdate
' Description:  Textbox after update actions
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
Private Sub tbxPctFines_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctFines.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctFines_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctWater_AfterUpdate
' Description:  Textbox after update actions
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
Private Sub tbxPctWater_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctWater.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctWater_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctFA_AfterUpdate
' Description:  Textbox after update actions
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
Private Sub tbxPctFA_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctFA.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctFA_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctLitter_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 12, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/12/2017 - initial version
' ---------------------------------
Private Sub tbxPctLitter_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctLitter.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctLitter_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctWoodyDebris_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 12, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/12/2017 - initial version
' ---------------------------------
Private Sub tbxPctWoodyDebris_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctWoodyDebris.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctWoodyDebris_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctWCC_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' SoWCCe/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub tbxPctWCC_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctWCC.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctWCC_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctStandingDead_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' SoPctStandingDeade/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub tbxPctStandingDead_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctStandingDead.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctStandingDead_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctSocialTrails_AfterUpdate
' Description:  Textbox after update actions
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
Private Sub tbxPctSocialTrails_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctSocialTrails.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctSocialTrails_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPlotDensity_AfterUpdate
' Description:  Textbox after update actions
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
Private Sub tbxPlotDensity_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPlotDensity.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPlotDensity_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglNoCanopyVeg_AfterUpdate
' Description:  Toggle button after update actions
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
Private Sub tglNoCanopyVeg_AfterUpdate()
On Error GoTo Err_Handler

    'display as checkbox
    ToggleCaption tglNoCanopyVeg, True
    
    If tglNoCanopyVeg > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglNoCanopyVeg_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglNoRootedVeg_AfterUpdate
' Description:  Toggle button after update actions
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
Private Sub tglNoRootedVeg_AfterUpdate()
On Error GoTo Err_Handler

    'display as checkbox
    ToggleCaption tglNoRootedVeg, True
    
    If tglNoRootedVeg Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglNoRootedVeg_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglNoIndicatorSpecies_AfterUpdate
' Description:  Toggle button after update actions
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
Private Sub tglNoIndicatorSpecies_AfterUpdate()
On Error GoTo Err_Handler

    'display as checkbox
    ToggleCaption tglNoIndicatorSpecies, True
    
    If tglNoIndicatorSpecies Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglNoIndicatorSpecies_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglBeaverBrowse_AfterUpdate
' Description:  Toggle after update actions
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
Private Sub tglBeaverBrowse_AfterUpdate()
On Error GoTo Err_Handler
    
    'display as checkbox
    ToggleCaption tglBeaverBrowse, True
    
    If tglBeaverBrowse Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglBeaverBrowse_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

'' ---------------------------------
'' Sub:          tglHasSocialTrails_AfterUpdate
'' Description:  Toggle after update actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 6/27/2016 - initial version
''   BLC - 1/11/2017 - revised to toggle w/ text change
'' ---------------------------------
'Private Sub tglHasSocialTrails_AfterUpdate()
'On Error GoTo Err_Handler
'
''    If tglHasSocialTrails Then
''        tglHasSocialTrails.Caption = StringFromCodepoint(uCheck)
''        ReadyForSave
''    Else
''        tglHasSocialTrails.Caption = ""
''    End If
'
'    'display as checkbox
'    ToggleCaption tglHasSocialTrails, True
'
'    If tglHasSocialTrails Then _
'        ReadyForSave
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tglHasSocialTrails_AfterUpdate[VegPlot form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' ---------------------------------
' Sub:          chkCalibrationPlot_Click
' Description:  Checkbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 1, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/1/2017 - initial version
'   BLC - 11/26/2017 - display plot # hint when calibration/replicate
'   BLC - 12/27/2017 - updated IF to avoid error #13 - type mismatch
' ---------------------------------
Private Sub chkCalibrationPlot_Click()
On Error GoTo Err_Handler
    
    'default
    lblHintPlotNumDist.Visible = False
    
    'unhide hint if calibration/replicate plot
    If chkCalibrationPlot = True Or chkReplicatePlot = True Then _
        lblHintPlotNumDist.Visible = True
    
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkCalibrationPlot_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          chkReplicatePlot_Click
' Description:  Checkbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 1, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/1/2017 - initial version
'   BLC - 11/26/2017 - display plot # hint when calibration/replicate
'   BLC - 12/27/2017 - updated IF to avoid error #13 - type mismatch
' ---------------------------------
Private Sub chkReplicatePlot_Click()
On Error GoTo Err_Handler
    
    'default
    lblHintPlotNumDist.Visible = False
    
    'unhide hint if calibration/replicate plot
    If chkCalibrationPlot = True Or chkReplicatePlot = True Then _
    lblHintPlotNumDist.Visible = True
    
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkReplicatePlot_Click[VegPlot form])"
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
'   BLC - 6/27/2016 - revised to use ClearForm()
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
            "Error encountered (#" & Err.Number & " - btnUndo_Click[VegPlot form])"
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
'   BLC - 1/11/2017 - revised checkboxes to toggle buttons &
'                     converted values to 1/0 vs. -1/0 for SQL clarity
'   BLC - 1/12/2017 - code cleanup, enabled buttons after tbxID > 0
'                     (plot saved & ID returned)
'   BLC - 9/25/2017 - revise for NCPN_framework.XX classes
'   BLC - 9/27/2017 - update to use Factory.NewClassXX() vs GetClass()
'   BLC - 11/10/2017 - update to handle unset values using Ne()
'   BLC - 11/11/2017 - update percent values
'   BLC - 11/26/2017 - replaced HasSocialTrails w/ PctSocialTrails
'   BLC - 12/5/2017 - add BeaverBrowse, Calibration & Replicate Plots
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
'    'Dim vp As New VegPlot
'    Dim vp As NCPN_framework.VegPlot
'    Set vp = Factory.NewVegPlot
'
'    With vp
'        'values passed into form
'        .SiteID = TempVars("SiteID")
'        .FeatureID = GetFeatureID(TempVars("ParkCode"), TempVars("FeatureID"))
'
'        'form values
'        .EventID = cbxEvent
'        .VegTransectID = Ne(Nz(cbxTransect, 0), 0)
'        .PlotNumber = Ne(tbxNumber, 0)
'        .PlotDistance = Ne(tbxDistance, 0)
'        .ModalSedimentSizeID = cbxModalSedSize
'
'        .PlotDensity = Ne(tbxPlotDensity, 0)
'
'        'pct values
'        .PctFines = SetTrace(tbxPctFines, 0.5)
'        .PctWater = SetTrace(tbxPctWater, 0.5)
'        .PctLitter = SetTrace(tbxPctLitter, 0.5)
'        .PctWoodyDebris = SetTrace(tbxPctWoodyDebris, 0.5)
'        .PctFilamentousAlgae = SetTrace(tbxPctFA, 0.5)
'        .PctStandingDead = SetTrace(tbxPctStandingDead, 0.5)
'        .PctSocialTrails = SetTrace(tbxPctSocialTrails, 0.5)
'        .PctModalSedimentSize = SetTrace(tbxPctMSS, 0.5)
'        .WoodyCanopyPctCover = SetTrace(Ne(tbxPctWCC, 0), 0.5)
'        .UnderstoryRootedPctCover = SetTrace(Ne(tbxPctURC, 0), 0.5)
'        .AllRootedPctCover = SetTrace(Ne(tbxPctARC, 0), 0.5)
'
'        'chk/tgl values -> change Access -1 (true) to clearer 1 for use in SQL
'        '                  so value of 1 = has no canopy veg, 0 = has canopy veg etc.
'        .NoCanopyVeg = IIf(tglNoCanopyVeg = -1, 1, 0)
'        .NoRootedVeg = IIf(tglNoRootedVeg = -1, 1, 0)
'        .NoIndicatorSpecies = IIf(tglNoIndicatorSpecies = -1, 1, 0)
'        '.HasSocialTrails = IIf(tglHasSocialTrails = -1, 1, 0) 'replaced w/ PctSocialTrails
'        .BeaverBrowse = IIf(tglBeaverBrowse = -1, 1, 0)
'
'        .CalibrationPlot = IIf(chkCalibrationPlot = True, 1, 0)
'        .ReplicatePlot = IIf(chkReplicatePlot = True, 1, 0)
'
'        .ID = tbxID '0 if new, edit if > 0
'        .SaveToDb
'
'        'set the tbxID.value
'        tbxID = .ID
'
'    End With
'
'    'clear values & refresh display
'
'    ReadyForSave
'
'    PopulateForm Me, tbxID.Value
'
'    If tbxID.Value > 0 Then
'        'highlight SetObserverRecorder button
'        btnSetObserverRecorder.BorderColor = lngYellow
'        lblMsg.ForeColor = lngYellow
'        lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
'        lblMsg.Caption = "Don't forget to set observer && recorder!"
'
'        'enable buttons
'        btnSubstrateCover.Enabled = True
'        btnWCC.Enabled = True
'        btnURC.Enabled = True
'        btnARC.Enabled = True
'        btnTaglines.Enabled = True
'    End If
'    'refresh list
'    Me.list.Requery
'
'    Me.Requery
    
    'set enable btnSave_Click save
    m_SaveOK = True
        
    UpsertRecord Me
    
    Me![list].Form.Requery
    
    'revert to disable non-btnSave_Click save
    m_SaveOK = False
    
    'clear fields
    ClearForm Me
    
    'clear control sources
    cbxEvent.ControlSource = ""
    'cbxLocation.Value = ""
    cbxTransect.ControlSource = ""
    cbxModalSedSize.ControlSource = ""
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSubstrateCover_Click
' Description:  Substrate cover button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 9, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/9/2017 - initial version
' ---------------------------------
Private Sub btnSubstrateCover_Click()
On Error GoTo Err_Handler
    
    'open substrate cover form
    DoCmd.OpenForm "SubstrateCover", acNormal, , , , , "VegPlot|" & tbxID.Text
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSubstrateCover_Click[VegWalk form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSetObserverRecorder_Click
' Description:  Set observer recorder button click actions
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
Private Sub btnSetObserverRecorder_Click()
On Error GoTo Err_Handler
    
    DoCmd.OpenForm "SetObserverRecorder", acNormal, , , , , "VegPlot|" & Me.tbxID & "|" & Me.cbxEvent.Column(1) 'Column 1 = event date

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSetObserverRecorder_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnTaglines_Click
' Description:  Tagline button click actions
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
Private Sub btnTaglines_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "Tagline", acNormal, , , , , Me.Name & "|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnTaglines_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnWCC_Click
' Description:  Woody Canopy Cover button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 11/26/2017 - revised to call VegSpecies vs VegWalk
' ---------------------------------
Private Sub btnWCC_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "VegSpecies", acNormal, , , , , "WoodyCanopySpecies|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnWCC_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnURC_Click
' Description:  Woody Canopy Cover button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 11/26/2017 - revised to call VegSpecies vs VegWalk
' ---------------------------------
Private Sub btnURC_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "VegSpecies", acNormal, , , , , "UnderstoryRootedSpecies|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnURC_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnARC_Click
' Description:  Woody Canopy Cover button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 11/26/2017 - revised to call VegSpecies vs VegWalk
' ---------------------------------
Private Sub btnARC_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "VegSpecies", acNormal, , , , , "AllRootedSpecies|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnARC_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAddEvent_Click
' Description:  Add event button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 2, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/2/2016 - initial version
' ---------------------------------
Private Sub btnAddEvent_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "Events", acNormal, , , , , Me.Name
    
    'refresh cbx (done from event form close)
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddEvent_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAddTransect_Click
' Description:  Add transect button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 11, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/11/2017 - initial version
'   BLC - 1/12/2017 - revised form name to VegTransect vs. Transect
' ---------------------------------
Private Sub btnAddTransect_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "VegTransect", acNormal, , , , , Me.Name & "|" & tbxID
    
    'refresh cbx (done from transect form close)
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddTransect_Click[VegPlot form])"
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
'   BLC - 11/10/2017 - fix error 2185(can't reference property/method unless control has focus)
'                      use tbxID vs tbxID.Text
' ---------------------------------
Private Sub btnComment_Click()
On Error GoTo Err_Handler
    
    'open comment form
    DoCmd.OpenForm "Comment", acNormal, , , , , "VegPlot|" & tbxID & "|255"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[VegPlot form])"
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
'   BLC - 10/24/2016 - revised to use CallingForm property
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'clear subform filter & order
'    With list.Form
'        .Filter = ""
'        .OrderBy = ""
'    End With

    'restore CallingForm
    ToggleForm Me.CallingForm, 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[VegPlot form])"
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
'   BLC - 1/9/2017  - adjusted for park specific modifications, substrate cover
'   BLC - 1/11/2017 - adjusted for toggle buttons vs. checkboxes
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    'requires:  EventID, SiteID, FeatureID, VegTransectID, PlotDistance_m,
    '           ModalSedSize, PctFine, PctWater, PctURC, PlotDensity,
    '           NoCanopyVeg, NoRootedVeg, HasSocialTrail, FA
    '           BLCA only: NoIndicatorSpecies
    
'    If Nz(tbxDistance.value, "") > 0 _
'        And Nz(cbxModalSedSize.value, "") > -1 _
'        And Nz(tbxPctFines.value, "") > -1 _
'        And Nz(tbxPctWater.value, "") > -1 _
'        And Nz(tbxPctURC.value, "") > -1 _
'        And Nz(tbxPlotDensity.value, "") > -1 _
'        And Nz(chkNoCanopyVeg.value, "") > -1 _
'        And Nz(chkNoRootedVeg.value, "") > -1 _
'        And Nz(chkHasSocialTrails.value, "") > -1 Then
    
'       And Nz(tbxPctFA.Value,"") > -1 _

    If Nz(tbxDistance.Value, "") > 0 _
        And Nz(cbxModalSedSize.Value, "") > -1 _
        And Nz(tbxPctFines.Value, "") > -1 _
        And Nz(tbxPctWater.Value, "") > -1 Then '_
'        And Nz(chkNoRootedVeg.Value, "") > -1 Then
        
        Select Case TempVars("ParkCode")
            Case "BLCA"
                'requires NoIndicatorSpecies
                'If Nz(chkNoIndicatorSpecies.Value, "") > -1 Then GoTo Exit_Handler

            Case "CANY"
            Case "DINO"
        End Select
        
        isOK = True
        
    End If
    
    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    'enable save button only for new sites (tbxID = 0)
    If tbxID = 0 Then btnSave.Enabled = isOK
    
    btnSubstrateCover.Enabled = IIf(tbxID.Value > 0, True, False)
    btnSetObserverRecorder.Enabled = IIf(tbxID.Value > 0, True, False)
    
    'refresh form
    Me.Requery
   
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub
