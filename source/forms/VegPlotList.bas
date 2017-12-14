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
    ItemSuffix =37
    Left =1065
    Top =6570
    Right =8445
    Bottom =10065
    DatasheetGridlinesColor =14806254
    OrderBy ="ID DESC"
    RecSrcDt = Begin
        0xe8a9c07c72c3e440
    End
    RecordSource ="VegPlot"
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
    FilterOnLoad =255
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
                    OverlapFlags =215
                    Left =180
                    Top =120
                    Width =7260
                    Height =840
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Edit or Delete Records using the buttons for the record at right.\015\012Icon co"
                        "des at left identify if record may be edited/deleted."
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =960
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =1020
                    Width =1860
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblPlotNumDist"
                    Caption ="Plot #/Distance (m)"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =840
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
                    LayoutCachedLeft =840
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1110
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =3120
                    Top =1020
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblPctMSS"
                    Caption ="MSS % Cover"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =3120
                    LayoutCachedTop =1020
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4500
                    Top =1020
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblModalSedSize"
                    Caption ="MSS Class"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =4500
                    LayoutCachedTop =1020
                    LayoutCachedWidth =5745
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =420
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1080
                    Width =1560
                    Height =300
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxReplicatePlot"
                    ControlSource ="ReplicatePlot"
                    GridlineColor =10921638

                    LayoutCachedLeft =1080
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Width =1560
                    Height =300
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxCalibrationPlot"
                    ControlSource ="CalibrationPlot"
                    GridlineColor =10921638

                    LayoutCachedLeft =2700
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5940
                    Width =1560
                    Height =300
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxDistance"
                    ControlSource ="PlotDistance_m"
                    GridlineColor =10921638

                    LayoutCachedLeft =5940
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4320
                    Width =1560
                    Height =300
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxNumber"
                    ControlSource ="PlotNumber"
                    GridlineColor =10921638

                    LayoutCachedLeft =4320
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =6000
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

                    LayoutCachedLeft =6000
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =360
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
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =15
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =15
                    LayoutCachedWidth =840
                    LayoutCachedHeight =315
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin CommandButton
                    OverlapFlags =247
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
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Width =480
                    Height =315
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxID"
                    ControlSource ="ID"
                    ConditionalFormat = Begin
                        0x01000000e0000000020000000100000000000000000000002000000001000000 ,
                        0xff990000ffffff000100000000000000210000003f0000000100000022b14c00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b00740062007800430061006c0069006200720061007400 ,
                        0x69006f006e0050006c006f0074005d003d0031002c0031002c00300029000000 ,
                        0x000049004900660028005b007400620078005200650070006c00690063006100 ,
                        0x7400650050006c006f0074005d003d0031002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =720
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =315
                    ConditionalFormat14 = Begin
                        0x010002000000010000000000000001000000ff990000ffffff001f0000004900 ,
                        0x4900660028005b00740062007800430061006c00690062007200610074006900 ,
                        0x6f006e0050006c006f0074005d003d0031002c0031002c003000290000000000 ,
                        0x000000000000000000000000000000000001000000000000000100000022b14c ,
                        0x00ffffff001d00000049004900660028005b007400620078005200650070006c ,
                        0x006900630061007400650050006c006f0074005d003d0031002c0031002c0030 ,
                        0x002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4200
                    Width =1560
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxModalSedSize"
                    ControlSource ="ModalSedimentSize_ID"
                    ConditionalFormat = Begin
                        0x01000000e0000000020000000100000000000000000000002000000001000000 ,
                        0xff990000ffffff000100000000000000210000003f0000000100000022b14c00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b00740062007800430061006c0069006200720061007400 ,
                        0x69006f006e0050006c006f0074005d003d0031002c0031002c00300029000000 ,
                        0x000049004900660028005b007400620078005200650070006c00690063006100 ,
                        0x7400650050006c006f0074005d003d0031002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4200
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010002000000010000000000000001000000ff990000ffffff001f0000004900 ,
                        0x4900660028005b00740062007800430061006c00690062007200610074006900 ,
                        0x6f006e0050006c006f0074005d003d0031002c0031002c003000290000000000 ,
                        0x000000000000000000000000000000000001000000000000000100000022b14c ,
                        0x00ffffff001d00000049004900660028005b007400620078005200650070006c ,
                        0x006900630061007400650050006c006f0074005d003d0031002c0031002c0030 ,
                        0x002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1440
                    Width =1560
                    Height =300
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxPlotNumDist"
                    ControlSource ="=IIf([tbxCalibrationPlot]=1,\"C\",\"\") & IIf([tbxReplicatePlot]=1,\"R\",\"\") &"
                        " IIf(TempVars(\"ParkCode\")=\"DINO\",[tbxNumber],[tbxDistance])"
                    ConditionalFormat = Begin
                        0x01000000e0000000020000000100000000000000000000002000000001000000 ,
                        0xff990000ffffff000100000000000000210000003f0000000100000022b14c00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b00740062007800430061006c0069006200720061007400 ,
                        0x69006f006e0050006c006f0074005d003d0031002c0031002c00300029000000 ,
                        0x000049004900660028005b007400620078005200650070006c00690063006100 ,
                        0x7400650050006c006f0074005d003d0031002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010002000000010000000000000001000000ff990000ffffff001f0000004900 ,
                        0x4900660028005b00740062007800430061006c00690062007200610074006900 ,
                        0x6f006e0050006c006f0074005d003d0031002c0031002c003000290000000000 ,
                        0x000000000000000000000000000000000001000000000000000100000022b14c ,
                        0x00ffffff001d00000049004900660028005b007400620078005200650070006c ,
                        0x006900630061007400650050006c006f0074005d003d0031002c0031002c0030 ,
                        0x002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3120
                    Width =840
                    Height =300
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxPctMSS"
                    ControlSource ="PctModalSedimentSize"
                    ConditionalFormat = Begin
                        0x01000000e0000000020000000100000000000000000000002000000001000000 ,
                        0xff990000ffffff000100000000000000210000003f0000000100000022b14c00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b00740062007800430061006c0069006200720061007400 ,
                        0x69006f006e0050006c006f0074005d003d0031002c0031002c00300029000000 ,
                        0x000049004900660028005b007400620078005200650070006c00690063006100 ,
                        0x7400650050006c006f0074005d003d0031002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3120
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010002000000010000000000000001000000ff990000ffffff001f0000004900 ,
                        0x4900660028005b00740062007800430061006c00690062007200610074006900 ,
                        0x6f006e0050006c006f0074005d003d0031002c0031002c003000290000000000 ,
                        0x000000000000000000000000000000000001000000000000000100000022b14c ,
                        0x00ffffff001d00000049004900660028005b007400620078005200650070006c ,
                        0x006900630061007400650050006c006f0074005d003d0031002c0031002c0030 ,
                        0x002900000000000000000000000000000000000000000000
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
' Form:         VegPlotList
' Level:        Application form
' Version:      1.03
' Basis:        Dropdown form
'
' Description:  List form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
'               BLC - 10/16/2017 - 1.01 - revised to use tbxID vs. ID on delete
'               BLC - 11/10/2017 - 1.02 - fix so passes form Me.Parent.Form vs. Me.Parent (btnEdit_Click())
'               BLC - 12/7/2017 - 1.03 - added sorting by header
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
    
    'populate Modal Sediment Size

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[VegPlotList form])"
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
            "Error encountered (#" & Err.Number & " - Form_Load[VegPlotList form])"
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
            "Error encountered (#" & Err.Number & " - Form_Current[VegPlotList form])"
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
'   BLC - 12/7/2016 - initial version
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
            "Error encountered (#" & Err.Number & " - lblHdrID_Click[VegPlotList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblPctMSS_Click
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
'   BLC - 12/7/2016 - initial version
' ---------------------------------
Private Sub lblPctMSS_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblPctMSS

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblPctMSS_Click[VegPlotList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblModalSedSize_Click
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
'   BLC - 12/7/2016 - initial version
' ---------------------------------
Private Sub lblModalSedSize_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblModalSedSize

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblModalSedSize_Click[VegPlotList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblPlotNumDist_Click
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
'   BLC - 12/7/2016 - initial version
' ---------------------------------
Private Sub lblPlotNumDist_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblPlotNumDist

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblPlotNumDist_Click[VegPlotList form])"
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
'   BLC - 11/10/2017 - fix so passes form Me.Parent.Form vs. Me.Parent
' ---------------------------------
Private Sub btnEdit_Click()
On Error GoTo Err_Handler
    
    'populate the parent form
    PopulateForm Me.Parent.Form, ID

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEdit_Click[VegPlotList form])"
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
'   BLC - 10/17/2017 - revised to use tbxID vs. ID on delete
' ---------------------------------
Private Sub btnDelete_Click()
On Error GoTo Err_Handler
    
    Dim result As Integer
    
    'identify the record ID
     result = MsgBox("Delete Record this record: #" & tbxID & " ?" _
                        & vbCrLf & "This action cannot be undone.", vbYesNo, "Delete Record?")

    If result = vbYes Then DeleteRecord "VegPlot", tbxID
    
    'clear the deleted record
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[VegPlotList form])"
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
            "Error encountered (#" & Err.Number & " - Form_Close[VegPlotList form])"
    End Select
    Resume Exit_Handler
End Sub
