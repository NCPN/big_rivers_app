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
    ItemSuffix =34
    Right =13155
    Bottom =11385
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x5dbc32126cd1e440
    End
    RecordSource ="Flags"
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
        Begin FormHeader
            Height =1530
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
                    OverlapFlags =93
                    Left =2520
                    Top =1215
                    Width =555
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblAlphaCode"
                    Caption ="Alpha"
                    GridlineColor =10921638
                    LayoutCachedLeft =2520
                    LayoutCachedTop =1215
                    LayoutCachedWidth =3075
                    LayoutCachedHeight =1530
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =960
                    Top =1215
                    Width =270
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblHdrID"
                    Caption ="ID"
                    GridlineColor =10921638
                    LayoutCachedLeft =960
                    LayoutCachedTop =1215
                    LayoutCachedWidth =1230
                    LayoutCachedHeight =1530
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =3960
                    Top =1215
                    Width =1440
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFlagName"
                    Caption ="Name"
                    GridlineColor =10921638
                    LayoutCachedLeft =3960
                    LayoutCachedTop =1215
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =1530
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3300
                    Top =1215
                    Width =555
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblNumericCode"
                    Caption ="#"
                    GridlineColor =10921638
                    LayoutCachedLeft =3300
                    LayoutCachedTop =1215
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =1530
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =2460
                    Top =840
                    Width =1380
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblCode"
                    Caption ="Flag Code"
                    GridlineColor =10921638
                    LayoutCachedLeft =2460
                    LayoutCachedTop =840
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =1155
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =1380
                    Top =1215
                    Width =600
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblGroup"
                    Caption ="Grp"
                    GridlineColor =10921638
                    LayoutCachedLeft =1380
                    LayoutCachedTop =1215
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =1530
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =1920
                    Top =1200
                    Width =600
                    Height =300
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblType"
                    Caption ="Type"
                    GridlineColor =10921638
                    LayoutCachedLeft =1920
                    LayoutCachedTop =1200
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =1500
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Line
                    OverlapFlags =87
                    Left =2460
                    Top =1155
                    BorderColor =16777062
                    Name ="lnFlag"
                    GridlineColor =10921638
                    LayoutCachedLeft =2460
                    LayoutCachedTop =1155
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =1155
                    BorderThemeColorIndex =-1
                End
            End
        End
        Begin Section
            Height =480
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =6000
                    Top =60
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
                    LayoutCachedTop =60
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =420
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
                    Top =75
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =75
                    LayoutCachedWidth =840
                    LayoutCachedHeight =375
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
                    Top =60
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
                    LayoutCachedTop =60
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =420
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
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2460
                    Top =60
                    Width =660
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxSiteCode"
                    ControlSource ="Code"
                    ConditionalFormat = Begin
                        0x01000000c4000000020000000100000000000000000000001800000001000000 ,
                        0x22b14c00ffffff000100000000000000190000003100000001000000ff000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004900730041006300740069007600650046006f007200500072006f007400 ,
                        0x6f0063006f006c005d003d003100000000005b00490073004100630074006900 ,
                        0x7600650046006f007200500072006f0074006f0063006f006c005d003d003000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2460
                    LayoutCachedTop =60
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000022b14c00ffffff00170000005b00 ,
                        0x4900730041006300740069007600650046006f007200500072006f0074006f00 ,
                        0x63006f006c005d003d0031000000000000000000000000000000000000000000 ,
                        0x00010000000000000001000000ff000000ffffff00170000005b004900730041 ,
                        0x006300740069007600650046006f007200500072006f0074006f0063006f006c ,
                        0x005d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Width =480
                    Height =315
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =900
                    LayoutCachedTop =60
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1440
                    Top =60
                    Width =420
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFlagGroupAbbrev"
                    ControlSource ="FlagGroup"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000002000000000000000200000001000000 ,
                        0x22b14c00ffffff000000000002000000030000000500000001000000ed1c2400 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedTop =60
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =375
                    ConditionalFormat14 = Begin
                        0x01000200000000000000020000000100000022b14c00ffffff00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x00ed1c2400ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3960
                    Top =60
                    Width =1980
                    Height =315
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxSiteName"
                    ControlSource ="FlagName"
                    ConditionalFormat = Begin
                        0x01000000c4000000020000000100000000000000000000001800000001000000 ,
                        0x22b14c00ffffff000100000000000000190000003100000001000000ed1c2400 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004900730041006300740069007600650046006f007200500072006f007400 ,
                        0x6f0063006f006c005d003d003100000000005b00490073004100630074006900 ,
                        0x7600650046006f007200500072006f0074006f0063006f006c005d003d003000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3960
                    LayoutCachedTop =60
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000022b14c00ffffff00170000005b00 ,
                        0x4900730041006300740069007600650046006f007200500072006f0074006f00 ,
                        0x63006f006c005d003d0031000000000000000000000000000000000000000000 ,
                        0x00010000000000000001000000ed1c2400ffffff00170000005b004900730041 ,
                        0x006300740069007600650046006f007200500072006f0074006f0063006f006c ,
                        0x005d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3240
                    Top =60
                    Width =600
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="Text29"
                    ControlSource ="NumericCode"
                    ConditionalFormat = Begin
                        0x01000000c4000000020000000100000000000000000000001800000001000000 ,
                        0x22b14c00ffffff000100000000000000190000003100000001000000ff000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004900730041006300740069007600650046006f007200500072006f007400 ,
                        0x6f0063006f006c005d003d003100000000005b00490073004100630074006900 ,
                        0x7600650046006f007200500072006f0074006f0063006f006c005d003d003000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =60
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000022b14c00ffffff00170000005b00 ,
                        0x4900730041006300740069007600650046006f007200500072006f0074006f00 ,
                        0x63006f006c005d003d0031000000000000000000000000000000000000000000 ,
                        0x00010000000000000001000000ff000000ffffff00170000005b004900730041 ,
                        0x006300740069007600650046006f007200500072006f0074006f0063006f006c ,
                        0x005d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =420
                    Height =315
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFlagTypeAbbrev"
                    ControlSource ="FlagType"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000002000000000000000200000001000000 ,
                        0x22b14c00ffffff000000000002000000030000000500000001000000ed1c2400 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =375
                    ConditionalFormat14 = Begin
                        0x01000200000000000000020000000100000022b14c00ffffff00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x00ed1c2400ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
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
' Form:         FlagList
' Level:        Application form
' Version:      1.02
' Basis:        Dropdown form
'
' Description:  List form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
'               BLC - 6/28/2016 - 1.01 - revised to use ToggleIsActive (mod_App_Data)
'               BLC - 10/16/2017 - 1.02 - fixed to use tbxID vs. ID on delete
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
    tbxIcon.Value = "" 'StringFromCodepoint(uLocked)
    tbxIcon.ForeColor = lngDkGreen
    lblDirections.ForeColor = lngLtBlue
    
    'set hover
    btnEdit.HoverColor = lngGreen
    btnDelete.HoverColor = lngGreen
    
    btnDelete.Caption = StringFromCodepoint(uDelete)
    btnDelete.ForeColor = lngRed

    'set data source
    Set Me.Recordset = GetRecords("s_flags")
    
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[FlagList form])"
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
'    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[FlagList form])"
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
            "Error encountered (#" & Err.Number & " - Form_Current[FlagList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnActivate_Click
' Description:  Activate button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
'   BLC - 6/28/2016 - added context to ToggleIsActive()
' ---------------------------------
Private Sub btnActivate_Click()
On Error GoTo Err_Handler
    
    'identify the record ID
'    ToggleIsActive "Flag", tbxID, Abs(tbxIsActive - 1)
    
    'clear the Activated record
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnActivate_Click[FlagList form])"
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
    PopulateForm Me.Parent, tbxID

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEdit_Click[FlagList form])"
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
'   BLC - 10/16/2017 - revised to use tbxID vs. ID on delete
' ---------------------------------
Private Sub btnDelete_Click()
On Error GoTo Err_Handler
    
    Dim result As Integer
    
    'identify the record ID
     result = MsgBox("Delete Record this record: #" & tbxID & " ?" _
                        & vbCrLf & "This action cannot be undone.", vbYesNo, "Delete Record?")

    If result = vbYes Then DeleteRecord "Flag", tbxID
    
    'clear the deleted record
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[FlagList form])"
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
            "Error encountered (#" & Err.Number & " - Form_Close[FlagList form])"
    End Select
    Resume Exit_Handler
End Sub
