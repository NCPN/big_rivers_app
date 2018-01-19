Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9060
    DatasheetFontHeight =11
    ItemSuffix =43
    Left =4035
    Top =3045
    Right =13980
    Bottom =14895
    DatasheetGridlinesColor =14806254
    Filter ="CommentType='event' AND CommentType_ID=38"
    RecSrcDt = Begin
        0x0a6c31995402e540
    End
    RecordSource ="AppComment"
    Caption ="Comment List"
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
                    Left =120
                    Width =7260
                    Height =840
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Sort records by clicking the header.\015\012Last modified date color reflects if"
                        " comment is recent (in the last month)."
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =840
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2880
                    Top =1020
                    Width =315
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblCommentTypeID"
                    Caption =" ID"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =2880
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3195
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =660
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
                    LayoutCachedLeft =660
                    LayoutCachedTop =1020
                    LayoutCachedWidth =930
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =3540
                    Top =1020
                    Width =1980
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblComment"
                    Caption ="Comment"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =3540
                    LayoutCachedTop =1020
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1500
                    Top =1020
                    Width =480
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblCommentType"
                    Caption ="Type"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =1500
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =5670
                    Top =780
                    Width =900
                    Height =555
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblLastModifiedDate"
                    Caption ="Last Modified"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =5670
                    LayoutCachedTop =780
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =1335
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7500
                    Top =900
                    Width =720
                    ForeColor =4210752
                    Name ="btnOpenTable"
                    Caption ="Add Record"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open tsys_Db_Templates table"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b0a090ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff0000000000000000b0a090ffe0c8c0ffd0c0b0ffd0b8b0ffd0b8b0ff ,
                        0xc0b0a0ffc0b0a0ffc0b0a0ffc0a8a0ffc0a890ffc0a890ffb0a090ffb0a090ff ,
                        0x604830ff0000000000000000b0a090fffffffffffffffffffff8ffffd0b8b0ff ,
                        0xfff0f0fffff0e0ffffe8e0ffc0a8a0fff0d8d0fff0d8c0fff0d0b0ffb0a090ff ,
                        0x604830ff0000000000000000b0a090ffffffffffffffffffffffffffd0c0b0ff ,
                        0xfff8f0fffff0f0fffff0e0ffc0b0a0ffffe0d0fff0d8d0fff0d8c0ffc0a890ff ,
                        0x604830ff0000000000000000b0a090ffe0d0d0ffd0c8c0ffd0c0c0ffd0c0b0ff ,
                        0xd0c0b0ffd0b8b0ffd0b8b0ffc0b0a0ffc0b0a0ffc0b0a0ffc0a8a0ffc0a890ff ,
                        0x604830ff0000000000000000c0a890ffffffffffffffffffffffffffd0c8c0ff ,
                        0xfffffffffff8fffffff8f0ffd0b8b0fffff0e0ffffe8e0ffffe0d0ffc0a8a0ff ,
                        0x604830ff0000000000000000c0a8a0ffffffffffffffffffffffffffd0c8c0ff ,
                        0xfffffffffffffffffff8ffffd0b8b0fffff0f0fffff0e0ffffe8e0ffc0a8a0ff ,
                        0x604830ff0000000000000000c0b0a0ffe0d8d0ffe0d0c0ffe0d0c0ffe0c8c0ff ,
                        0xd0c8c0ffd0c8c0ffd0c0b0ffd0c0b0ffd0b8b0ffd0b8b0ffc0b0a0ffc0b0a0ff ,
                        0x604830ff0000000000000000d0b0a0ffffffffffffffffffffffffffe0d0c0ff ,
                        0xffffffffffffffffffffffffd0c0b0fffff8fffffff8f0fffff0f0ffc0b0a0ff ,
                        0x604830ff0000000000000000d0b8a0ffffffffffffffffffffffffffe0d0c0ff ,
                        0xffffffffffffffffffffffffd0c8c0fffffffffffff8fffffff8f0ffd0b8b0ff ,
                        0x604830ff0000000000000000f0a890fff0a890fff0a890fff0a880fff0a080ff ,
                        0xe09870ffe09060ffe08850ffe08050ffe07840ffe07040ffe07040ffe07040ff ,
                        0xd06030ff0000000000000000f0a890ffffc0a0ffffc0a0ffffc0a0ffffb890ff ,
                        0xffb890ffffb090ffffa880ffffa880fff0a070fff0a070fff09870fff09860ff ,
                        0xd06830ff0000000000000000f0a890fff0a890fff0a890fff0a890fff0a880ff ,
                        0xf0a080fff09870ffe09870ffe09060ffe08860ffe08050ffe07840ffe07840ff ,
                        0xe07040ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =7500
                    LayoutCachedTop =900
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =1260
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
                    TextAlign =2
                    Left =1200
                    Top =660
                    Width =2040
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblReference"
                    Caption ="Reference"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =660
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =975
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =480
            Name ="Detail"
            OnMouseMove ="[Event Procedure]"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =45
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =2
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =45
                    LayoutCachedWidth =480
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontFamily =2
                    Left =8280
                    Width =720
                    FontSize =14
                    TabIndex =1
                    ForeColor =255
                    Name ="btnDelete"
                    Caption ="í ½í·´"
                    OnClick ="[Event Procedure]"
                    FontName ="Academy Engraved LET"
                    GridlineColor =10921638

                    LayoutCachedLeft =8280
                    LayoutCachedWidth =9000
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
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3540
                    Top =45
                    Width =1980
                    Height =315
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxComment"
                    ControlSource ="=IIf(Len([Comment])>15,Truncate([Comment],15,\"L\") & \"...\",[Comment])"
                    OnMouseMove ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3540
                    LayoutCachedTop =45
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =600
                    Top =45
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =45
                    LayoutCachedWidth =960
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1050
                    Top =45
                    Width =1590
                    Height =315
                    FontSize =8
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxCommentType"
                    ControlSource ="CommentType"
                    GridlineColor =10921638

                    LayoutCachedLeft =1050
                    LayoutCachedTop =45
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5580
                    Top =30
                    Width =1800
                    Height =330
                    FontSize =8
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxLastModified"
                    ControlSource ="LastModified"
                    ConditionalFormat = Begin
                        0x0100000094000000010000000100000000000000000000001900000001000000 ,
                        0x3333ff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004c006100730074004d006f006400690066006900650064005d003e004400 ,
                        0x610074006500280029002d003300300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =5580
                    LayoutCachedTop =30
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010000003333ff00ffffff00180000005b00 ,
                        0x4c006100730074004d006f006400690066006900650064005d003e0044006100 ,
                        0x74006500280029002d0033003000000000000000000000000000000000000000 ,
                        0x000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3900
                    Top =60
                    Height =315
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="lblType"
                    ControlSource ="=Left([TemplateName],1)"
                    GridlineColor =10921638

                    LayoutCachedLeft =3900
                    LayoutCachedTop =60
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7500
                    Width =720
                    ForeColor =4210752
                    Name ="btnEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Edit comment"
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

                    LayoutCachedLeft =7500
                    LayoutCachedWidth =8220
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
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2760
                    Top =60
                    Width =480
                    Height =315
                    FontSize =8
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxCommentTypeID"
                    ControlSource ="CommentType_ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2760
                    LayoutCachedTop =60
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
            End
        End
        Begin FormFooter
            Visible = NotDefault
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
' Form:         CommentList
' Level:        Application form
' Version:      1.01
' Basis:        Dropdown form
'
' Description:  List form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, October 18, 2017
' References:   -
' Revisions:    BLC - 10/18/2017 - 1.00 - initial version
'               BLC - 11/24/2017 - 1.01 - fixed to delete from AppComment vs tsys_Db_Templates
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
        'Me.btnEdit.Caption = m_ButtonCaption
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
' Source/date:  Bonnie Campbell, October 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/18/2017 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

Debug.Print Me.OpenArgs

    'setup filter
'Debug.Print Me.Parent.lblContext.Caption
'    If Len(Nz(Me.Parent.lblContext.Caption, "")) > 0 Then
'        Dim ary() As String
'
'        If InStr(Me.Parent.lblContext.Caption, "-") Then
'            ary = Split(Me.Parent.lblContext.Caption, "-")
'
'            Me.CommentType = ary(0)
'            Me.CommentType_ID = ary(1)
'
'            Me.Filter = "CommentType=" & Me.CommentType & " AND CommentType_ID=" & Me.CommentType_ID
'        End If
'    End If
        
    Me.Caption = "Comment List"
    lblTitle.Caption = ""
    lblDirections.Caption = "Sort records by clicking the header." _
                            & vbCrLf & "Last modified date color reflects if comment is recent (in the last month)."
    tbxIcon.Value = StringFromCodepoint(uLocked)
    tbxIcon.ForeColor = lngDkGreen
    lblDirections.ForeColor = lngLtBlue
    
    'set hover
    btnDelete.HoverColor = lngGreen
    btnEdit.HoverColor = lngGreen
    btnOpenTable.HoverColor = lngGreen
    
    btnDelete.Caption = StringFromCodepoint(uDelete)
    btnDelete.ForeColor = lngRed

    'enable textbox to ensure scrollbar is available for longer text
    tbxComment.Enabled = True
    
    'cover to avoid data entry
    
    'default
    btnOpenTable.Visible = False
    
    'allow viewing ALL comments via table
    If DEV_MODE Then btnOpenTable.Visible = True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[CommentList form])"
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
            "Error encountered (#" & Err.Number & " - Form_Load[CommentList form])"
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
'   BLC - 2/1/2017 - handles giving focus to new template after TemplateAdd
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

'    If FormIsOpen("TemplateAdd") Then
'
'        'go to the new record (assumes new record is last)
'        'must sort & give form focus first --> sort is a toggle so do version then ID
'        '                                      to give low to high ID
'        Call lblCommentType_Click
'        Call lblHdrID_Click
'        Me.SetFocus
'        DoCmd.GoToRecord acDataForm, Me.Name, acLast
''        DoCmd.GoToRecord acDataForm, Me.Name, acPrevious, 30 << causes endless loop
'
'        'close form
'        DoCmd.Close acForm, "TemplateAdd"
'    Else
'        Debug.Print "not open"
'    End If
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[CommentList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnOpenTable_Click
' Description:  Open table button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 10, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/10/2017 - initial version
' ---------------------------------
Private Sub btnOpenTable_Click()
On Error GoTo Err_Handler
    
    'minimize CommentList
    ToggleForm "CommentList", -1
    
    DoCmd.OpenTable "AppComment", acViewNormal

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnOpenTable_Click[CommentList form])"
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
            "Error encountered (#" & Err.Number & " - btnEdit_Click[CommentList form])"
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
'   BLC - 11/24/2017 - fixed to delete from AppComment vs tsys_Db_Templates
' ---------------------------------
Private Sub btnDelete_Click()
On Error GoTo Err_Handler
    
    Dim result As Integer
    
    'identify the record ID
     result = MsgBox("Delete Record this record: #" & tbxID & " ?" _
                        & vbCrLf & "This action cannot be undone.", vbYesNo, "Delete Record?")

    If result = vbYes Then DeleteRecord "AppComment", tbxID
    
    'clear the deleted record
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[CommentList form])"
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
            "Error encountered (#" & Err.Number & " - lblHdrID_Click[CommentList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblCommentType_Click
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
Private Sub lblCommentType_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblCommentType

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblCommentType_Click[CommentList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblCommentTypeID_Click
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
Private Sub lblCommentTypeID_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblCommentTypeID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblCommentTypeID_Click[CommentList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblComment_Click
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
Private Sub lblComment_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblComment

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblComment_Click[CommentList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLastModifiedDate_Click
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
Private Sub lblLastModifiedDate_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblLastModifiedDate
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLastModifiedDate_Click[CommentList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxComment_MouseMove
' Description:  mouse move (hover) actions
' Assumptions:  -
'               Comment textbox is disabled, so control tips won't display
'               Otherwise this would be tbxComment_MouseMove instead & tbxComment would
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
Private Sub tbxComment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Me.tbxComment.ControlTipText = Nz(FetchAddlData("AppComment", "Comment", Me.tbxID)(0), "")
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxComment_MouseMove[CommentList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Detail_MouseMove
' Description:  mouse move (hover) actions
' Assumptions:  -
'               Comment textbox is disabled, so control tips won't display
'               Otherwise this would be tbxComment_MouseMove instead & tbxComment would
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

    Me.tbxComment.ControlTipText = ""
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_MouseMove[CommentList form])"
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
'   BLC - 10/18/2017 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[CommentList form])"
    End Select
    Resume Exit_Handler
End Sub
