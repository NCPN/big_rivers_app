Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6912
    DatasheetFontHeight =11
    ItemSuffix =55
    Left =3144
    Top =864
    Right =11832
    Bottom =8592
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xc1f3db6ed487e440
    End
    RecordSource ="tbl_Target_Areas"
    Caption ="Photo Details"
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
    SplitFormSplitterBar =0
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
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
        Begin FormHeader
            Height =6780
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =2424
                    Width =6912
                    Height =360
                    BackColor =15858167
                    BorderColor =15858167
                    Name ="rctSubjectHdr"
                    GridlineColor =10921638
                    LayoutCachedTop =2424
                    LayoutCachedWidth =6912
                    LayoutCachedHeight =2784
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Width =6912
                    Height =360
                    BackColor =15266810
                    BorderColor =10921638
                    Name ="rctPhotogHdr"
                    GridlineColor =10921638
                    LayoutCachedWidth =6912
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5040
                    Top =6300
                    Width =1620
                    Height =372
                    ForeColor =16711680
                    Name ="btnNext"
                    Caption ="Save && Next >>"
                    StatusBarText ="Next to add photo and move to next one"
                    GridlineColor =10921638

                    LayoutCachedLeft =5040
                    LayoutCachedTop =6300
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =6672
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =6750156
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =52377
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =3407769
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =52224
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =2375487
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =6750156
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =1260
                    Top =480
                    Width =1680
                    Height =418
                    ColumnOrder =0
                    TabIndex =1
                    BorderColor =10921638
                    Name ="optgUSDS"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="Upstream or Downstream - double click to clear options"
                    GridlineColor =10921638

                    LayoutCachedLeft =1260
                    LayoutCachedTop =480
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =898
                    Begin
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1380
                            Top =568
                            Width =536
                            Height =210
                            OptionValue =1
                            BorderColor =10921638
                            Name ="optUS"
                            ControlTipText ="Upstream"
                            GridlineColor =10921638

                            LayoutCachedLeft =1380
                            LayoutCachedTop =568
                            LayoutCachedWidth =1916
                            LayoutCachedHeight =778
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1610
                                    Top =540
                                    Width =330
                                    Height =285
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="lblUS"
                                    Caption ="US"
                                    ControlTipText ="Upstream"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1610
                                    LayoutCachedTop =540
                                    LayoutCachedWidth =1940
                                    LayoutCachedHeight =825
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =2100
                            Top =568
                            Width =536
                            Height =216
                            TabIndex =1
                            OptionValue =2
                            BorderColor =10921638
                            Name ="optDS"
                            ControlTipText ="Downstream"
                            GridlineColor =10921638

                            LayoutCachedLeft =2100
                            LayoutCachedTop =568
                            LayoutCachedWidth =2636
                            LayoutCachedHeight =784
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =2330
                                    Top =540
                                    Width =330
                                    Height =300
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="lblDS"
                                    Caption ="DS"
                                    ControlTipText ="Downstream"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =2330
                                    LayoutCachedTop =540
                                    LayoutCachedWidth =2660
                                    LayoutCachedHeight =840
                                End
                            End
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    Left =1260
                    Top =900
                    Width =1680
                    Height =418
                    ColumnOrder =1
                    TabIndex =2
                    BorderColor =10921638
                    Name ="optgRRRL"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="River Right or River Left - double click to clear options"
                    GridlineColor =10921638

                    LayoutCachedLeft =1260
                    LayoutCachedTop =900
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1318
                    Begin
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1380
                            Top =988
                            Width =536
                            Height =210
                            OptionValue =1
                            BorderColor =10921638
                            Name ="optRR"
                            ControlTipText ="River Right"
                            GridlineColor =10921638

                            LayoutCachedLeft =1380
                            LayoutCachedTop =988
                            LayoutCachedWidth =1916
                            LayoutCachedHeight =1198
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1610
                                    Top =960
                                    Width =330
                                    Height =285
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="lblRR"
                                    Caption ="RR"
                                    ControlTipText ="River Right"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1610
                                    LayoutCachedTop =960
                                    LayoutCachedWidth =1940
                                    LayoutCachedHeight =1245
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =2100
                            Top =988
                            Width =596
                            TabIndex =1
                            OptionValue =2
                            BorderColor =10921638
                            Name ="optRL"
                            ControlTipText ="River Left"
                            GridlineColor =10921638

                            LayoutCachedLeft =2100
                            LayoutCachedTop =988
                            LayoutCachedWidth =2696
                            LayoutCachedHeight =1228
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =2330
                                    Top =960
                                    Width =330
                                    Height =300
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="lblRL"
                                    Caption ="RL"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =2330
                                    LayoutCachedTop =960
                                    LayoutCachedWidth =2660
                                    LayoutCachedHeight =1260
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =540
                    Top =480
                    Width =660
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFacing"
                    Caption ="Facing"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="Direction Facing - double click to clear options"
                    GridlineColor =10921638
                    LayoutCachedLeft =540
                    LayoutCachedTop =480
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =795
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =1440
                    Width =2760
                    Height =315
                    ColumnOrder =2
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPhotog"
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedTop =1440
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1755
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =540
                            Top =1440
                            Width =690
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPhotog"
                            Caption ="Name"
                            GridlineColor =10921638
                            LayoutCachedLeft =540
                            LayoutCachedTop =1440
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =1755
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =240
                    Top =19
                    Width =1350
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPhotogHdr"
                    Caption ="Photographer"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =19
                    LayoutCachedWidth =1590
                    LayoutCachedHeight =334
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =1872
                    Width =2760
                    Height =315
                    ColumnOrder =3
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPhotogLoc"
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedTop =1872
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =2187
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =540
                            Top =1872
                            Width =855
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPhotogLoc"
                            Caption ="Location"
                            GridlineColor =10921638
                            LayoutCachedLeft =540
                            LayoutCachedTop =1872
                            LayoutCachedWidth =1395
                            LayoutCachedHeight =2187
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =2940
                    Width =2760
                    Height =315
                    ColumnOrder =4
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSubjectLoc"
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedTop =2940
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =3255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =540
                            Top =2940
                            Width =855
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSubjectLoc"
                            Caption ="Location"
                            GridlineColor =10921638
                            LayoutCachedLeft =540
                            LayoutCachedTop =2940
                            LayoutCachedWidth =1395
                            LayoutCachedHeight =3255
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =240
                    Top =2460
                    Width =1350
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSubjectHdr"
                    Caption ="Subject"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =2460
                    LayoutCachedWidth =1590
                    LayoutCachedHeight =2775
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3600
                    Top =540
                    Width =1200
                    Height =300
                    ColumnOrder =5
                    TabIndex =6
                    BorderColor =10921638
                    Name ="cbxCloseup"
                    ControlTipText ="Is photo a Closeup?"
                    GridlineColor =10921638

                    LayoutCachedLeft =3600
                    LayoutCachedTop =540
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3830
                            Top =510
                            Width =930
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblCloseup"
                            Caption ="Closeup?"
                            GridlineColor =10921638
                            LayoutCachedLeft =3830
                            LayoutCachedTop =510
                            LayoutCachedWidth =4760
                            LayoutCachedHeight =825
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3600
                    Top =930
                    Width =1620
                    Height =300
                    ColumnOrder =6
                    TabIndex =7
                    BorderColor =10921638
                    Name ="cbxReplacement"
                    ControlTipText ="Is photo a Replacement?"
                    GridlineColor =10921638

                    LayoutCachedLeft =3600
                    LayoutCachedTop =930
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3830
                            Top =900
                            Width =1410
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblReplacement"
                            Caption ="Replacement?"
                            GridlineColor =10921638
                            LayoutCachedLeft =3830
                            LayoutCachedTop =900
                            LayoutCachedWidth =5240
                            LayoutCachedHeight =1215
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =3960
                    Width =2760
                    Height =315
                    ColumnOrder =7
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPhotoNum"
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedTop =3960
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =4275
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =540
                            Top =3960
                            Width =855
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPhotoNum"
                            Caption ="Photo #"
                            GridlineColor =10921638
                            LayoutCachedLeft =540
                            LayoutCachedTop =3960
                            LayoutCachedWidth =1395
                            LayoutCachedHeight =4275
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =4800
                    Width =4080
                    Height =1560
                    ColumnOrder =8
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxComments"
                    GridlineColor =10921638

                    LayoutCachedLeft =780
                    LayoutCachedTop =4800
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =6360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =540
                            Top =4428
                            Width =1080
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label43"
                            Caption ="Comments"
                            GridlineColor =10921638
                            LayoutCachedLeft =540
                            LayoutCachedTop =4428
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =4743
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4500
                    Top =1872
                    Width =2280
                    Height =300
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblPhotogLocHint"
                    Caption ="Photog loc hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =4500
                    LayoutCachedTop =1872
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =2172
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4500
                    Top =2940
                    Width =2280
                    Height =300
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblSubjectLocHint"
                    Caption ="Subject loc hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =4500
                    LayoutCachedTop =2940
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =3240
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4500
                    Top =3840
                    Width =2280
                    Height =840
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblPhotoNumHint"
                    Caption ="Photo # hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =4500
                    LayoutCachedTop =3840
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =4680
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4980
                    Top =4800
                    Width =1800
                    Height =660
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblCommentHint"
                    Caption ="Comment hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =4980
                    LayoutCachedTop =4800
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =5460
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =5340
                    Top =420
                    Width =1440
                    Height =420
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblCloseupHint"
                    Caption ="Closeup hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =5340
                    LayoutCachedTop =420
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =840
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =5340
                    Top =900
                    Width =1440
                    Height =480
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblReplacementHint"
                    Caption ="Replacement hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =5340
                    LayoutCachedTop =900
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =1380
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =3420
                    Width =6912
                    Height =360
                    BackColor =16381933
                    BorderColor =16381933
                    Name ="rctPhotoDetailHdr"
                    GridlineColor =10921638
                    LayoutCachedTop =3420
                    LayoutCachedWidth =6912
                    LayoutCachedHeight =3780
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =240
                    Top =3456
                    Width =1350
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPhotoDetails"
                    Caption ="Photo Details"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =3456
                    LayoutCachedWidth =1590
                    LayoutCachedHeight =3771
                End
            End
        End
        Begin Section
            Height =0
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' FORM:         Form_fsub_Photo_FTOR_Details
' Description:  Photo detail functions & procedures for feature, transect, overview & reference photos
'
' Source/date:  Bonnie Campbell, 7/13/2015
' Revisions:    BLC - 7/13/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  Actions for form loading
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, July 13, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/13/2015 - initial version
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler


    'disable Next to start
    btnNext.Enabled = False

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[form_fsub_Photo_FTOR_Details])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxPark_Change
' Description:  Actions to take when a park is selected
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
'   BLC - 6/12/2015 - changed the check from 0 to 3 (Park_Code = 4 characters) and
'                     added enabling continue button, changed TempVars.item("... to TempVars("...
' ---------------------------------
Private Sub cbxPark_Change()
On Error GoTo Err_Handler
    
    'set park & enable continue when a 4-letter park code is selected
    If Len(cbxPark.Value) > 3 Then
        'set park
        TempVars("park") = Trim(cbxPark.Value)
        
        'enable the continue button
        If Len(cbxPark) > 3 And TempVars("TgtYear") > 0 Then
            btnContinue.Enabled = True
        End If
    End If
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPark_Change[form_fsub_Select_Park_Year])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxYear_Change
' Description:  Actions to take when a task action is selected
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015 - initial version
'   BLC - 6/12/2015 - added enabling continue button, changed TempVars.item("... to TempVars("...
' ---------------------------------
Private Sub cbxYear_Change()
On Error GoTo Err_Handler

    If Len(Trim(cbxYear)) > 0 Then
        'set year
        TempVars("TgtYear") = cbxYear.Value
        
        'enable the continue button
        If Len(cbxPark) > 3 And TempVars("TgtYear") > 0 Then
            btnContinue.Enabled = True
        End If
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxYear_Change[form_fsub_Select_Park_Year])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnNext_Click
' Description:  Save photo info & go to next actions
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, July 13, 2015 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 7/13/2015 - initial version
' ---------------------------------
Private Sub btnNext_Click()
On Error GoTo Err_Handler
       
    'clear year & park (prevents NULL errors & click continue if values aren't set)
'    cbxYear.Value = "SEL"
'    cbxPark.Value = ""
       
    'open target species list
'    DoCmd.OpenForm "frm_Tgt_Species", acNormal, , , , , TempVars("TgtYear")
    
    ' save & move to next photo in tree
    
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNext_Click[form_fsub_Photo_FTOR_Details])"
    End Select
    Resume Exit_Sub
End Sub

' http://www.geeksengine.com/article/unselect_access_radio_buttons.html
Private Sub lblFacing_DblClick(Cancel As Integer)
    'clear options
    optgRRRL.Value = ""
    optgUSDS.Value = ""
End Sub

' http://www.geeksengine.com/article/unselect_access_radio_buttons.html
Private Sub optgRRRL_DblClick(Cancel As Integer)
    'clear options
    optgRRRL.Value = ""
End Sub

' http://www.geeksengine.com/article/unselect_access_radio_buttons.html
Private Sub optgUSDS_DblClick(Cancel As Integer)
    'clear options
    optgUSDS.Value = ""
End Sub
