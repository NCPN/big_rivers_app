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
    Width =6780
    DatasheetFontHeight =11
    ItemSuffix =56
    Right =12855
    Bottom =11790
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xcb7a7f08cdc4e440
    End
    Caption ="Photo Details"
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
            Height =6660
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
                    Width =6780
                    Height =360
                    BackColor =15858167
                    BorderColor =15858167
                    Name ="rctSubjectHdr"
                    GridlineColor =10921638
                    LayoutCachedTop =2424
                    LayoutCachedWidth =6780
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
                    Width =6780
                    Height =360
                    BackColor =15266810
                    BorderColor =10921638
                    Name ="rctPhotogHdr"
                    GridlineColor =10921638
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1260
                    Top =6240
                    Width =1620
                    Height =372
                    ForeColor =16711680
                    Name ="btnNext"
                    Caption ="Save && Next >>"
                    StatusBarText ="Next to add photo and move to next one"
                    GridlineColor =10921638

                    LayoutCachedLeft =1260
                    LayoutCachedTop =6240
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =6612
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
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
                    ColumnOrder =1
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
                    ColumnOrder =2
                    TabIndex =3
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
                    AfterUpdate ="[Event Procedure]"
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
                    AfterUpdate ="[Event Procedure]"
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
                    Name ="chkCloseup"
                    OnClick ="[Event Procedure]"
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
                    Name ="chkReplacement"
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
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPhotoNum"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedTop =3960
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =4275
                    BackThemeColorIndex =-1
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
                    Height =1380
                    ColumnOrder =8
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxComments"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =780
                    LayoutCachedTop =4800
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =6180
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
                    Width =1860
                    Height =405
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblPhotogLocHint"
                    Caption ="Photog loc hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =4500
                    LayoutCachedTop =1872
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =2277
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4500
                    Top =2940
                    Width =1860
                    Height =360
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblSubjectLocHint"
                    Caption ="Subject loc hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =4500
                    LayoutCachedTop =2940
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =3300
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4500
                    Top =3960
                    Width =2220
                    Height =720
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblPhotoNumHint"
                    Caption ="Photo # hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =4500
                    LayoutCachedTop =3960
                    LayoutCachedWidth =6720
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
                    Top =480
                    Width =1440
                    Height =360
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblCloseupHint"
                    Caption ="Closeup hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =5340
                    LayoutCachedTop =480
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
                    Height =360
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
                    LayoutCachedHeight =1260
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =3420
                    Width =6780
                    Height =360
                    BackColor =16381933
                    BorderColor =16381933
                    Name ="rctPhotoDetailHdr"
                    GridlineColor =10921638
                    LayoutCachedTop =3420
                    LayoutCachedWidth =6780
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
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =4680
                    Top =6240
                    Width =1800
                    TabIndex =10
                    ForeColor =4210752
                    Name ="btnSave"
                    Caption ="Save && Next >>"
                    StatusBarText ="Save photo details & move to next photo"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save photo details & move to next photo"
                    GridlineColor =10921638

                    LayoutCachedLeft =4680
                    LayoutCachedTop =6240
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =6600
                    PictureCaptionArrangement =5
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
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1620
                    Top =1440
                    Width =2760
                    Height =315
                    ColumnOrder =0
                    TabIndex =2
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x01000000a2000000020000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff00010000000000000004000000200000000100000000000000 ,
                        0xfff2000000000000000000000000000000000000000000000000000000000000 ,
                        0x220022000000000049004900660028004c0065006e0028005b00630062007800 ,
                        0x500068006f0074006f0067005d0029003d0030002c0031002c00300029000000 ,
                        0x0000
                    End
                    Name ="cbxPhotog"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1440"
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedTop =1440
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1755
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000fff200001b00000049004900660028004c0065006e0028005b ,
                        0x00630062007800500068006f0074006f0067005d0029003d0030002c0031002c ,
                        0x0030002900000000000000000000000000000000000000000000
                    End
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =4440
                    Top =1440
                    Width =1320
                    TabIndex =11
                    ForeColor =4210752
                    Name ="btnContacts"
                    Caption ="Add Contact"
                    StatusBarText ="Add new contact"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add new contact"
                    GridlineColor =10921638

                    LayoutCachedLeft =4440
                    LayoutCachedTop =1440
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =1800
                    PictureCaptionArrangement =5
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
' FORM:         PhotoFTORDetailsDetails Form
' Description:  Photo detail functions & procedures for feature, transect, overview & reference photos
'
' Source/date:  Bonnie Campbell, 7/13/2015
' Revisions:    BLC - 7/13/2015 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------

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
    
    'set hover
    btnSave.HoverColor = lngGreen
      
    'defaults
    btnSave.Enabled = False
    cbxPhotog.BackColor = lngYellow
    tbxPhotoNum.BackColor = lngYellow
  
    'initialize values
    Set Me.cbxPhotog.Recordset = GetRecords("s_contact_list")
    
'    ClearForm Me
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[PhotoFTORDetails form])"
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
            "Error encountered (#" & Err.Number & " - Form_Load[PhotoFTORDetails form])"
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
              
      'If tbxID > 0 Then btnComment.Enabled = True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[PhotoFTORDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxComments_KeyPress
' Description:  Textbox keypress actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/14/2016 - initial version
' ---------------------------------
Private Sub tbxComments_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler

    LimitKeyPress Me.tbxComments, 255, KeyAscii
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxComments_KeyPress[PhotoFTORDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxContact_Change
' Description:  Combobox change actions
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, October 14, 2016 - for NCPN tools
' Revisions:
'   BLC - 10/14/2016 - initial version
' ---------------------------------
Private Sub cbxContact_Change()
On Error GoTo Err_Handler
    
'    'set park & enable continue when a 4-letter park code is selected
'    If Len(cbxPark.value) > 3 Then
'        'set park
'        TempVars("park") = Trim(cbxPark.value)
'
'        'enable the continue button
'        If Len(cbxPark) > 3 And TempVars("TgtYear") > 0 Then
'            btnContinue.Enabled = True
'        End If
'    End If
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxContact_Change[PhotoFTORDetailsDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxComments_Change
' Description:  Textbox change actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/14/2016 - initial version
' ---------------------------------
Private Sub tbxComments_Change()
On Error GoTo Err_Handler

    LimitChange Me.tbxComments, 255
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxComments_Change[PhotoFTORDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPhotog_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/14/2016 - initial version
' ---------------------------------
Private Sub tbxPhotog_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPhotog_AfterUpdate[PhotoFTORDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPhotogLoc_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/14/2016 - initial version
' ---------------------------------
Private Sub tbxPhotogLoc_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPhotogLoc_AfterUpdate[PhotoFTORDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxSujbectLoc_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/14/2016 - initial version
' ---------------------------------
Private Sub tbxSubjectLoc_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxSubjectLoc_AfterUpdate[PhotoFTORDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPhotoNum_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/14/2016 - initial version
' ---------------------------------
Private Sub tbxPhotoNum_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPhotoNum_AfterUpdate[PhotoFTORDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxComments_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/14/2016 - initial version
' ---------------------------------
Private Sub tbxComments_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxComments_AfterUpdate[PhotoFTORDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnContacts_Click
' Description:  Add contact button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/14/2016 - initial version
' ---------------------------------
Private Sub btnContacts_Click()
On Error GoTo Err_Handler
    
    DoCmd.OpenForm "Contact", acNormal, , , , , "Tree"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnContacts_Click[PhotoFTORDetails form])"
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
            "Error encountered (#" & Err.Number & " - btnUndo_Click[PhotoFTORDetails form])"
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
    
'    Dim s As New Site
'
'    With s
'        'values passed into form
'        .Park = TempVars("ParkCode")
'        .River = TempVars("River")
'
'        'form values
'        .Code = tbxSiteCode.Value
'        .Name = tbxSiteName.Value
'        .Directions = Nz(tbxSiteDirections.Value, "")
'        .Description = Nz(tbxDescription.Value, "")
'
'        'assumed
'        .IsActiveForProtocol = 1 'all sites assumed active when added
'
'        .ID = tbxID.Value '0 if new, edit if > 0
'        .SaveToDb
'
'        'set the tbxID.value
'        'tbxID = .ID #can't assign value to object
'
'    End With
'
'    'clear values & refresh display
'
'    ReadyForSave
'
'    PopulateForm Me, tbxID.Value
'
'    'refresh list
'    Me.list.Requery
'
'    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[PhotoFTORDetails form])"
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
' ---------------------------------
Private Sub btnComment_Click()
On Error GoTo Err_Handler
    
    'open comment form
'    DoCmd.OpenForm "Comment", acNormal, , , , , "event|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[PhotoFTORDetails form])"
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

'    'restore Main
'    ToggleForm "Main", 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[PhotoFTORDetails form])"
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
    'requires: direction facing, photog (comments optional)
    If Len(Nz(cbxPhotog.Value, "")) > 0 _
        And Len(Nz(tbxPhotoNum.Value, "")) > 0 Then
        isOK = True
    End If
    
'    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    btnSave.Enabled = isOK
    
    'refresh form
    Me.Requery
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[PhotoFTORDetails form])"
    End Select
    Resume Exit_Handler
End Sub
'
'
'' ---------------------------------
'' SUB:          Form_Load
'' Description:  Actions for form loading
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   none
'' Source/date:
'' Adapted:      Bonnie Campbell, July 13, 2015 - for NCPN tools
'' Revisions:
''   BLC - 7/13/2015 - initial version
'' ---------------------------------
'Private Sub Form_Load()
'On Error GoTo Err_Handler
'
'    'disable Next to start
'    btnNext.Enabled = False
'
'Exit_Handler:
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - Form_Load[PhotoFTORDetailsDetails form])"
'    End Select
'    Resume Exit_Handler
'End Sub
'
'' ---------------------------------
'' SUB:          cbxYear_Change
'' Description:  Actions to take when a task action is selected
'' Assumptions:  -
'' Parameters:   N/A
'' Returns:      N/A
'' Throws:       none
'' References:   none
'' Source/date:
'' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
'' Revisions:
''   BLC - 2/23/2015 - initial version
''   BLC - 6/12/2015 - added enabling continue button, changed TempVars.item("... to TempVars("...
'' ---------------------------------
'Private Sub cbxYear_Change()
'On Error GoTo Err_Handler
'
''    If Len(Trim(cbxYear)) > 0 Then
''        'set year
''        TempVars("TgtYear") = cbxYear.value
''
''        'enable the continue button
''        If Len(cbxPark) > 3 And TempVars("TgtYear") > 0 Then
''            btnContinue.Enabled = True
''        End If
''    End If
'
'Exit_Handler:
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - cbxYear_Change[PhotoFTORDetailsDetails form])"
'    End Select
'    Resume Exit_Handler
'End Sub
'
'' ---------------------------------
'' SUB:          btnNext_Click
'' Description:  Save photo info & go to next actions
'' Assumptions:  -
'' Parameters:   N/A
'' Returns:      N/A
'' Throws:       none
'' References:   none
'' Source/date:  Bonnie Campbell, July 13, 2015 - for NCPN tools
'' Adapted:
'' Revisions:
''   BLC, 7/13/2015 - initial version
'' ---------------------------------
'Private Sub btnNext_Click()
'On Error GoTo Err_Handler
'
'    'clear year & park (prevents NULL errors & click continue if values aren't set)
''    cbxYear.Value = "SEL"
''    cbxPark.Value = ""
'
'    'open target species list
''    DoCmd.OpenForm "frm_Tgt_Species", acNormal, , , , , TempVars("TgtYear")
'
'    ' save & move to next photo in tree
'
'
'Exit_Handler:
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - btnNext_Click[PhotoFTORDetailsDetails form])"
'    End Select
'    Resume Exit_Handler
'End Sub

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
