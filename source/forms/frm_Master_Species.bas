Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11580
    DatasheetFontHeight =9
    ItemSuffix =162
    Left =3720
    Top =2340
    Right =11820
    Bottom =8055
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x748529458d1ee340
    End
    RecordSource ="tlu_NCPN_Plants"
    Caption ="Master Species Lookup"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =7200
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4080
                    Top =120
                    Width =3780
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Master Species Lookup"
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =11
                    ListRows =30
                    ListWidth =14400
                    Left =180
                    Top =900
                    Width =2040
                    BoundColumn =1
                    Name ="Select_Code"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1;1440;1152;1872;1728;864;1872;864;1440;864;1872"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2220
                    Top =1440
                    Width =3600
                    TabIndex =1
                    Name ="Master_Species"
                    ControlSource ="Master_Species"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =1440
                            Width =1980
                            Height =240
                            FontWeight =700
                            Name ="Label4"
                            Caption ="Master Species (ITIS)"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7200
                    Top =1440
                    Width =2040
                    TabIndex =2
                    Name ="Master_Family"
                    ControlSource ="Master_Family"

                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =5820
                            Top =1440
                            Width =1320
                            Height =240
                            FontWeight =700
                            Name ="Label6"
                            Caption ="Master Family"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2280
                    Top =1860
                    Width =2940
                    TabIndex =3
                    Name ="Master_Common_Name"
                    ControlSource ="Master_Common_Name"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =1860
                            Width =2040
                            Height =240
                            FontWeight =700
                            Name ="Label8"
                            Caption ="Master Common Name"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =87
                    Left =180
                    Top =660
                    Width =2040
                    Height =240
                    FontWeight =700
                    Name ="Combo_Caption"
                    Caption ="Master PLANTS Code"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7200
                    Top =1860
                    Width =3660
                    Height =600
                    TabIndex =4
                    Name ="Add_Synonyms"
                    ControlSource ="Add_Synonyms"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5280
                            Top =1860
                            Width =1860
                            Height =240
                            FontWeight =700
                            Name ="Label11"
                            Caption ="Additional Synonyms"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1860
                    Top =2460
                    Width =9000
                    Height =420
                    TabIndex =5
                    Name ="Taxonomic_Notes"
                    ControlSource ="Taxonomic_Notes"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =2460
                            Width =1620
                            Height =240
                            FontWeight =700
                            Name ="Label13"
                            Caption ="Taxonomic Notes:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Top =3420
                    Width =480
                    Height =240
                    FontWeight =700
                    Name ="Label14"
                    Caption ="Utah"
                End
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Top =3840
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label16"
                    Caption ="Colorado"
                End
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Top =4260
                    Width =900
                    Height =240
                    FontWeight =700
                    Name ="Label17"
                    Caption ="Wyoming"
                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =5
                    Left =60
                    Top =3000
                    Width =11520
                    Name ="Line18"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1320
                    Top =3420
                    Width =960
                    TabIndex =6
                    Name ="Utah_PLANT_Code"
                    ControlSource ="Utah_PLANT_Code"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =1260
                            Top =3120
                            Width =1080
                            Height =240
                            FontWeight =700
                            Name ="Label20"
                            Caption ="PLANT Code"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1320
                    Top =3840
                    Width =959
                    TabIndex =7
                    Name ="Co_PLANT_Code"
                    ControlSource ="Co_PLANT_Code"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1320
                    Top =4260
                    Width =959
                    TabIndex =8
                    Name ="Wy_PLANT_code"
                    ControlSource ="Wy_PLANT_code"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2460
                    Top =3420
                    Width =3660
                    TabIndex =9
                    Name ="Utah_Species"
                    ControlSource ="Utah_Species"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2460
                            Top =3120
                            Width =720
                            Height =240
                            FontWeight =700
                            Name ="Label25"
                            Caption ="Species"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2460
                    Top =3840
                    Width =3660
                    TabIndex =10
                    Name ="Co_Species"
                    ControlSource ="Co_Species"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2460
                    Top =4260
                    Width =3660
                    TabIndex =11
                    Name ="Wy_Species"
                    ControlSource ="Wy_Species"

                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =6120
                    Top =3840
                    Width =1980
                    Height =240
                    Name ="Label29"
                    Caption ="(Weber && Wittmann 2001)"
                End
                Begin Label
                    OverlapFlags =87
                    Left =6120
                    Top =3420
                    Width =1380
                    Height =240
                    Name ="Label30"
                    Caption ="(Welsh et al 2003)"
                End
                Begin Label
                    OverlapFlags =87
                    Left =6120
                    Top =4260
                    Width =960
                    Height =240
                    Name ="Label31"
                    Caption ="(Dorn 2001)"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8820
                    Top =3420
                    Width =1800
                    TabIndex =12
                    Name ="UT_Family"
                    ControlSource ="UT_Family"
                    StatusBarText ="Utah Family"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8820
                            Top =3120
                            Width =660
                            Height =240
                            FontWeight =700
                            Name ="Label34"
                            Caption ="Family"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8820
                    Top =3840
                    Width =1800
                    TabIndex =13
                    Name ="CO_Family"
                    ControlSource ="CO_Family"
                    StatusBarText ="Colorado Family"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8820
                    Top =4260
                    Width =1800
                    TabIndex =14
                    Name ="WY_Family"
                    ControlSource ="WY_Family"
                    StatusBarText ="Wyoming Family"

                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =600
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =15
                    Name ="Check91"
                    ControlSource ="=IIf([ARCH]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =480
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label92"
                            Caption ="ARCH"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7920
                    Top =60
                    Width =420
                    TabIndex =16
                    Name ="ARCH"
                    ControlSource ="ARCH"
                    StatusBarText ="Park presence descriptor for ARCH"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =8340
                    Top =60
                    Width =420
                    TabIndex =17
                    Name ="BLCA"
                    ControlSource ="BLCA"
                    StatusBarText ="Park presence descriptor for BLCA"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =8760
                    Top =60
                    Width =420
                    TabIndex =18
                    Name ="BRCA"
                    ControlSource ="BRCA"
                    StatusBarText ="Park presence descriptor for BRCA"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =9180
                    Top =60
                    Width =420
                    TabIndex =19
                    Name ="CANY"
                    ControlSource ="CANY"
                    StatusBarText ="Park presence descriptor for CANY"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =9600
                    Top =60
                    Width =420
                    TabIndex =20
                    Name ="CARE"
                    ControlSource ="CARE"
                    StatusBarText ="Park presence descriptor for CARE"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =10020
                    Top =60
                    Width =420
                    TabIndex =21
                    Name ="CEBR"
                    ControlSource ="CEBR"
                    StatusBarText ="Park presence descriptor for CEBR"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =10440
                    Top =60
                    Width =420
                    TabIndex =22
                    Name ="COLM"
                    ControlSource ="COLM"
                    StatusBarText ="Park presence descriptor for COLM"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =10860
                    Top =60
                    Width =420
                    TabIndex =23
                    Name ="CURE"
                    ControlSource ="CURE"
                    StatusBarText ="Park presence descriptor for CURE"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =8340
                    Top =300
                    Width =420
                    TabIndex =24
                    Name ="FOBU"
                    ControlSource ="FOBU"
                    StatusBarText ="Park presence descriptor for FOBU"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =8760
                    Top =300
                    Width =420
                    TabIndex =25
                    Name ="GOSP"
                    ControlSource ="GOSP"
                    StatusBarText ="Park presence descriptor for GOSP"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =9600
                    Top =300
                    Width =420
                    TabIndex =26
                    Name ="NABR"
                    ControlSource ="NABR"
                    StatusBarText ="Park presence descriptor for NABR"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =10020
                    Top =300
                    Width =420
                    TabIndex =27
                    Name ="PISP"
                    ControlSource ="PISP"
                    StatusBarText ="Park presence descriptor for PISP"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =10440
                    Top =300
                    Width =420
                    TabIndex =28
                    Name ="TICA"
                    ControlSource ="TICA"
                    StatusBarText ="Park presence descriptor for TICA"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =10860
                    Top =300
                    Width =420
                    TabIndex =29
                    Name ="ZION"
                    ControlSource ="ZION"
                    StatusBarText ="Park presence descriptor for ZION"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5160
                    Top =4740
                    Width =1500
                    Height =240
                    FontWeight =700
                    Name ="Label111"
                    Caption ="Present in Parks"
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =1260
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =30
                    Name ="Check113"
                    ControlSource ="=IIf([BLCA]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1140
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label114"
                            Caption ="BLCA"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =1920
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =31
                    Name ="Check115"
                    ControlSource ="=IIf([BRCA]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1800
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label116"
                            Caption ="BRCA"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =2580
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =32
                    Name ="Check117"
                    ControlSource ="=IIf([CANY]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2460
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label118"
                            Caption ="CANY"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =3240
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =33
                    Name ="Check119"
                    ControlSource ="=IIf([CARE]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label120"
                            Caption ="CARE"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =7920
                    Top =300
                    Width =420
                    TabIndex =34
                    Name ="DINO(UT)"
                    ControlSource ="DINO(UT)"
                    StatusBarText ="Park presence descriptor for DINO - Utah"
                    EventProcPrefix ="DINO_UT_"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =7920
                    Top =540
                    Width =420
                    TabIndex =35
                    Name ="DINO(CO)"
                    ControlSource ="DINO(CO)"
                    StatusBarText ="Park presence descriptor for DINO - Colorado"
                    EventProcPrefix ="DINO_CO_"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =9180
                    Top =300
                    Width =420
                    TabIndex =36
                    Name ="HOVE(UT)"
                    ControlSource ="HOVE(UT)"
                    StatusBarText ="Park presence descriptor for HOVE - Utah"
                    EventProcPrefix ="HOVE_UT_"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =9180
                    Top =540
                    Width =420
                    TabIndex =37
                    Name ="HOVE(CO)"
                    ControlSource ="HOVE(CO)"
                    StatusBarText ="Park presence descriptor for HOVE - Colorado"
                    EventProcPrefix ="HOVE_CO_"

                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =3900
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =38
                    Name ="Check125"
                    ControlSource ="=IIf([CEBR]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3780
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label126"
                            Caption ="CEBR"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =4560
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =39
                    Name ="Check127"
                    ControlSource ="=IIf([COLM]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4440
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label128"
                            Caption ="COLM"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =5220
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =40
                    Name ="Check129"
                    ControlSource ="=IIf([CURE]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5100
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label130"
                            Caption ="CURE"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =5880
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =41
                    Name ="Check131"
                    ControlSource ="=IIf(([DINO(UT)]=\"Present\") Or ([DINO(CO)]=\"Present\"),-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5760
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label132"
                            Caption ="DINO"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =6540
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =42
                    Name ="Check133"
                    ControlSource ="=IIf([FOBU]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6420
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label134"
                            Caption ="FOBU"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =7200
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =43
                    Name ="Check135"
                    ControlSource ="=IIf([GOSP]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7080
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label136"
                            Caption ="GOSP"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =7860
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =44
                    Name ="Check137"
                    ControlSource ="=IIf(([HOVE(UT)]=\"Present\") Or ([HOVE(CO)]=\"Present\"),-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7740
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label138"
                            Caption ="HOVE"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =8520
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =45
                    Name ="Check139"
                    ControlSource ="=IIf([NABR]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8400
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label140"
                            Caption ="NABR"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =9180
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =46
                    Name ="Check141"
                    ControlSource ="=IIf([PISP]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9060
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label142"
                            Caption ="PISP"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =9840
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =47
                    Name ="Check143"
                    ControlSource ="=IIf([TICA]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9720
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label144"
                            Caption ="TICA"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =10500
                    Top =5340
                    Width =240
                    Height =180
                    TabIndex =48
                    Name ="Check145"
                    ControlSource ="=IIf([ZION]=\"Present\",-1,0)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10380
                            Top =5040
                            Width =495
                            Height =240
                            Name ="Label146"
                            Caption ="ZION"
                        End
                    End
                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =5
                    Left =60
                    Top =4620
                    Width =11460
                    Name ="Line147"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5220
                    Top =6060
                    Width =1140
                    Height =405
                    TabIndex =49
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin OptionGroup
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =2400
                    Top =720
                    Width =5706
                    Height =478
                    TabIndex =50
                    Name ="Sort_Options"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            Left =2520
                            Top =600
                            Width =1530
                            Height =240
                            BackColor =-2147483633
                            Name ="Label150"
                            Caption ="Sort List by MASTER"
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =2580
                            Top =928
                            OptionValue =1
                            Name ="Option152"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =2810
                                    Top =900
                                    Width =960
                                    Height =240
                                    Name ="Label153"
                                    Caption ="PLANT Code"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =3960
                            Top =928
                            OptionValue =2
                            Name ="Option154"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =4190
                                    Top =900
                                    Width =615
                                    Height =240
                                    Name ="Label155"
                                    Caption ="Species"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =4980
                            Top =928
                            OptionValue =3
                            Name ="Option156"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =5210
                                    Top =900
                                    Width =1125
                                    Height =240
                                    Name ="Label157"
                                    Caption ="Family-Species"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =6480
                            Top =928
                            OptionValue =4
                            Name ="Option158"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6710
                                    Top =900
                                    Width =1140
                                    Height =240
                                    Name ="Label159"
                                    Caption ="Common Name"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2040
                    Top =240
                    Width =959
                    TabIndex =51
                    Name ="Master_PLANT_Code"
                    ControlSource ="Master_PLANT_Code"
                    StatusBarText ="Master Species PLANTS Code"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =240
                            Width =1800
                            Height =240
                            FontWeight =700
                            Name ="Label160"
                            Caption ="Master PLANTS Code"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7740
                    Top =6360
                    Width =3510
                    Height =405
                    TabIndex =52
                    ForeColor =3767809
                    Name ="ButtonSave"
                    Caption ="Save this plant in detail record and close form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  Dim strSQL As String
  Dim strSortBy As String
  Dim strCol1 As String

      strSortBy = "Master_Plant_Code"
      strCol1 = "Master_Plant_Code"
      Me!Master_PLANT_Code.visible = False
  
  strSQL = "SELECT " & strCol1 & ", Master_PLANT_Code AS [Master Plant Code], Master_Family AS [Master Family], Master_Species AS [Master Species], " & _
  "Master_Common_Name AS [Master Common Name], Utah_PLANT_Code AS [Utah Code], Utah_Species AS [Utah Species], " & _
  "Co_PLANT_Code AS [CO Code], Co_Species AS [CO Species], Wy_PLANT_code AS [WY Code], Wy_Species AS [WY Species] " & _
  "FROM tlu_NCPN_Plants ORDER BY "
  Me!Select_Code.RowSource = strSQL & strSortBy
End Sub

Private Sub Form_Open(Cancel As Integer)
  Me![Select_Code] = Me![Master_PLANT_Code]
End Sub

Private Sub Select_Code_AfterUpdate()
    Me.RecordsetClone.FindFirst "[Master_PLANT_Code] = '" & Me![Select_Code].Column(1) & "'"
    Me.Bookmark = Me.RecordsetClone.Bookmark
End Sub
Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click

    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

Private Sub Sort_Options_AfterUpdate()
  Dim strSQL As String
  Dim strSortBy As String
  Dim strCol1 As String
  
  Select Case Sort_Options
    Case 1
      strSortBy = "Master_Plant_Code"
      strCol1 = "Master_Plant_Code"
      Me!Combo_Caption.Caption = "Master PLANTS Code"
      Me!Master_PLANT_Code.visible = False
    Case 2
      strSortBy = "Master_Species"
      strCol1 = "Master_Species"
      Me!Combo_Caption.Caption = "Master Species"
      Me!Master_PLANT_Code.visible = True
    Case 3
      strSortBy = "Master_Family, Master_Species"
      strCol1 = "Master_Family"
      Me!Combo_Caption.Caption = "Master Family"
      Me!Master_PLANT_Code.visible = True
    Case Else
      strSortBy = "Master_Common_Name"
      strCol1 = "Master_Common_Name"
      Me!Combo_Caption.Caption = "Master Common Name"
      Me!Master_PLANT_Code.visible = True
  End Select
  
  strSQL = "SELECT " & strCol1 & ", Master_PLANT_Code AS [Master Plant Code], Master_Family AS [Master Family], Master_Species AS [Master Species], " & _
  "Master_Common_Name AS [Master Common Name], Utah_PLANT_Code AS [Utah Code], Utah_Species AS [Utah Species], " & _
  "Co_PLANT_Code AS [CO Code], Co_Species AS [CO Species], Wy_PLANT_code AS [WY Code], Wy_Species AS [WY Species] " & _
  "FROM tlu_NCPN_Plants ORDER BY "
  Me!Select_Code.RowSource = strSQL & strSortBy
  Me!Select_Code.Requery
End Sub
Private Sub ButtonSave_Click()
On Error GoTo Err_ButtonSave_Click

  If Not IsNull(Me.OpenArgs) Then
    If Me.OpenArgs = "fsub_Quadrat_Shrubs" Then  ' Find the right calling form
      Forms!frm_Data_Entry!frm_Quadrat_Transect.Form!fsub_Quadrat.Form!fsub_Quadrat_Shrubs.Form!Master_Code = Me![Master_PLANT_Code]
      Forms!frm_Data_Entry!frm_Quadrat_Transect.Form!fsub_Quadrat.Form!fsub_Quadrat_Shrubs.Form!Master_Code.Requery
    ElseIf Me.OpenArgs = "fsub_Species" Then
      Forms!frm_Data_Entry!frm_Quadrat_Transect.Form!fsub_Quadrat.Form!fsub_Species.Form!Master_Code = Me![Master_PLANT_Code]
      Forms!frm_Data_Entry!frm_Quadrat_Transect.Form!fsub_Quadrat.Form!fsub_Species.Form!Master_Code.Requery
    ElseIf Me.OpenArgs = "fsub_LP_Belt_Shrub" Then
      If Not IsNull(DLookup("[Shrub_ID]", "tbl_LP_Shrub", "[Transect_ID] = '" & Forms!frm_Data_Entry!frm_LP_Belt_Transect.Form!fsub_LP_Belt_Shrub.Form!Transect_ID & "' AND [Species] = '" & Me![Master_PLANT_Code] & "'")) Then
        MsgBox "This species is already recorded for this transect."
      Else
        Forms!frm_Data_Entry!frm_LP_Belt_Transect.Form!fsub_LP_Belt_Shrub.Form!Species = Me![Master_PLANT_Code]
        Forms!frm_Data_Entry!frm_LP_Belt_Transect.Form!fsub_LP_Belt_Shrub.Form!Species.Requery
      End If
    ElseIf Me.OpenArgs = "fsub_LP_Exotic" Then
      If Not IsNull(DLookup("[Exotic_ID]", "tbl_LP_Exotic", "[Transect_ID] = '" & Forms!frm_Data_Entry!frm_LP_Belt_Transect.Form!fsub_LP_Exotic.Form!Transect_ID & "' AND [Species] = '" & Me![Master_PLANT_Code] & "'")) Then
        MsgBox "This species is already recorded for this transect."
      Else
        Forms!frm_Data_Entry!frm_LP_Belt_Transect.Form!fsub_LP_Exotic.Form!Species = Me![Master_PLANT_Code]
        Forms!frm_Data_Entry!frm_LP_Belt_Transect.Form!fsub_LP_Exotic.Form!Species.Requery
      End If
    ElseIf Me.OpenArgs = "fsub_LP_Intercept" Then
      If Not IsNull(DLookup("[LC_ID]", "tbl_LP_Lower_Canopy", "[Intercept_ID] = '" & Forms!frm_Data_Entry!frm_LP_Transect.Form!fsub_LP_Intercept.Form!fsub_LP_Lower_Canopy.Form!Intercept_ID & "' AND [Species] = '" & Me![Master_PLANT_Code] & "'")) Then
        MsgBox "This species is already recorded for this point."
      Else
        Forms!frm_Data_Entry!frm_LP_Transect.Form!fsub_LP_Intercept.Form!Top = Me![Master_PLANT_Code]
        Forms!frm_Data_Entry!frm_LP_Transect.Form!fsub_LP_Intercept.Form!Top.Requery
        Forms!frm_Data_Entry!frm_LP_Transect.Form!fsub_LP_Intercept.Form!Alive.Enabled = True
      End If
    Else
      If Not IsNull(DLookup("[LC_ID]", "tbl_LP_Lower_Canopy", "[Intercept_ID] = '" & Forms!frm_Data_Entry!frm_LP_Transect.Form!fsub_LP_Intercept.Form!fsub_LP_Lower_Canopy.Form!Intercept_ID & "' AND [Species] = '" & Me![Master_PLANT_Code] & "'")) Or Not IsNull(DLookup("[Intercept_ID]", "tbl_LP_Intercept", "[Intercept_ID] = '" & Forms!frm_Data_Entry!frm_LP_Transect.Form!fsub_LP_Intercept.Form!fsub_LP_Lower_Canopy.Form!Intercept_ID & "' AND [Top] = '" & Me![Master_PLANT_Code] & "'")) Then
        MsgBox "This species is already recorded for this point."
      Else
        Forms!frm_Data_Entry!frm_LP_Transect.Form!fsub_LP_Intercept.Form!fsub_LP_Lower_Canopy.Form!Species = Me![Master_PLANT_Code]
        Forms!frm_Data_Entry!frm_LP_Transect.Form!fsub_LP_Intercept.Form!fsub_LP_Lower_Canopy.Form!Species.Requery
      End If
    End If ' End if for form name tests
  End If   ' End if for null OpenArgs test
  DoCmd.Close acForm, "frm_Master_Species"

Exit_ButtonSave_Click:
    Exit Sub

Err_ButtonSave_Click:
    MsgBox Err.Description
    Resume Exit_ButtonSave_Click
    
End Sub
