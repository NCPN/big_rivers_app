Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =16342
    DatasheetFontHeight =11
    ItemSuffix =24
    Right =10272
    Bottom =8556
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x06c7dc6ed487e440
    End
    RecordSource ="tbl_Target_Species"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
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
        Begin FormHeader
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =1668
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTgtAreaHdr"
                    Caption ="Target Species"
                    GridlineColor =10921638
                    LayoutCachedWidth =1668
                    LayoutCachedHeight =372
                End
                Begin Label
                    OverlapFlags =85
                    Top =660
                    Width =7188
                    Height =576
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblChooseSpeciesType"
                    Caption ="Set the priority and target area information for the target species chosen.     "
                        "  Click Save List to save the target list with the priority and targeting inform"
                        "ation."
                    GridlineColor =10921638
                    LayoutCachedTop =660
                    LayoutCachedWidth =7188
                    LayoutCachedHeight =1236
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1140
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =360
                    Width =420
                    Height =300
                    ForeColor =4210752
                    Name ="cmdDeleteTgtArea_x"
                    Caption ="Delete Target Area"
                    ControlTipText ="Delete Record"
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Action ="GoToControl"
                            Argument ="=[Screen].[PreviousControl].[Name]"
                        End
                        Begin
                            Action ="ClearMacroError"
                        End
                        Begin
                            Condition ="Not [Form].[NewRecord]"
                            Action ="RunCommand"
                            Argument ="223"
                        End
                        Begin
                            Condition ="[Form].[NewRecord] And Not [Form].[Dirty]"
                            Action ="Beep"
                        End
                        Begin
                            Condition ="[Form].[NewRecord] And [Form].[Dirty]"
                            Action ="RunCommand"
                            Argument ="292"
                        End
                        Begin
                            Condition ="[MacroError]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"cmdDeleteTgtArea_x\" xmlns=\"http://schemas.microsoft.com/of"
                                "fice/accessservices/2009/11/application\" xmlns:a=\"http://schemas.microsoft.com"
                                "/office/accessservices/2009/11"
                        End
                        Begin
                            Comment ="_AXL:/forms\"><Statements><Action Name=\"OnError\"/><Action Name=\"GoToControl\""
                                "><Argument Name=\"ControlName\">=[Screen].[PreviousControl].[Name]</Argument></A"
                                "ction><Action Name=\"ClearMacroError\"/><ConditionalBlock><If><Condition>Not [Fo"
                                "rm].[NewRecord]</Condi"
                        End
                        Begin
                            Comment ="_AXL:tion><Statements><Action Name=\"DeleteRecord\"/></Statements></If></Conditi"
                                "onalBlock><ConditionalBlock><If><Condition>[Form].[NewRecord] And Not [Form].[Di"
                                "rty]</Condition><Statements><Action Name=\"Beep\"/></Statements></If></Condition"
                                "alBlock><Conditio"
                        End
                        Begin
                            Comment ="_AXL:nalBlock><If><Condition>[Form].[NewRecord] And [Form].[Dirty]</Condition><S"
                                "tatements><Action Name=\"UndoRecord\"/></Statements></If></ConditionalBlock><Con"
                                "ditionalBlock><If><Condition>[MacroError]&lt;&gt;0</Condition><Statements><Actio"
                                "n Name=\"Message"
                        End
                        Begin
                            Comment ="_AXL:Box\"><Argument Name=\"Message\">=[MacroError].[Description]</Argument></Ac"
                                "tion></Statements></If></ConditionalBlock></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b0a090ff302010ff302010ff302010ff302010ff302010ff ,
                        0x302010ff302010ff302010ff302010ff302010ff302010ff302010ff00000000 ,
                        0x0000000000000000b0a090fffff8f0fffff0f0ffffe8e0fff0e8e0fff0e0d0ff ,
                        0xf0d8d0fff0d8c0fff0d8c0fff0d8c0fff0d8c0fff0d8c0ff302010ff00000000 ,
                        0x0000000000000000b0a090ffffffffffe06830ffe06830ffe06830ffd06830ff ,
                        0xd06830ffd06830ffd06030ffc06030ff904820ffffe0d0ff302010ff00000000 ,
                        0x0000000000000000b0a090ffffffffffd06830ffffb080ffffa880ffffa070ff ,
                        0xf09870fff09060ffa0b0f0ff1020e0ffc0c8f0ffffe0d0ff302010ff00000000 ,
                        0x00000000a0a8f0ffb0a090ffffffffffe06830ffe06830ffe06830ffd06830ff ,
                        0xd06830ffe0e0f0ff0028ffff1028f0ff4050d0ffffe0d0ff302010ff00000000 ,
                        0x4050e0ff0010b0ffb0a090ffffffffffffffffffffffffffffffffffffffffff ,
                        0xfff8f0ffffe8e0ff2048ffff1038ffff1028ffffe0e8f0ff302010ff7088f0ff ,
                        0x0018c0ff6078f0ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ff ,
                        0xb0a090ffb0a090ffe0e0f0ff3050ffff2040ffff8090f0ffb0b8f0ff0028f0ff ,
                        0x4058f0ff00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000d0d8f0ff4060ffff3050ffff2040ffff3050ffff ,
                        0xe0e8f0ff00000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c0d0f0ff4068ffff4060ffffc0c8f0ff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000c0c8f0ff6078ffff6078ffff6080ffff5070ffff ,
                        0xe0e0f0ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b0b8f0ff6078ffff6078ffffb0c0f0fff0f0f0ff7088ffff ,
                        0x6078ffffc0d0f0ff000000000000000000000000000000000000000000000000 ,
                        0x0000000090a0ffff6078ffff6078ffffd0d8f0ff000000000000000000000000 ,
                        0xb0b8f0ff8098ffff000000000000000000000000000000000000000000000000 ,
                        0x000000008098ffff6080ffffd0d8f0ff00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =360
                    LayoutCachedWidth =480
                    LayoutCachedHeight =660
                    Gradient =0
                    BackThemeColorIndex =1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =3
                    WebImagePaddingBottom =3
                End
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =360
                    Width =2640
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblMasterPlantSpecies_x"
                    Caption ="Master Plant Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =780
                    LayoutCachedTop =360
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =660
                End
                Begin Label
                    OverlapFlags =85
                    Left =3660
                    Top =360
                    Width =1740
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLUCode_x"
                    Caption ="Master Plant Code"
                    GridlineColor =10921638
                    LayoutCachedLeft =3660
                    LayoutCachedTop =360
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =660
                End
                Begin Label
                    OverlapFlags =85
                    Left =5760
                    Top =360
                    Width =1740
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblUTSpecies_x"
                    Caption ="Utah Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =5760
                    LayoutCachedTop =360
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =660
                End
                Begin Label
                    OverlapFlags =85
                    Left =7800
                    Top =360
                    Width =1740
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCOSpecies_x"
                    Caption ="CO Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =7800
                    LayoutCachedTop =360
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =660
                End
                Begin Label
                    OverlapFlags =85
                    Left =9720
                    Top =360
                    Width =1740
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblWYSpecies_x"
                    Caption ="WY Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =9720
                    LayoutCachedTop =360
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =660
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =11760
                    Top =420
                    TabIndex =1
                    BorderColor =10921638
                    Name ="cbxTransectOnly_x"
                    ControlSource ="Transect_Only"
                    DefaultValue ="0"
                    ControlTipText ="Select if target species should be monitored on Transect Only"
                    GridlineColor =10921638

                    LayoutCachedLeft =11760
                    LayoutCachedTop =420
                    LayoutCachedWidth =12020
                    LayoutCachedHeight =660
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =12120
                    Top =360
                    Width =1920
                    Height =300
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cbxTgtAreas_x"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [tbl_Target_Areas].[Target_Area_ID], [tbl_Target_Areas].[Target_Area] FRO"
                        "M tbl_Target_Areas ORDER BY [Target_Area]; "
                    ColumnWidths ="0;1440"
                    GridlineColor =10921638
                    ListItemsEditForm ="frmTgtAreas"

                    LayoutCachedLeft =12120
                    LayoutCachedTop =360
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =660
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =14160
                    Top =360
                    Width =238
                    Height =346
                    FontSize =10
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =29
                    RightMargin =29
                    BottomMargin =29
                    BackColor =13882323
                    BorderColor =8355711
                    ForeColor =8224125
                    Name ="lblAddTgtArea"
                    Caption ="+"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add new target area"
                    GridlineColor =10921638
                    LayoutCachedLeft =14160
                    LayoutCachedTop =360
                    LayoutCachedWidth =14398
                    LayoutCachedHeight =706
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =720
                    Width =2640
                    Height =300
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTgtSpeciesName"
                    ControlSource ="Species_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =780
                    LayoutCachedTop =720
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =1020
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3660
                    Top =720
                    Width =2640
                    Height =300
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text21"
                    ControlSource ="Master_Plant_Code_FK"
                    GridlineColor =10921638

                    LayoutCachedLeft =3660
                    LayoutCachedTop =720
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    Width =2640
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPark"
                    Caption ="Park"
                    GridlineColor =10921638
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =300
                End
                Begin Label
                    OverlapFlags =85
                    Left =2940
                    Width =2640
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear"
                    Caption ="Year"
                    GridlineColor =10921638
                    LayoutCachedLeft =2940
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =300
                End
            End
        End
        Begin FormFooter
            Height =360
            Name ="FormFooter"
            AutoHeight =1
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
' MODULE:       Form_frm_Tgt_Species_List
' Description:  Target species functions & procedures
'
' Source/date:  Bonnie Campbell, 2/11/2015
' Revisions:    BLC - 2/11/2015 - initial version
' =================================

' ---------------------------------
' SUB:          lblAddTgtArea_Click
' Description:  Open add target area form
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 11, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/11/2015 - initial version
'   BLC, 4/30/2015 - integrated into Invasives Reporting tool & updated form naming
' ---------------------------------
Private Sub lblAddTgtArea_Click()
    DoCmd.OpenForm "frm_Tgt_Areas", acNormal
End Sub
