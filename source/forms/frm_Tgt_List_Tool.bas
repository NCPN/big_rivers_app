Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =12300
    DatasheetFontHeight =10
    ItemSuffix =194
    Left =3525
    Right =15825
    Bottom =10560
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x80d28b4cb201e340
    End
    RecordSource ="tsys_App_Defaults"
    Caption ="NCPN Invasives Reporting Tool"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnGotFocus ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
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
        Begin BoundObjectFrame
            AddColon = NotDefault
            SizeMode =3
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =-1800
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin CustomControl
            SpecialEffect =2
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =10575
            BackColor =11830108
            Name ="Detail"
            AlternateBackColor =15710311
            Begin
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =540
                    Top =1425
                    Width =9360
                    Height =485
                    Name ="optGrpTabs"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =540
                    LayoutCachedTop =1425
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =1910
                    Begin
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =540
                            Top =1425
                            Width =1740
                            Height =480
                            FontWeight =700
                            OptionValue =1
                            ForeColor =6250335
                            Name ="tglCreate"
                            StatusBarText ="Create options"
                            Caption ="CREATE"
                            ControlTipText ="Create options"

                            LayoutCachedLeft =540
                            LayoutCachedTop =1425
                            LayoutCachedWidth =2280
                            LayoutCachedHeight =1905
                            UseTheme =1
                            Shape =2
                            BackColor =16772300
                            OldBorderStyle =0
                            HoverColor =10092492
                            PressedColor =16772300
                            HoverForeColor =9974127
                            PressedForeColor =0
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =2280
                            Top =1425
                            Width =1740
                            Height =480
                            FontWeight =700
                            TabIndex =1
                            OptionValue =2
                            ForeColor =6250335
                            Name ="tglSearch"
                            StatusBarText ="Search options"
                            Caption ="SEARCH"
                            ControlTipText ="Search options"

                            LayoutCachedLeft =2280
                            LayoutCachedTop =1425
                            LayoutCachedWidth =4020
                            LayoutCachedHeight =1905
                            UseTheme =1
                            Shape =2
                            BackColor =12835293
                            BackThemeColorIndex =3
                            BackShade =90.0
                            OldBorderStyle =0
                            HoverColor =10092492
                            PressedColor =12835293
                            PressedThemeColorIndex =3
                            PressedShade =90.0
                            HoverForeColor =9974127
                            PressedForeColor =0
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =5760
                            Top =1425
                            Width =1740
                            Height =480
                            FontWeight =700
                            TabIndex =2
                            OptionValue =4
                            ForeColor =6250335
                            Name ="tglExport"
                            StatusBarText ="Export options"
                            Caption ="EXPORT"
                            ControlTipText ="Export options"

                            LayoutCachedLeft =5760
                            LayoutCachedTop =1425
                            LayoutCachedWidth =7500
                            LayoutCachedHeight =1905
                            UseTheme =1
                            Shape =2
                            BackColor =13434828
                            OldBorderStyle =0
                            HoverColor =10092492
                            PressedColor =13434828
                            HoverForeColor =9974127
                            PressedForeColor =0
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =7500
                            Top =1425
                            Width =1740
                            Height =480
                            FontWeight =700
                            TabIndex =3
                            OptionValue =5
                            ForeColor =6250335
                            Name ="tglDbAdmin"
                            StatusBarText ="Db Admin options"
                            Caption ="DB ADMIN"
                            ControlTipText ="Db Admin options"

                            LayoutCachedLeft =7500
                            LayoutCachedTop =1425
                            LayoutCachedWidth =9240
                            LayoutCachedHeight =1905
                            UseTheme =1
                            Shape =2
                            BackColor =14806254
                            BackThemeColorIndex =3
                            OldBorderStyle =0
                            HoverColor =10092492
                            PressedColor =14211288
                            HoverForeColor =9974127
                            PressedForeColor =0
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =4020
                            Top =1425
                            Width =1740
                            Height =480
                            FontWeight =700
                            TabIndex =4
                            OptionValue =3
                            ForeColor =6250335
                            Name ="tglReports"
                            StatusBarText ="Report options"
                            Caption ="REPORTS"
                            ControlTipText ="Report options"

                            LayoutCachedLeft =4020
                            LayoutCachedTop =1425
                            LayoutCachedWidth =5760
                            LayoutCachedHeight =1905
                            UseTheme =1
                            Shape =2
                            BackColor =16764108
                            OldBorderStyle =0
                            HoverColor =10092492
                            PressedColor =16764108
                            HoverForeColor =9974127
                            PressedForeColor =0
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Overlaps =1
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =255
                    Left =480
                    Top =1800
                    Width =11160
                    Height =6840
                    BackColor =16764108
                    Name ="rctMainTabColor"
                    LayoutCachedLeft =480
                    LayoutCachedTop =1800
                    LayoutCachedWidth =11640
                    LayoutCachedHeight =8640
                End
                Begin Label
                    OverlapFlags =93
                    Left =60
                    Top =60
                    Width =3780
                    Height =600
                    FontSize =12
                    FontWeight =600
                    ForeColor =14277081
                    Name ="lblNetwork"
                    Caption ="Target Species List Tool"
                    FontName ="Tahoma"
                    ShortcutMenuBar ="Double-click to open website"
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                End
                Begin Label
                    OverlapFlags =255
                    Left =6420
                    Top =2175
                    Width =4980
                    Height =2340
                    FontSize =11
                    BorderColor =16777215
                    ForeColor =8355711
                    Name ="lblInstructions"
                    Caption ="Choose the park and year for your list. Click 'Continue' to prepare your report."
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =6420
                    LayoutCachedTop =2175
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =4515
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =255
                    Left =780
                    Top =2175
                    Width =5460
                    Height =5400
                    BackColor =13434879
                    BorderColor =12835293
                    Name ="rctLeftInset"
                    GridlineColor =10921638
                    LayoutCachedLeft =780
                    LayoutCachedTop =2175
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =7575
                    BorderThemeColorIndex =3
                    BorderShade =90.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Tab
                    OverlapFlags =255
                    BackStyle =0
                    Style =2
                    Top =480
                    Width =12240
                    Height =8640
                    TabIndex =1
                    Name ="tabMain"

                    LayoutCachedTop =480
                    LayoutCachedWidth =12240
                    LayoutCachedHeight =9120
                    UseTheme =1
                    Shape =2
                    BackColor =9211020
                    OldBorderStyle =0
                    Begin
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =75
                            Top =555
                            Width =12090
                            Height =8490
                            BorderColor =10921638
                            Name ="tabCreate"
                            StatusBarText ="Create"
                            ControlTipText ="Click to create lists"
                            Caption ="QA/QC"
                            GridlineColor =10921638
                            LayoutCachedLeft =75
                            LayoutCachedTop =555
                            LayoutCachedWidth =12165
                            LayoutCachedHeight =9045
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Rectangle
                                    BackStyle =1
                                    OverlapFlags =255
                                    Left =2280
                                    Top =3000
                                    Width =2400
                                    Height =2880
                                    BackColor =8829910
                                    BorderColor =12835293
                                    Name ="rctInsetCtr_Create"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =2280
                                    LayoutCachedTop =3000
                                    LayoutCachedWidth =4680
                                    LayoutCachedHeight =5880
                                    BorderThemeColorIndex =3
                                    BorderShade =90.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    FontUnderline = NotDefault
                                    BorderWidth =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2280
                                    Top =3660
                                    Width =2400
                                    Height =420
                                    FontSize =11
                                    BackColor =8648143
                                    ForeColor =16711680
                                    Name ="btnCreateTgtLists"
                                    Caption ="Create Target List"
                                    FontName ="Calibri"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Create species target list"
                                    LayoutCachedLeft =2280
                                    LayoutCachedTop =3660
                                    LayoutCachedWidth =4680
                                    LayoutCachedHeight =4080
                                    ThemeFontIndex =1
                                End
                                Begin Label
                                    FontUnderline = NotDefault
                                    BorderWidth =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2280
                                    Top =4320
                                    Width =2399
                                    Height =420
                                    FontSize =11
                                    ForeColor =16711680
                                    Name ="btnAddTgtArea"
                                    Caption ="Add Target Area(s)"
                                    FontName ="Calibri"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Add target area"
                                    LayoutCachedLeft =2280
                                    LayoutCachedTop =4320
                                    LayoutCachedWidth =4679
                                    LayoutCachedHeight =4740
                                    ThemeFontIndex =1
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =75
                            Top =555
                            Width =12090
                            Height =8490
                            BorderColor =10921638
                            Name ="tabSearch"
                            Caption ="SEARCH"
                            GridlineColor =10921638
                            LayoutCachedLeft =75
                            LayoutCachedTop =555
                            LayoutCachedWidth =12165
                            LayoutCachedHeight =9045
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Rectangle
                                    BackStyle =1
                                    OverlapFlags =255
                                    Left =2279
                                    Top =2999
                                    Width =2400
                                    Height =2880
                                    BackColor =8829910
                                    BorderColor =12835293
                                    Name ="rctInsetCtr_Logs"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =2279
                                    LayoutCachedTop =2999
                                    LayoutCachedWidth =4679
                                    LayoutCachedHeight =5879
                                    BorderThemeColorIndex =3
                                    BorderShade =90.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    FontUnderline = NotDefault
                                    BorderWidth =1
                                    OverlapFlags =255
                                    TextAlign =2
                                    Left =2205
                                    Top =4455
                                    Width =2415
                                    Height =420
                                    FontSize =11
                                    ForeColor =16711680
                                    Name ="btnEditLog"
                                    Caption ="Edit Log"
                                    FontName ="Calibri"
                                    ControlTipText ="Open edit log"
                                    LayoutCachedLeft =2205
                                    LayoutCachedTop =4455
                                    LayoutCachedWidth =4620
                                    LayoutCachedHeight =4875
                                    ThemeFontIndex =1
                                End
                                Begin Rectangle
                                    SpecialEffect =0
                                    BackStyle =1
                                    OverlapFlags =255
                                    Left =2279
                                    Top =2999
                                    Width =2399
                                    Height =2880
                                    BackColor =8829910
                                    BorderColor =12835293
                                    Name ="rctInsetCtr_Search"
                                    LayoutCachedLeft =2279
                                    LayoutCachedTop =2999
                                    LayoutCachedWidth =4678
                                    LayoutCachedHeight =5879
                                    BorderThemeColorIndex =3
                                    BorderShade =90.0
                                End
                                Begin Label
                                    FontUnderline = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2280
                                    Top =3540
                                    Width =2415
                                    Height =420
                                    FontSize =11
                                    ForeColor =16711680
                                    Name ="btnSearch"
                                    Caption ="Species Search"
                                    FontName ="Calibri"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open species search"
                                    LayoutCachedLeft =2280
                                    LayoutCachedTop =3540
                                    LayoutCachedWidth =4695
                                    LayoutCachedHeight =3960
                                    ThemeFontIndex =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =75
                            Top =555
                            Width =12090
                            Height =8490
                            BorderColor =10921638
                            Name ="tabReports"
                            Caption ="REPORTS"
                            GridlineColor =10921638
                            LayoutCachedLeft =75
                            LayoutCachedTop =555
                            LayoutCachedWidth =12165
                            LayoutCachedHeight =9045
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Rectangle
                                    BackStyle =1
                                    OverlapFlags =255
                                    Left =2280
                                    Top =3000
                                    Width =2400
                                    Height =2880
                                    BackColor =8829910
                                    BorderColor =12835293
                                    Name ="rctInsetLeft_Reports"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =2280
                                    LayoutCachedTop =3000
                                    LayoutCachedWidth =4680
                                    LayoutCachedHeight =5880
                                    BorderThemeColorIndex =3
                                    BorderShade =90.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    FontUnderline = NotDefault
                                    BackStyle =1
                                    BorderWidth =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2305
                                    Top =3595
                                    Width =2340
                                    Height =420
                                    FontSize =11
                                    TopMargin =72
                                    BackColor =8648143
                                    ForeColor =16711680
                                    Name ="btnCrewSpeciesList"
                                    Caption ="Crew Species List"
                                    FontName ="Calibri"
                                    OnClick ="[Event Procedure]"
                                    ShortcutMenuBar ="Run Species List report"
                                    ControlTipText ="Prepare field crew target species list"
                                    LayoutCachedLeft =2305
                                    LayoutCachedTop =3595
                                    LayoutCachedWidth =4645
                                    LayoutCachedHeight =4015
                                    ThemeFontIndex =1
                                End
                                Begin Label
                                    FontUnderline = NotDefault
                                    BorderWidth =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2305
                                    Top =4135
                                    Width =2340
                                    Height =420
                                    FontSize =11
                                    TopMargin =72
                                    ForeColor =16711680
                                    Name ="btnSpeciesListByPark"
                                    Caption ="Species List by Park"
                                    FontName ="Calibri"
                                    OnClick ="[Event Procedure]"
                                    ShortcutMenuBar ="Run target species list by park report"
                                    ControlTipText ="Run target species list by park report"
                                    LayoutCachedLeft =2305
                                    LayoutCachedTop =4135
                                    LayoutCachedWidth =4645
                                    LayoutCachedHeight =4555
                                    ThemeFontIndex =1
                                End
                                Begin Label
                                    FontUnderline = NotDefault
                                    BorderWidth =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2305
                                    Top =4675
                                    Width =2340
                                    Height =420
                                    FontSize =11
                                    TopMargin =72
                                    BackColor =8648143
                                    ForeColor =16711680
                                    Name ="btnTgtListAnnualSummary"
                                    Caption ="Annual Summary"
                                    FontName ="Calibri"
                                    OnClick ="[Event Procedure]"
                                    ShortcutMenuBar ="Run species list annual summary report"
                                    ControlTipText ="Run species list annual summary report"
                                    LayoutCachedLeft =2305
                                    LayoutCachedTop =4675
                                    LayoutCachedWidth =4645
                                    LayoutCachedHeight =5095
                                    ThemeFontIndex =1
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =75
                            Top =555
                            Width =12090
                            Height =8490
                            BorderColor =10921638
                            Name ="tabExport"
                            Caption ="EXPORT"
                            GridlineColor =10921638
                            LayoutCachedLeft =75
                            LayoutCachedTop =555
                            LayoutCachedWidth =12165
                            LayoutCachedHeight =9045
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =75
                            Top =555
                            Width =12090
                            Height =8490
                            BorderColor =10921638
                            Name ="tabDbAdmin"
                            StatusBarText ="Database Administration"
                            ControlTipText ="Database Administration Tasks"
                            Caption ="DB ADMIN"
                            GridlineColor =10921638
                            LayoutCachedLeft =75
                            LayoutCachedTop =555
                            LayoutCachedWidth =12165
                            LayoutCachedHeight =9045
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    BorderWidth =3
                                    Left =720
                                    Top =2775
                                    Width =10800
                                    Height =5340
                                    Name ="fsub_DbAdmin"
                                    SourceObject ="Form.fsub_DbAdmin"
                                    StatusBarText ="Database Administration"

                                    LayoutCachedLeft =720
                                    LayoutCachedTop =2775
                                    LayoutCachedWidth =11520
                                    LayoutCachedHeight =8115
                                End
                                Begin Label
                                    OverlapFlags =255
                                    Left =600
                                    Top =2640
                                    Width =10980
                                    Height =5640
                                    BackColor =12835293
                                    Name ="lblCover"
                                    LayoutCachedLeft =600
                                    LayoutCachedTop =2640
                                    LayoutCachedWidth =11580
                                    LayoutCachedHeight =8280
                                    BackShade =75.0
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    Left =840
                    Top =2235
                    Width =5100
                    Height =315
                    FontSize =14
                    BorderColor =16777215
                    ForeColor =5855577
                    Name ="lblInsetTitle"
                    Caption ="Reports > Field Crew Species List"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedLeft =840
                    LayoutCachedTop =2235
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =2550
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Subform
                    OverlapFlags =247
                    Left =6540
                    Top =4620
                    Width =4170
                    Height =2895
                    TabIndex =2
                    Name ="fsub_Filter"
                    SourceObject ="Form.fsub_Select_Tgt_Lists"

                    LayoutCachedLeft =6540
                    LayoutCachedTop =4620
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =7515
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
Option Explicit

' =================================
' FORM NAME:    frm_Tgt_List_Tool
' Description:  Standard form - main screen of the target list tool user interface
' Data source:  tsys_App_Defaults
' Data access:  edit only, no additions, moving between records or deletions
' Pages:        pgDefaults, pgAbout, pgSettings
' Functions:    none
' References:   AppSetup, DeleteFile, FileExists, MakeBackup, ParseFileExt,
'                   SaveFile, TableExists, VerifyConnections
' Source/date:  John R. Boetsch, May 24, 2006
' Adapted/date: Bonnie L. Campbell, April 21, 2015 from frm_Switchboard (WQ Utilities tool)
'               --------------------------------------------------------------------------------------
' Revisions:    BLC, 4/21/2015 - Adapted for NCPN Invasives Reports - Species Target List tool
'                   Converted tabQAQC to tabCreate, tabLogs to tabView
'               BLC, 4/22/2015 - Renamed
'               BLC, 6/4/2015  - Changed View to Search tab
' =================================

' ********************************************************************************************
'  NOTE: When making design changes to this form (e.g., changing the name of a control),
'   certain standard module functions (especially fxnAppSetup) may need to be updated so that
'   any external references to the controls in this form are correct. When renaming controls,
'   it may be a good idea to do global find/replace; when adding or removing controls, more
'   careful searching may be necessary.
' ********************************************************************************************

' =================================
'   Properties
' =================================

' =================================
'   Main Target List Tool Form Functionality
' =================================

' ---------------------------------
' SUB:     Form_Open
' Description:  Open form & run table checks
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Adapted from John Boetsch
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/31/2014 - updated to use TempVars vs gvars
'               BLC, 8/25/2014 - added 'update' flag for setUserAcess to ensure app mode settings are updated
'               --------------------------------------------------------------------------------------
'               BLC, 4/21/2015 - Adapted for NCPN Invasives Reports - Species Target List tool
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    Dim strSQL As String
    Dim aryFields As Variant

    'initialize app settings --> initApp moved to subform fsub_DbAdmin
    '                            since subform loads first *BEFORE* form

    'set default app mode & initialize controls
    setUserAccess Me, "update"
    
    ' Initialize UI
    optGrpTabs.Value = 1 'Create
    
    Me.optGrpTabs_Click
    Me!fsub_Filter.visible = False
    
Exit_Sub:
    DoCmd.SetWarnings True
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Form_frm_Tgt_List_Tool])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:     Form_GotFocus
' Description:  Handle UI display for slow startup
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Adapted from John Boetsch
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/31/2014 - initial version
' ---------------------------------
Private Sub Form_GotFocus()
    On Error GoTo Err_Handler

    ' Helps if startup is slow and user opens another window in the meantime
    Me.Repaint

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_Current
' Description:  Updates subform fields & displays
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, June 17, 2014
' Revisions:    6/17/2014 - BLC - initial version
' ---------------------------------
Private Sub Form_Current()
    Me.Repaint
End Sub

' ---------------------------------
' SUB:     cmdExit_Click
' Description:  Compact, repair & exit database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Adapted from John Boetsch
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/31/2014 - updated to use TempVars vs gvars
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub cmdExit_Click()
    On Error GoTo Err_Handler

    ' Prompt for backups, depending on application settings
    '   Note:  only relevant for Access back-end files
    If TempVars("Connected") And TempVars("HasAccessBE") And Me.chkBackupOnExit Then MakeBackup

    ' Compact and repair back-end database prior to exit, depending on
    '   default settings and on whether there is a valid link to the database
    '   Note:  only relevant for Access back-end files
    If TempVars("Connected") And TempVars("HasAccessBE") And TempVars("WritePermission") And Me.chkCompactBEOnExit Then

        Dim rst As DAO.Recordset
        Dim strOrigFile As String
        Dim strFileExt As String
        Dim strNewFile As String
        Dim intCount As Integer

On Error GoTo Quit_procedure
        ' Set the recordset to the system table
        Set rst = CurrentDb.OpenRecordset("SELECT * " & _
            "FROM tsys_Link_Dbs;", dbOpenSnapshot)

        rst.MoveFirst
        Do Until rst.EOF
            ' Check for empty string or non-existent file first
            If IsNull(rst![File_path]) = False Then
                strOrigFile = rst![File_path]
                strFileExt = ParseFileExt(strOrigFile)
                If FileExists(strOrigFile) Then
                    intCount = 0
                    ' If needed, loop through temporary name alternatives until an unused
                    '   name is found
                    Do
                        intCount = intCount + 1
                        strNewFile = Left(strOrigFile, Len(strOrigFile) - Len(strFileExt)) _
                            & "_" & CStr(intCount) & strFileExt
                    Loop Until FileExists(strNewFile) = False
                    DBEngine.CompactDatabase strOrigFile, strNewFile
                    ' If successful deleting the original, uncompacted file then rename
                    '   the compacted file to the original name
                    If DeleteFile(strOrigFile) Then Name strNewFile As strOrigFile
                End If
            End If
            rst.MoveNext
        Loop
    End If

Quit_procedure:
    ' Close the application
    DoCmd.Quit acQuitSaveNone

Exit_Procedure:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3356, 70
        ' The back-end database is already open when trying to compact ...
        DoCmd.Quit acQuitSaveNone
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered exiting the main menu"
        DoCmd.Quit acQuitSaveNone
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_Close
' Description: Close switchboard/main form & log user
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Adapted from John Boetsch
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/31/2014 - updated to use TempVars vs gvars
'               BLC, 5/27/2015 - added check for if main form is open (frm_Main_Menu)
' ---------------------------------
Private Sub Form_Close()
    On Error GoTo Err_Handler

    'log user exit
    logUserAction Me

    'check if form is open
    If FormIsOpen("frm_Main_Menu") Then
        'refocus on main menu
        Forms("frm_Main_Menu").SetFocus
        DoCmd.Restore
    End If
    
Exit_Procedure:
    DoCmd.SetWarnings True
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
'   Main Tab Navigation
' =================================
'  NOTE:
'    To provide colored tabs & tabbed control navigation, standard tab control tabs are
'  replaced with an option group (optGrpTabs) tab control tabs are set to None & background
'  is transparent. A rectangle (rctMainTabColor) is used to color the body of the tab control.
'  rctMainTabColor is below the transparent tabMain and optGrpTabs lies below it to create
'  the appearance of a colored tab & tab control.
' =================================

' =================================
' SUB:     optGrpTabs_Click
' Description:  Changes display of main tab control
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, June 2014
'               Adapted from Tina T, 03/19/2013 post comments
'               http://www.utteraccess.com/forum/Tabbed-Form-Styles-Colou-t2001626.html
'               Created 06/11/2014 blc; Last modified 06/11/2014 blc.
' Revisions:    BLC, 6/16/2014 - Changed to Public Sub (vs Private) to make available to
'                                switchboard open event for initializing tab UI
'               BLC, 8/21/2014 - Added code to unhighlight any highlighted label buttons
'                                on tab change
'               --------------------------------------------------------------------------------------
'               BLC, 4/21/2015 - Adapted for NCPN Invasives Reports - Species Target List tool
'               BLC, 6/4/2015  - Changed View to Search, changed TempVars.item("xx") to TempVars("xx")
'                                Shifted PopulateInsetTitle & set tab color to before check if DbAdmin
'                                to make DbAdmin tab appear correct color when selected
'               BLC, 6/9/2015  - Removed blnLeftInsetHide to expose DbAdmin display similar to Export tab
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' =================================
Public Sub optGrpTabs_Click()
On Error GoTo Err_Handler

Dim strTab As String, strAction As String
Dim lngTabColor As Long
Dim blnLeftInsetHide As Boolean

    'default
    lngTabColor = HTMLConvert("#FFFFFF") 'white
    strAction = ""
    
    'identify current tab
    Select Case optGrpTabs.Value
        Case 1 'Create
            strTab = "tabCreate"
            strAction = "tbl"
            lngTabColor = HTMLConvert("#CCECFF") 'lt blue CCECFF RGB(204,236,255) 13430015
        Case 2 'Search
            strTab = "tabSearch"
            strAction = "frm"
            lngTabColor = RGB(221, 217, 195) 'tan DDD9C3 RGB(221,217,195) 14539203
            'update instructions & make visible
            PopulateInstructions Me!lblInstructions, "Search"
            Me!lblInstructions.visible = True
        Case 3 'Reports
            strTab = "tabReports"
            strAction = "rpt"
            lngTabColor = HTMLConvert("#CCCCFF") 'lt purple CCCCFF RGB(204,187,255) 13417471
        Case 4 'Exports
            strTab = "tabExport"
            strAction = "exp"
            lngTabColor = HTMLConvert("#CCFFCC") 'lt green CCFFCC RGB(204,255,204) 13434828
            'Application.LoadCustomUI "tabExportOptions", GetRibbonXML("Export")
            'update instructions & make visible
            PopulateInstructions Me!lblInstructions, "Export"
            Me!lblInstructions.visible = True
        Case 5 'DB Admin
            strTab = "tabDbAdmin"
            lngTabColor = HTMLConvert("#D8D8D8") 'lt gray D8D8D8 RGB(216,216,216) 14211288
            'blnLeftInsetHide = True
            lblCover.backcolor = lngTabColor
            lblCover.visible = True
            'update instructions & make visible
            PopulateInstructions Me!lblInstructions, "DbAdmin"
            Me!lblInstructions.visible = True
    End Select
    
    'identify action type (qry = query, rpt = report, exp = export, tbl = table, frm = form)
    If IsNull(TempVars("action")) Then
        TempVars.Add "action", strAction
    Else
        TempVars("action") = strAction
    End If
    
    'clear instructions
    If Not (strTab = "tabExport" Or strTab = "tabSearch" Or strTab = "tabDbAdmin") Then
        Me!lblInstructions.Caption = " "
    End If
    
    'uncolor label buttons
    buttonUnHighlight Me!lblInstructions
    
    'hide filter
    Me.Controls.item("fsub_Filter").visible = False
    
    'unhide left inset rectangle
    Me.Controls.item("rctLeftInset").visible = True
    If blnLeftInsetHide Then
        Me.Controls.item("rctLeftInset").visible = False
    End If
    
    'update title & make visible
    PopulateInsetTitle Me!lblInsetTitle, Replace(strTab, "tab", "")
        
    'set tab color
    rctMainTabColor.backcolor = lngTabColor
        
    'keep tabs hidden
    If strTab = "tabView" Or strTab = "tabDbAdmin" Then
        
        ' set hide only to hide other tab insets
        tabPageUnhide tabMain, strTab, True
        GoTo Exit_Sub
    End If
    
    'hide all tabs & unhide current tab
    tabPageUnhide tabMain, strTab
        
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - optGrpTabs_Click[frm_Tgt_List_Tool])"
    End Select
    Resume Exit_Sub
End Sub

' =================================
'   Create Tab Functionality
' =================================
' =================================
' TAB (PAGE) NAME:    CREATE (tabCreate)
' Description:  Create target species lists
' Unbound ctls:
' Subforms:     fsub_Select_Park_Year
' =================================
' ---------------------------------
' SUB:     btnCreateTgtLists_Click
' Description:  Create species target lists
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, April 22, 2015
' Revisions:    4/22/2015 - BLC - initial version
' ---------------------------------
Private Sub btnCreateTgtLists_Click()
    
    'button actions
    ActionButton_Click btnCreateTgtLists
    
End Sub

' ---------------------------------
' SUB:     btnAddTgtArea_Click
' Description:  Add new target area(s)
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, April 22, 2015
' Revisions:    4/22/2015 - BLC - initial version
' ---------------------------------
Private Sub btnAddTgtArea_Click()
    
    'button actions
    ActionButton_Click btnAddTgtArea, "frm_Tgt_Areas"
    
    fsub_Filter.visible = False
       
End Sub

' ---------------------------------
' SUB:          ActionButton_Click
' Description:  Generic form button actions
' Assumptions:  -
' Parameters:   btn - action control (control)
'               strForm - optional form name (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, August 20, 2014 - for NCPN tools
' Adapted:      -
' Revisions:   BLC - 8/20/2014 - initial version
'              BLC - 6/4/2015  - added blnDisableDoubleClick to turn off species search LUCode_DoubleClick event
' ---------------------------------
Public Sub ActionButton_Click(btn As Control, Optional strForm As String)
    Dim strArgs As String
    'default
    strArgs = ""
    
    'highlight btn
    buttonHighlight btn, 1

    'status bar
    Application.SysCmd acSysCmdSetStatus, btn.ControlTipText

    'prepare display
    If Len(strForm) > 0 Then
        If strForm = "frm_Species_Search" Then strArgs = "DisableDoubleClick"
        DoCmd.OpenForm strForm, acNormal, , , , acWindowNormal, strArgs
    Else
        OpenFilterSubform btn
    End If
End Sub

' ---------------------------------
' SUB:          OpenFilterSubform
' Description:  Prepares the filter subform & displays it
' Parameters:   ctrl - control button clicked (control)
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, June 17, 2014
' Revisions:    BLC - 6/17/2014 - initial version
'               BLC - 5/1/2015  - handled alternate fsub_Filter subform control source forms
'               BLC - 5/12/2015  - adapted for unused tabs
'               BLC - 5/13/2015 - fixed so "Select Year" is displayed initially for cbxYear
'               BLC - 6/4/2015  - added "Search" action for alternate fsub_Filter subform control handling
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Public Sub OpenFilterSubform(ctrl As Control)
        
On Error GoTo Err_Handler:
    
    'set breadcrumb hidden values
    Dim strCrumbs As String, strField As String
    strField = ""
    'If ControlExists("lblField", fsub_Filter) Then
        'strField = lblField.Caption & "|"
    'End If

    strCrumbs = tglCreate.Caption & "|" & lblInsetTitle.Caption & "|" _
                    & strField & Replace(ctrl.name, "btn", "")
                    
                    '& lblField.Caption & "|" & Replace(ctrl.name, "btn", "")
    
    'Prepare & open filter
    'DoCmd.OpenForm ("fsub_Filter"), , , , , , strCrumbs
    
    'breadcrumbs
    Dim aryCrumbs As Variant
    aryCrumbs = CrumbsToArray(strCrumbs)
    'PrepareCrumbs Me.fsub_Filter, aryCrumbs
    
    'set & pass variables
    Dim strAction As String
    strAction = Replace(ctrl.name, "btn", "")
    
    'check for existence, add analysis tempvar if needed
    If IsNull(TempVars("analysis")) Then
        TempVars("analysis") = strAction
    Else
        TempVars.Add "analysis", strAction
    End If
    
    'update title & make visible
    PopulateInsetTitle Me!lblInsetTitle, TempVars("analysis")
    
    'update instructions & make visible
    PopulateInstructions Me!lblInstructions, TempVars("analysis")

    'set fsub_Filter subform
    Select Case TempVars("action")
        Case "tbl"
            fsub_Filter.SourceObject = "fsub_Select_Park_Year"
        Case "rpt"
            Select Case TempVars("rpt")
                Case "CrewSpeciesList", "SpeciesListByPark"
                    fsub_Filter.SourceObject = "fsub_Select_Tgt_Lists"
                Case "TgtListAnnualSummary"
                    fsub_Filter.SourceObject = "fsub_Select_Year"
            End Select
    End Select

    'reveal filter subform & clear fields
    With fsub_Filter
        .visible = True
        Select Case fsub_Filter.SourceObject
            Case "fsub_Select_Park_Year"
                !cbxPark.Value = ""
                !cbxYear.Value = "SEL"
            Case "fsub_Select_Tgt_Lists"
            Case "fsub_Select_Year"
        End Select
    End With
    
    'hide filter subform for unused tabs
    If strAction = "DbAdmin" Or strAction = "View" Or strAction = "Search" Then
        lblCover.visible = True
    End If
    
    Me.Repaint

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - OpenFilterSubform[Form_frm_Tgt_List_Tool])"
    End Select
    Resume Exit_Procedure
    
End Sub

' =================================
'   Search Tab Functionality
' =================================
' =================================
' TAB (PAGE) NAME:    Search (tabSearch)
' Description:  Search for species info
' Unbound ctls:
' Subforms:     none
' =================================
' ---------------------------------
' SUB:     btnSearch_Click
' Description:  Opens the search form without double click functionality
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, June 4, 2015
' Revisions:    BLC - 6/4/2015 - initial version
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub btnSearch_Click()

    TempVars("frm") = "Species_Search"

    'button actions
    ActionButton_Click btnSearch, "frm_Species_Search"
    
End Sub

' =================================
'   Reports Tab Functionality
' =================================
' =================================
' TAB (PAGE) NAME:    Reports (tabReports)
' Description:  Run results reports
' Unbound ctls:
' Subforms:     none
' =================================
' ---------------------------------
' SUB:     btnCrewSpeciesList_Click
' Description:  Opens the field crew species target list report
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, April 22, 2015
' Revisions:    BLC - 4/22/2015 - initial version
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub btnCrewSpeciesList_Click()

    TempVars("rpt") = "CrewSpeciesList"

    'button actions
    ActionButton_Click btnCrewSpeciesList
    
End Sub

' ---------------------------------
' SUB:     btnSpeciesListByPark_Click
' Description:  Opens the species target list by park report
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, April 22, 2015
' Revisions:    BLC - 4/22/2015 - initial version
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub btnSpeciesListByPark_Click()

    TempVars("rpt") = "SpeciesListByPark"
    
    'button actions
    ActionButton_Click btnSpeciesListByPark
    
End Sub

' ---------------------------------
' SUB:     btnTgtListAnnualSummary_Click
' Description:  Opens the species target list annual summary report
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, April 22, 2015
' Revisions:    BLC - 4/22/2015 - initial version
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub btnTgtListAnnualSummary_Click()

    TempVars("rpt") = "TgtListAnnualSummary"

    'button actions
    ActionButton_Click btnTgtListAnnualSummary
    
End Sub

' =================================
'   Export Tab Functionality
' =================================
' =================================
' TAB (PAGE) NAME:    Export (tabExport)
' Description:  Export results
' Unbound ctls:
' Subforms:     none
' =================================
'  see instructions

' =================================
'   DbAdmin Tab Functionality
' =================================
' =================================
' TAB (PAGE) NAME:    DbAdmin (tabDbAdmin)
' Description:  Db Administrative functions
' Unbound ctls:
' Subforms:     fsub_DbAdmin
' =================================
'  see subform
