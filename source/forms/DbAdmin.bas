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
    ItemSuffix =72
    Left =4275
    Top =2730
    Right =12135
    Bottom =9375
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x7982ea74fa02e540
    End
    RecordSource ="SELECT Project, Release_ID, DataTimeframe, UserName, Park, BackupPromptOnStartUp"
        ", BackupPromptOnExit, CompactBEOnExit, WebURL, AppContactName, AppContactOrg, Ap"
        "pContactPhone, AppContactEmail, ar.ReleaseDate, ar.IsSupported, DatabaseTitle, V"
        "ersionNumber, FileName FROM tsys_App_Defaults AS ad INNER JOIN tsys_App_Releases"
        " AS ar ON ar.ID = ad.Release_ID; "
    Caption ="NCPN Big Rivers App"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnActivate ="[Event Procedure]"
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
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =840
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =60
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Db Admin"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =180
                    Top =420
                    Width =6840
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Choose the desired action below."
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =6960
                    Top =180
                    Width =720
                    ForeColor =16711680
                    Name ="btnComment"
                    Caption ="í ½í·©"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Enter comment"
                    GridlineColor =10921638

                    LayoutCachedLeft =6960
                    LayoutCachedTop =180
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =540
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
                Begin Line
                    OverlapFlags =87
                    SpecialEffect =5
                    Top =735
                    Width =7859
                    Name ="lnHdr"
                    LayoutCachedTop =735
                    LayoutCachedWidth =7859
                    LayoutCachedHeight =735
                    BorderThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ToggleButton
                    OverlapFlags =215
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =5370
                    Top =150
                    Width =1500
                    Height =420
                    ColumnOrder =0
                    FontSize =8
                    FontWeight =500
                    TabIndex =1
                    Name ="tglDevMode"
                    StatusBarText ="Turn DEV MODE on or off"
                    DefaultValue ="False"
                    Caption ="DEV MODE"
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Turn DEV MODE on or off"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5370
                    LayoutCachedTop =150
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =570
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =1
                    Bevel =0
                    Gradient =12
                    BackColor =5880731
                    BackThemeColorIndex =6
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderColor =5880731
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =15788753
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =9699294
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16724787
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16724787
                    PressedForeThemeColorIndex =-1
                    Shadow =-1
                    QuickStyle =25
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5820
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Width =7860
                    Height =5820
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =5820
                    BackThemeColorIndex =-1
                End
                Begin Line
                    OverlapFlags =87
                    SpecialEffect =5
                    Left =3345
                    Top =3240
                    Width =1872
                    Name ="lnDbAdmin"
                    LayoutCachedLeft =3345
                    LayoutCachedTop =3240
                    LayoutCachedWidth =5217
                    LayoutCachedHeight =3240
                    BorderThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5640
                    Top =3840
                    Width =1800
                    Height =405
                    ForeColor =4210752
                    Name ="btnX"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =3840
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =4245
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
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =3420
                    Top =1320
                    Width =1800
                    Height =405
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnNavCoords"
                    Caption ="Navigation coords"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Generate navigation target coordinates for upload to GPS"
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =1320
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1725
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =84
                    Left =3420
                    Top =840
                    Width =1800
                    Height =405
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnTaskList"
                    Caption ="&Task list"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="View the list of tasks associated with sample locations"
                    UnicodeAccessKey =84
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =840
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1245
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =69
                    Left =3420
                    Top =360
                    Width =1800
                    Height =405
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnEnter"
                    Caption ="&Enter / edit data"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open the gateway form to access the data entry forms"
                    UnicodeAccessKey =69
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =360
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =765
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =80
                    Left =5640
                    Top =3300
                    Width =1800
                    Height =405
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnUISetup"
                    Caption ="UI Setu&p"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Update back-end database connections"
                    UnicodeAccessKey =112
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =3300
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =3705
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5640
                    Top =360
                    Width =1800
                    Height =405
                    TabIndex =5
                    ForeColor =4210752
                    Name ="btnSummaries"
                    Caption ="Data summaries"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open the data summarization tool"
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =360
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =765
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5640
                    Top =2340
                    Width =1800
                    Height =405
                    TabIndex =6
                    ForeColor =4210752
                    Name ="btnQAReport"
                    Caption ="Quality review rpt"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="View the quality review results as a report"
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =2340
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =2745
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5640
                    Top =1860
                    Width =1800
                    Height =405
                    TabIndex =7
                    ForeColor =4210752
                    Name ="btnTaskListRpt"
                    Caption ="Task list report"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Generate the task list report"
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =1860
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =2265
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5640
                    Top =1380
                    Width =1800
                    Height =405
                    TabIndex =8
                    ForeColor =4210752
                    Name ="btnNavReport"
                    Caption ="Navigation report"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Generate the sample location navigation report"
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =1380
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =1785
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5640
                    Top =900
                    Width =1800
                    Height =405
                    TabIndex =9
                    ForeColor =4210752
                    Name ="btnSpeciesListRpt"
                    Caption ="Species lists"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Generate transect species lists"
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =900
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =1305
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5880
                    Top =4740
                    Width =1800
                    Height =405
                    TabIndex =10
                    ForeColor =4210752
                    Name ="btnEditLog"
                    Caption ="Edit log"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Log edits to certified data"
                    GridlineColor =10921638

                    LayoutCachedLeft =5880
                    LayoutCachedTop =4740
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =5145
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =81
                    Left =3960
                    Top =4740
                    Width =1800
                    Height =405
                    TabIndex =11
                    ForeColor =4210752
                    Name ="btnQA"
                    Caption ="&QA checks"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open the data validation and quality review tool"
                    UnicodeAccessKey =81
                    GridlineColor =10921638

                    LayoutCachedLeft =3960
                    LayoutCachedTop =4740
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =5145
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =76
                    Left =2040
                    Top =4740
                    Width =1800
                    Height =405
                    TabIndex =12
                    ForeColor =4210752
                    Name ="btnLookups"
                    Caption ="&Lookup tables"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Manage lookup domains"
                    UnicodeAccessKey =76
                    GridlineColor =10921638

                    LayoutCachedLeft =2040
                    LayoutCachedTop =4740
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =5145
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =68
                    Left =120
                    Top =4740
                    Width =1800
                    Height =405
                    TabIndex =13
                    ForeColor =4210752
                    Name ="btnBrowser"
                    Caption ="&Data browser"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open the project data browser to view data by sample location"
                    UnicodeAccessKey =68
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =4740
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =5145
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =66
                    Left =3420
                    Top =3850
                    Width =1800
                    Height =405
                    TabIndex =14
                    ForeColor =4210752
                    Name ="btnBackup"
                    Caption ="&Backup data"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Make a backup copy of the back-end database(s)"
                    UnicodeAccessKey =66
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =3850
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =4255
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =79
                    Left =3420
                    Top =3360
                    Width =1800
                    Height =405
                    TabIndex =15
                    ForeColor =4210752
                    Name ="btnDbWindow"
                    Caption ="View db &objects"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="View the database object navigation pane:  tables, queries, forms"
                    UnicodeAccessKey =111
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =3360
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =3765
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =85
                    Left =3420
                    Top =2740
                    Width =1800
                    Height =405
                    TabIndex =16
                    ForeColor =4210752
                    Name ="btnSetRoles"
                    Caption ="Set &user roles"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set database user access roles"
                    UnicodeAccessKey =117
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =2740
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =3145
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
                    TabStop = NotDefault
                    OverlapFlags =223
                    AccessKey =67
                    Left =3420
                    Top =2250
                    Width =1800
                    Height =405
                    TabIndex =17
                    ForeColor =4210752
                    Name ="btnReconnect"
                    Caption ="Db &connections"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Update back-end database connections"
                    UnicodeAccessKey =99
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =2250
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =2655
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
                    Left =3360
                    Top =60
                    Width =1920
                    Height =255
                    FontSize =9
                    FontWeight =700
                    BackColor =11056034
                    ForeColor =16777215
                    Name ="lblDataEntry"
                    Caption ="Data Entry and Edits"
                    FontName ="Arial"
                    LayoutCachedLeft =3360
                    LayoutCachedTop =60
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =315
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =5760
                    Top =3000
                    Width =1575
                    Height =270
                    FontSize =9
                    FontWeight =700
                    BackColor =11056034
                    ForeColor =16777215
                    Name ="lblUISetup"
                    Caption ="User Interface"
                    FontName ="Arial"
                    LayoutCachedLeft =5760
                    LayoutCachedTop =3000
                    LayoutCachedWidth =7335
                    LayoutCachedHeight =3270
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Line
                    OverlapFlags =87
                    SpecialEffect =5
                    Left =5580
                    Top =825
                    Width =2016
                    Name ="lnSummaries"
                    LayoutCachedLeft =5580
                    LayoutCachedTop =825
                    LayoutCachedWidth =7596
                    LayoutCachedHeight =825
                    BorderThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Line
                    OverlapFlags =87
                    SpecialEffect =5
                    Left =2160
                    Top =4500
                    Width =5472
                    Name ="lnUI"
                    LayoutCachedLeft =2160
                    LayoutCachedTop =4500
                    LayoutCachedWidth =7632
                    LayoutCachedHeight =4500
                    BorderThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =5520
                    Top =60
                    Width =2100
                    Height =255
                    FontSize =9
                    FontWeight =700
                    BackColor =11056034
                    ForeColor =16777215
                    Name ="lblOutput"
                    Caption ="Summaries and Output"
                    FontName ="Arial"
                    LayoutCachedLeft =5520
                    LayoutCachedTop =60
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =315
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Line
                    OverlapFlags =87
                    SpecialEffect =5
                    Left =5520
                    Top =2820
                    Width =2016
                    Name ="lnRpt"
                    LayoutCachedLeft =5520
                    LayoutCachedTop =2820
                    LayoutCachedWidth =7536
                    LayoutCachedHeight =2820
                    BorderThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Line
                    OverlapFlags =87
                    SpecialEffect =5
                    Left =3360
                    Top =1860
                    Width =1872
                    Name ="lnDataEntry"
                    LayoutCachedLeft =3360
                    LayoutCachedTop =1860
                    LayoutCachedWidth =5232
                    LayoutCachedHeight =1860
                    BorderThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =120
                    Top =4380
                    Width =1860
                    Height =255
                    FontSize =9
                    FontWeight =700
                    BackColor =11056034
                    ForeColor =16777215
                    Name ="lblMgmtTools"
                    Caption ="Management Tools"
                    FontName ="Arial"
                    LayoutCachedLeft =120
                    LayoutCachedTop =4380
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =4635
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =3540
                    Top =1980
                    Width =1575
                    Height =270
                    FontSize =9
                    FontWeight =700
                    BackColor =11056034
                    ForeColor =16777215
                    Name ="lblDbAdmin"
                    Caption ="Database Admin"
                    FontName ="Arial"
                    LayoutCachedLeft =3540
                    LayoutCachedTop =1980
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =2250
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Tab
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextFontFamily =18
                    Width =3225
                    Height =4275
                    FontSize =9
                    FontWeight =700
                    TabIndex =18
                    Name ="PageTabs"
                    FontName ="Arial"

                    LayoutCachedWidth =3225
                    LayoutCachedHeight =4275
                    ThemeFontIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    BackColor =14922894
                    BackThemeColorIndex =2
                    BackTint =40.0
                    BackShade =100.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverThemeColorIndex =-1
                    PressedColor =-2147483625
                    PressedThemeColorIndex =-1
                    HoverForeColor =-2147483619
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =-2147483619
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    ForeTint =100.0
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =75
                            Top =435
                            Width =3075
                            Height =3765
                            BorderColor =10921638
                            Name ="pgDefaults"
                            Caption =" Defaults"
                            GridlineColor =10921638
                            LayoutCachedLeft =75
                            LayoutCachedTop =435
                            LayoutCachedWidth =3150
                            LayoutCachedHeight =4200
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Rectangle
                                    OverlapFlags =223
                                    Left =195
                                    Top =795
                                    Width =2820
                                    Height =2805
                                    Name ="boxDefaultPane"
                                    LayoutCachedLeft =195
                                    LayoutCachedTop =795
                                    LayoutCachedWidth =3015
                                    LayoutCachedHeight =3600
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =255
                                    Top =495
                                    Width =1584
                                    Height =252
                                    FontSize =9
                                    FontWeight =700
                                    BackColor =11056034
                                    ForeColor =-2147483603
                                    Name ="lblDefaults"
                                    Caption ="Current defaults"
                                    FontName ="Arial"
                                    LayoutCachedLeft =255
                                    LayoutCachedTop =495
                                    LayoutCachedWidth =1839
                                    LayoutCachedHeight =747
                                    ThemeFontIndex =-1
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderTint =100.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    SpecialEffect =2
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =1395
                                    Top =1395
                                    Width =1500
                                    FontSize =8
                                    BackColor =-2147483629
                                    ForeColor =-2147483607
                                    Name ="tbxTimeframe"
                                    ControlSource ="DataTimeframe"
                                    FontName ="Arial"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =1395
                                    LayoutCachedTop =1395
                                    LayoutCachedWidth =2895
                                    LayoutCachedHeight =1635
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =315
                                            Top =1395
                                            Width =960
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            ForeColor =-2147483603
                                            Name ="lblTimeframe"
                                            Caption ="Timeframe"
                                            FontName ="Arial"
                                            LayoutCachedLeft =315
                                            LayoutCachedTop =1395
                                            LayoutCachedWidth =1275
                                            LayoutCachedHeight =1647
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    SpecialEffect =2
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =1035
                                    Top =1755
                                    Width =1860
                                    FontSize =8
                                    TabIndex =1
                                    BackColor =-2147483629
                                    ForeColor =-2147483607
                                    Name ="tbxUser"
                                    ControlSource ="UserName"
                                    FontName ="Arial"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =1035
                                    LayoutCachedTop =1755
                                    LayoutCachedWidth =2895
                                    LayoutCachedHeight =1995
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =315
                                            Top =1755
                                            Width =468
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            ForeColor =-2147483603
                                            Name ="lblUser"
                                            Caption ="User"
                                            FontName ="Arial"
                                            LayoutCachedLeft =315
                                            LayoutCachedTop =1755
                                            LayoutCachedWidth =783
                                            LayoutCachedHeight =2007
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    SpecialEffect =2
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =1035
                                    Top =2115
                                    Width =1860
                                    FontSize =8
                                    TabIndex =2
                                    BackColor =-2147483629
                                    ForeColor =-2147483607
                                    Name ="tbxPark"
                                    ControlSource ="Park"
                                    FontName ="Arial"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =1035
                                    LayoutCachedTop =2115
                                    LayoutCachedWidth =2895
                                    LayoutCachedHeight =2355
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =315
                                            Top =2115
                                            Width =444
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            ForeColor =-2147483603
                                            Name ="lblPark"
                                            Caption ="Park"
                                            FontName ="Arial"
                                            LayoutCachedLeft =315
                                            LayoutCachedTop =2115
                                            LayoutCachedWidth =759
                                            LayoutCachedHeight =2367
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    SpecialEffect =2
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =1035
                                    Top =2475
                                    Width =1860
                                    FontSize =8
                                    TabIndex =3
                                    BackColor =-2147483629
                                    ForeColor =-2147483607
                                    Name ="tbxDatum"
                                    ControlSource ="Datum"
                                    FontName ="Arial"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =1035
                                    LayoutCachedTop =2475
                                    LayoutCachedWidth =2895
                                    LayoutCachedHeight =2715
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =315
                                            Top =2475
                                            Width =600
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            ForeColor =-2147483603
                                            Name ="lblDatum"
                                            Caption ="Datum"
                                            FontName ="Arial"
                                            LayoutCachedLeft =315
                                            LayoutCachedTop =2475
                                            LayoutCachedWidth =915
                                            LayoutCachedHeight =2727
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    SpecialEffect =2
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =1035
                                    Top =2835
                                    Width =1860
                                    FontSize =8
                                    TabIndex =4
                                    BackColor =-2147483629
                                    ForeColor =-2147483607
                                    Name ="tbxDeclination"
                                    ControlSource ="Declination"
                                    FontName ="Arial"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =1035
                                    LayoutCachedTop =2835
                                    LayoutCachedWidth =2895
                                    LayoutCachedHeight =3075
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =315
                                            Top =2835
                                            Width =1020
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            ForeColor =-2147483603
                                            Name ="lblDeclination"
                                            Caption ="Declin."
                                            FontName ="Arial"
                                            LayoutCachedLeft =315
                                            LayoutCachedTop =2835
                                            LayoutCachedWidth =1335
                                            LayoutCachedHeight =3087
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    SpecialEffect =2
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =1035
                                    Top =3195
                                    Width =1860
                                    Height =270
                                    FontSize =8
                                    TabIndex =5
                                    BackColor =-2147483629
                                    ForeColor =-2147483607
                                    Name ="tbxGPS_model"
                                    ControlSource ="GPS_model"
                                    FontName ="Arial"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =1035
                                    LayoutCachedTop =3195
                                    LayoutCachedWidth =2895
                                    LayoutCachedHeight =3465
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =315
                                            Top =3195
                                            Width =435
                                            Height =270
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            ForeColor =-2147483603
                                            Name ="lblGPS_model"
                                            Caption ="GPS"
                                            FontName ="Arial"
                                            LayoutCachedLeft =315
                                            LayoutCachedTop =3195
                                            LayoutCachedWidth =750
                                            LayoutCachedHeight =3465
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =1800
                                    Top =3780
                                    Width =1215
                                    Height =255
                                    FontSize =9
                                    TabIndex =6
                                    ForeColor =-2147483607
                                    Name ="tbxAppMode"
                                    StatusBarText ="Current user access level for this application"
                                    FontName ="Arial"
                                    ControlTipText ="Current user access level for this application"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =1800
                                    LayoutCachedTop =3780
                                    LayoutCachedWidth =3015
                                    LayoutCachedHeight =4035
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =195
                                            Top =3780
                                            Width =1560
                                            Height =270
                                            FontSize =9
                                            FontWeight =700
                                            ForeColor =-2147483602
                                            Name ="lblAppMode"
                                            Caption ="Application mode:"
                                            FontName ="Arial"
                                            LayoutCachedLeft =195
                                            LayoutCachedTop =3780
                                            LayoutCachedWidth =1755
                                            LayoutCachedHeight =4050
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    SpecialEffect =2
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =1035
                                    Top =1035
                                    Width =1860
                                    FontSize =8
                                    TabIndex =7
                                    BackColor =-2147483629
                                    ForeColor =-2147483607
                                    Name ="tbxProject"
                                    ControlSource ="Project"
                                    FontName ="Arial"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =1035
                                    LayoutCachedTop =1035
                                    LayoutCachedWidth =2895
                                    LayoutCachedHeight =1275
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =315
                                            Top =1035
                                            Width =672
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            ForeColor =-2147483603
                                            Name ="lblProject"
                                            Caption ="Project"
                                            FontName ="Arial"
                                            LayoutCachedLeft =315
                                            LayoutCachedTop =1035
                                            LayoutCachedWidth =987
                                            LayoutCachedHeight =1287
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    Left =1980
                                    Top =540
                                    Width =855
                                    Height =405
                                    TabIndex =8
                                    ForeColor =4210752
                                    Name ="btnChangeDefaults"
                                    Caption ="Change"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Change application defaults"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =1980
                                    LayoutCachedTop =540
                                    LayoutCachedWidth =2835
                                    LayoutCachedHeight =945
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
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =75
                            Top =435
                            Width =3075
                            Height =3765
                            BorderColor =10921638
                            Name ="pgAbout"
                            Caption =" Db Info"
                            GridlineColor =10921638
                            LayoutCachedLeft =75
                            LayoutCachedTop =435
                            LayoutCachedWidth =3150
                            LayoutCachedHeight =4200
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =2880
                                    Left =135
                                    Top =495
                                    Width =2880
                                    Height =270
                                    FontSize =8
                                    FontWeight =700
                                    ForeColor =-2147483607
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                                    Name ="cbxVersion"
                                    ControlSource ="Release_ID"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tsys_App_Releases.ID, 'Version ' & [VersionNumber] & ' (' & [ReleaseDate]"
                                        " & ')' AS Version FROM tsys_App_Releases; "
                                    ColumnWidths ="0;2880"
                                    FontName ="Arial"
                                    ControlTipText ="Version number of this application"
                                    AllowValueListEdits =0
                                    InheritValueList =0

                                    LayoutCachedLeft =135
                                    LayoutCachedTop =495
                                    LayoutCachedWidth =3015
                                    LayoutCachedHeight =765
                                    ThemeFontIndex =-1
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ForeThemeColorIndex =-1
                                    ForeShade =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =255
                                    Top =915
                                    Width =2700
                                    Height =270
                                    FontSize =8
                                    TabIndex =1
                                    BackColor =11056034
                                    ForeColor =-2147483607
                                    Name ="tbxContact_Name"
                                    ControlSource ="AppContactName"
                                    FontName ="Arial"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =255
                                    LayoutCachedTop =915
                                    LayoutCachedWidth =2955
                                    LayoutCachedHeight =1185
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =255
                                    Top =1235
                                    Width =2700
                                    Height =270
                                    FontSize =8
                                    TabIndex =2
                                    BackColor =11056034
                                    ForeColor =-2147483607
                                    Name ="tbxContact_Org"
                                    ControlSource ="AppContactOrg"
                                    FontName ="Arial"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =255
                                    LayoutCachedTop =1235
                                    LayoutCachedWidth =2955
                                    LayoutCachedHeight =1505
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =255
                                    Top =1555
                                    Width =2700
                                    Height =270
                                    FontSize =8
                                    TabIndex =3
                                    BackColor =11056034
                                    ForeColor =-2147483607
                                    Name ="tbxContact_Phone"
                                    ControlSource ="AppContactPhone"
                                    FontName ="Arial"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =255
                                    LayoutCachedTop =1555
                                    LayoutCachedWidth =2955
                                    LayoutCachedHeight =1825
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    FontUnderline = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    FELineBreak = NotDefault
                                    IsHyperlink = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =255
                                    Top =1875
                                    Width =2700
                                    Height =285
                                    FontSize =8
                                    TabIndex =4
                                    BackColor =11056034
                                    ForeColor =-2147483607
                                    Name ="tbxContact_Email"
                                    ControlSource ="AppContactEmail"
                                    FontName ="Arial"
                                    AsianLineBreak =0

                                    LayoutCachedLeft =255
                                    LayoutCachedTop =1875
                                    LayoutCachedWidth =2955
                                    LayoutCachedHeight =2160
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    FontUnderline = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    SpecialEffect =2
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =135
                                    Top =2535
                                    Width =2880
                                    Height =480
                                    FontSize =8
                                    TabIndex =5
                                    BackColor =-2147483604
                                    ForeColor =16711680
                                    Name ="tbxWebURL"
                                    ControlSource ="WebURL"
                                    StatusBarText ="Web address for application downloads"
                                    OnDblClick ="[Event Procedure]"
                                    FontName ="Arial"

                                    LayoutCachedLeft =135
                                    LayoutCachedTop =2535
                                    LayoutCachedWidth =3015
                                    LayoutCachedHeight =3015
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =135
                                            Top =2235
                                            Width =2880
                                            Height =270
                                            FontSize =9
                                            FontWeight =700
                                            ForeColor =-2147483603
                                            Name ="lblWeb_address"
                                            Caption ="Web address for updates:"
                                            FontName ="Arial"
                                            LayoutCachedLeft =135
                                            LayoutCachedTop =2235
                                            LayoutCachedWidth =3015
                                            LayoutCachedHeight =2505
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    Left =540
                                    Top =3120
                                    Width =2070
                                    Height =405
                                    TabIndex =6
                                    ForeColor =4210752
                                    Name ="btnReleaseHistory"
                                    Caption ="View release history"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="View application release history"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =540
                                    LayoutCachedTop =3120
                                    LayoutCachedWidth =2610
                                    LayoutCachedHeight =3525
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
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    Left =540
                                    Top =3645
                                    Width =2070
                                    Height =405
                                    TabIndex =7
                                    ForeColor =4210752
                                    Name ="btnReportBug"
                                    Caption ="Report a bug"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Report an application bug"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =540
                                    LayoutCachedTop =3645
                                    LayoutCachedWidth =2610
                                    LayoutCachedHeight =4050
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
                            End
                        End
                        Begin Page
                            Enabled = NotDefault
                            OverlapFlags =247
                            Left =75
                            Top =435
                            Width =3075
                            Height =3765
                            BorderColor =10921638
                            Name ="pgSettings"
                            OnClick ="[Event Procedure]"
                            Caption =" Settings"
                            GridlineColor =10921638
                            LayoutCachedLeft =75
                            LayoutCachedTop =435
                            LayoutCachedWidth =3150
                            LayoutCachedHeight =4200
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin CheckBox
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    Left =165
                                    Top =945
                                    Name ="chkBackupOnStartup"
                                    ControlSource ="BackupPromptOnStartup"
                                    StatusBarText ="Whether or not the application prompts for backups upon startup"

                                    LayoutCachedLeft =165
                                    LayoutCachedTop =945
                                    LayoutCachedWidth =425
                                    LayoutCachedHeight =1185
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =395
                                            Top =915
                                            Width =2532
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            ForeColor =-2147483607
                                            Name ="lblBackupOnStartup"
                                            Caption ="Prompt for backup on startup"
                                            FontName ="Arial"
                                            LayoutCachedLeft =395
                                            LayoutCachedTop =915
                                            LayoutCachedWidth =2927
                                            LayoutCachedHeight =1167
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    Left =165
                                    Top =1305
                                    TabIndex =1
                                    Name ="chkBackupOnExit"
                                    ControlSource ="BackupPromptOnExit"
                                    StatusBarText ="Whether or not the application prompts for backups upon exiting"

                                    LayoutCachedLeft =165
                                    LayoutCachedTop =1305
                                    LayoutCachedWidth =425
                                    LayoutCachedHeight =1545
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =395
                                            Top =1275
                                            Width =2244
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            ForeColor =-2147483607
                                            Name ="lblBackupOnExit"
                                            Caption ="Prompt for backup on exit"
                                            FontName ="Arial"
                                            LayoutCachedLeft =395
                                            LayoutCachedTop =1275
                                            LayoutCachedWidth =2639
                                            LayoutCachedHeight =1527
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    Left =165
                                    Top =1667
                                    TabIndex =2
                                    Name ="chkCompactBEOnExit"
                                    ControlSource ="CompactBEOnExit"
                                    StatusBarText ="Whether or not the application compacts the back-end db upon exiting"

                                    LayoutCachedLeft =165
                                    LayoutCachedTop =1667
                                    LayoutCachedWidth =425
                                    LayoutCachedHeight =1907
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =393
                                            Top =1635
                                            Width =2376
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            ForeColor =-2147483607
                                            Name ="lblCompactBEOnExit"
                                            Caption ="Compact back-end on exit"
                                            FontName ="Arial"
                                            LayoutCachedLeft =393
                                            LayoutCachedTop =1635
                                            LayoutCachedWidth =2769
                                            LayoutCachedHeight =1887
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    Left =165
                                    Top =2027
                                    TabIndex =3
                                    Name ="chkVerifyOnStartup"
                                    ControlSource ="VerifyLinksOnStartup"
                                    StatusBarText ="Whether or not the application verifies table connections upon startup"

                                    LayoutCachedLeft =165
                                    LayoutCachedTop =2027
                                    LayoutCachedWidth =425
                                    LayoutCachedHeight =2267
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =393
                                            Top =1995
                                            Width =2700
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            ForeColor =-2147483607
                                            Name ="lblVerifyOnStartup"
                                            Caption ="Test all connections on startup"
                                            FontName ="Arial"
                                            LayoutCachedLeft =393
                                            LayoutCachedTop =1995
                                            LayoutCachedWidth =3093
                                            LayoutCachedHeight =2247
                                            ThemeFontIndex =-1
                                            BackThemeColorIndex =-1
                                            BorderThemeColorIndex =-1
                                            BorderTint =100.0
                                            ForeThemeColorIndex =-1
                                            ForeTint =100.0
                                            GridlineThemeColorIndex =-1
                                            GridlineShade =100.0
                                        End
                                    End
                                End
                                Begin CommandButton
                                    Enabled = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    Left =420
                                    Top =2700
                                    Width =2280
                                    Height =405
                                    TabIndex =4
                                    ForeColor =4210752
                                    Name ="btnChangeDbInfo"
                                    Caption ="Set application info"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Update database version and contact info"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =420
                                    LayoutCachedTop =2700
                                    LayoutCachedWidth =2700
                                    LayoutCachedHeight =3105
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
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    Left =420
                                    Top =3300
                                    Width =2280
                                    Height =405
                                    TabIndex =5
                                    ForeColor =4210752
                                    Name ="btnManageLinks"
                                    Caption ="Manage back-end links"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Manage back-end links"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =420
                                    LayoutCachedTop =3300
                                    LayoutCachedWidth =2700
                                    LayoutCachedHeight =3705
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
                            End
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5880
                    Top =5280
                    Width =1800
                    Height =405
                    TabIndex =19
                    ForeColor =4210752
                    Name ="btnSOPs"
                    Caption ="SOPs"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5880
                    LayoutCachedTop =5280
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =5685
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =3960
                    Top =5280
                    Width =1800
                    Height =405
                    TabIndex =20
                    ForeColor =4210752
                    Name ="btn3"
                    GridlineColor =10921638

                    LayoutCachedLeft =3960
                    LayoutCachedTop =5280
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =5685
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
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =2040
                    Top =5280
                    Width =1800
                    Height =405
                    TabIndex =21
                    ForeColor =4210752
                    Name ="btnImportCSV"
                    Caption ="Import CSV"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2040
                    LayoutCachedTop =5280
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =5685
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =120
                    Top =5280
                    Width =1800
                    Height =405
                    TabIndex =22
                    ForeColor =4210752
                    Name ="btnViewTemplates"
                    Caption ="View Templates"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open the list of SQL & other templates"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =5280
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =5685
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
' Form:         DbAdmin
' Level:        Framework form
' Version:      1.10
' Basis:        -
'
' Description:  DbAdmin form object related properties, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
'               BLC - 6/12/2016 - 1.01 - adapted to framework & big rivers
'               BLC - 8/31/2016 - 1.02 - added user access level controls
'               BLC - 9/14/2016 - 1.03 - added btnViewTemplates & additional blanks
'                                        for Management Tools
'               BLC - 10/19/2016 - 1.04 - added Import CSV button, added callingform property
'               BLC - 1/12/2017 - 1.05 - added Version button
'               BLC - 1/19/2017 - 1.06 - revised to SOPs from Version (button)
'               BLC - 1/26/2017 - 1.07 - hid unused defaults, temporarily redirected button
'                                        clicks to messages
'               BLC - 9/6/2017  - 1.08 - added tglDevMode button, ToggleDevMode()
'               BLC - 10/17/2017 - 1.09 - commented out initApp (already called in PreSplash form)
'               BLC - 10/19/2017 - 1.10 - added comment length
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_ButtonCaption As String
Private m_CallingForm As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidLabel(Value As String)
Public Event InvalidCaption(Value As String)
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

Public Property Let ButtonCaption(Value As String)
    If Len(Value) > 0 Then
        m_ButtonCaption = Value

        'set the form button caption
        'Me.btnSave.Caption = m_ButtonCaption
    Else
        RaiseEvent InvalidCaption(Value)
    End If
End Property

Public Property Get ButtonCaption() As String
    ButtonCaption = m_ButtonCaption
End Property

Public Property Let CallingForm(Value As String)
    If Len(Value) > 0 Then
        m_CallingForm = Value
    Else
        RaiseEvent InvalidCallingForm(Value)
    End If
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
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 10/19/2016 - adjust to use calling form property
'   BLC - 1/12/2017 - add versions button
'   BLC - 1/26/2017 - hid unused defaults
'   BLC - 10/17/2017 - commented out initApp (already called in PreSplash form)
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "Main"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize Main
    ToggleForm Me.CallingForm, -1
    
    Title = "Db Admin"
    Directions = "Choose the desired action below."
    lblDirections.ForeColor = lngLtBlue
    btnComment.Caption = StringFromCodepoint(uComment)
    btnComment.ForeColor = lngBlue
    
    'set mode display
    tbxAppMode.ForeColor = lngGreen
    tbxAppMode.BorderStyle = 0  '0-transparent, 1-normal
    tbxAppMode.TextAlign = 2    '0-general, 1-left, 2-center, 3-right
    
    'set hovers
    btnComment.HoverColor = lngGreen
    btnBackup.HoverColor = lngGreen
    btnBrowser.HoverColor = lngGreen
    btnChangeDbInfo.HoverColor = lngGreen
    btnChangeDefaults.HoverColor = lngGreen
    btnDbWindow.HoverColor = lngGreen
    btnEditLog.HoverColor = lngGreen
    btnEnter.HoverColor = lngGreen
    btnLookups.HoverColor = lngGreen
    btnManageLinks.HoverColor = lngGreen
    btnNavCoords.HoverColor = lngGreen
    btnNavReport.HoverColor = lngGreen
    btnQA.HoverColor = lngGreen
    btnQAReport.HoverColor = lngGreen
    btnReconnect.HoverColor = lngGreen
    btnReleaseHistory.HoverColor = lngGreen
    btnReportBug.HoverColor = lngGreen
    btnSetRoles.HoverColor = lngGreen
    btnSpeciesListRpt.HoverColor = lngGreen
    btnSummaries.HoverColor = lngGreen
    btnTaskList.HoverColor = lngGreen
    btnTaskListRpt.HoverColor = lngGreen
    btnUISetup.HoverColor = lngGreen
    btnViewTemplates.HoverColor = lngGreen
    btnImportCSV.HoverColor = lngGreen
    btnSOPs.HoverColor = lngGreen
      
    'defaults
    Me.RecordSource = GetTemplate("s_db_admin_info") '"tsys_App_Defaults"
    cbxVersion.RowSource = GetTemplate("s_app_releases")
    cbxVersion.ControlSource = "Release_ID" '"ID"
    
    tbxAppMode.Value = TempVars("UserAccessLevel")
    
    'hide unused defaults
    lblDatum.Visible = False
    tbxDatum.Visible = False
    lblDeclination.Visible = False
    tbxDeclination.Visible = False
    lblGPS_model.Visible = False
    tbxGPS_model.Visible = False
    
    ' Update the DbAdmin switchboard settings according to application mode
    setUserAccess Me
    
    'initialize app settings
'    initApp    '<< DUPE call (already called in PreSplash form)
    
    ' If there is an Access back-end, open the always-open form (to maintain a connection
    '   to the back-end and avoid unnecessary create/delete/updates to its .ldb lock file)
    If Nz(TempVars("HasAccessBE"), False) Then DoCmd.OpenForm "LockBE", , , , , acHidden

    ' If there is an Access back-end, make the backups button visible
    Me!btnBackup.Visible = Nz(TempVars("HasAccessBE"), False)
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[DbAdmin form])"
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


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Activate
' Description:  form actions when open form gets focus
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Adapted:      -
' Revisions:
'   BLC - 7/28/2017 - initial version
'   BLC - 7/31/2017 - revised to ensure Dev Mode toggle button updates w/ current state
' ---------------------------------
Private Sub Form_Activate()

    'set toggle based on current value
    Me.tglDevMode.Value = DEV_MODE
    
    ToggleDevMode

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Activate[DbAmin form])"
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
            "Error encountered (#" & Err.Number & " - Form_Current[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglDevMode_Click
' Description:  Sets value for DEV_MODE true or false based on toggle
'                   Up = True, Down = False
' Assumptions:  DEV_MODE sets visibility of ID controls
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Adapted:      -
' Revisions:
'   BLC - 7/28/2017 - initial version
'   BLC - 7/31/2017 - revised to shift code to ToggleDevMode
' ---------------------------------
Private Sub tglDevMode_Click()
On Error GoTo Err_Handler

    ToggleDevMode
 
Exit_Handler:
    DoCmd.Hourglass False
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglDevMode_Click[DbAdmin form])"
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
'   BLC - 10/19/2017 - added comment length
' ---------------------------------
Private Sub btnComment_Click()
On Error GoTo Err_Handler
    
    'open comment form
    DoCmd.OpenForm "Comment", acNormal, , , , , "DbAdmin|" & "" & "|255"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[DbAdmin form])"
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
'   BLC - 6/24/2016 - added check for open form, restore main form
'   BLC - 10/19/2016 - adjusted to use callingform property
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore calling form
    ToggleForm Me.CallingForm, 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   DEV_MODE Functionality
' =================================
' ---------------------------------
' Sub:          ToggleDevMode
' Description:  Sets value for DEV_MODE true or false based on toggle
'                   Up = True, Down = False
' Assumptions:  DEV_MODE sets visibility of ID controls
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Adapted:      -
' Revisions:
'   BLC - 7/31/2017 - initial version
' ---------------------------------
Private Sub ToggleDevMode()
On Error GoTo Err_Handler

    'set global based on toggle
    If Me!tglDevMode = True Then
        'true = up
        DEV_MODE = True
        
        With Me.tglDevMode
            .Caption = "DEV MODE ON"
            .BackColor = lngLtLime
            .fontBold = True
            .ForeColor = lngBlue
        End With
    Else
        'false = down
        DEV_MODE = False
    
        With Me.tglDevMode
            .Caption = "DEV MODE OFF"
            .BackColor = lngLtrYellow
            .fontBold = False
            .ForeColor = lngRed
        End With
    End If
 
Exit_Handler:
    DoCmd.Hourglass False
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleDevMode[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   DbAdmin Tab Functionality
' =================================
' =================================
' TAB (PAGE) NAME:    DbAdmin (tabDbAdmin)
' Description:  Db Administrative functions
' Unbound ctls:
' Subforms:     fsub_DbAdmin
' =================================
' ---------------------------------
' SUB:          btnReconnect_Click
' Description:  Opens form for reconnecting backend dbs
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 6/13/2014 - initial version
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub btnReconnect_Click()
    On Error GoTo Err_Handler

    ' Reconnect back end tables
    DoCmd.OpenForm "ConnectDbs"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReconnect_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnSetRoles_Click
' Description:  Open form for setting user roles
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 7/31/2014 - changed gvars to TempVars
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub btnSetRoles_Click()
    On Error GoTo Err_Handler

    ' Open the form to set user roles for this project (if in admin / power user mode)
    If TempVars("Connected") = False Then
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    Else
        Select Case TempVars("UserAccessLevel")
          Case "admin", "power user"
            DoCmd.OpenForm "Contact"
        End Select
    End If

'REMOVE / FIX:
    DoCmd.OpenForm "Contact"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSetRoles_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDbWindow_Click
' Description:  opens database window
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 6/13/2014 - initial version
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub btnDbWindow_Click()
    On Error GoTo Err_Handler

    ' Show the database window.  To re-hide: DoCmd.RunCommand acCmdWindowHide
    DoCmd.SelectObject acForm, "", True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDbWindow_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:     btnBackup_Click
' Description:  Backup the current backend databases
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 7/31/2014 - Changed gvarConnected, gvarHasAccessBE, gvarWritePermission to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub btnBackup_Click()
    On Error GoTo Err_Handler

    ' Make sure that the database is connected
    If TempVars("Connected") = False Then
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
        GoTo Exit_Handler
    ElseIf TempVars("HasAccessBE") = False Then
        MsgBox "There are no Access back-ends currently connected ...", _
            vbExclamation, "No backup made"
        GoTo Exit_Handler
    ElseIf TempVars("WritePermission") = False Then
        MsgBox "The back-end database is in a read-only state ...", _
            vbExclamation, "No backup made"
        GoTo Exit_Handler
    Else
        Call MakeBackup
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnBackup_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   DbAdmin: Database admin functions
' =================================

' ---------------------------------
' PAGE NAME:    Application Defaults (pgDefaults)
' Description:  system defaults for the run-time environment
' Bound ctls:   various fields for displaying default values
' Unbound ctls: btnChangeDefaults - opens a popup for changing default values
' Subforms:     none
' ---------------------------------

' ---------------------------------
' SUB:          btnChangeDefaults_Click
' Description:  Open form for setting database defaults
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 7/31/2014 - Changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub btnChangeDefaults_Click()
    On Error GoTo Err_Handler

    ' Perform data validation
    If TempVars("Connected") Then
        DoCmd.OpenForm "SetDefaults", , , , , , 4
    Else
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnChangeDefaults_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' PAGE NAME:    Database Information (pgAbout)
' Description:  database development and release information
' Unbound ctls: btnReleaseHistory, btnReportBug
' Subforms:     none
' ---------------------------------

' ---------------------------------
' SUB:          tbxWebURL_DblClick
' Description:  Opens website address
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 6/13/2014 - initial version
'               BLC - 8/25/2014 - renamed control tbx vs txt
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub tbxWebURL_DblClick(Cancel As Integer)
On Error GoTo Err_Handler

    ' Upon clicking the project web address, open the website
    DoCmd.Hourglass True
    Application.FollowHyperlink Me.tbxWebURL, , True, False

Exit_Handler:
    DoCmd.Hourglass False
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxWebURL_DblClick[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnReleaseHistory_Click
' Description:  Opens database release history
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/28/2014 - updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 7/31/2014 - changed gvars to TempVars
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub btnReleaseHistory_Click()
On Error GoTo Err_Handler

    ' View the release history form
    If Nz(TempVars("Connected"), False) Then
        If TempVars("UserAccessLevel") = "admin" Then
            DoCmd.OpenForm "AppReleases"
        Else    ' read-only for all but admin users
            DoCmd.OpenForm "AppReleases", , , , acFormReadOnly
        End If
    Else
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReleaseHistory_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnReportBug_Click
' Description:  Opens bug reporting form
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub btnReportBug_Click()
On Error GoTo Err_Handler

    Dim strMsg As String

    ' Report an application bug
    strMsg = vbCrLf & "and provide the following details:" & vbCrLf & vbCrLf & _
        vbTab & "- Your name and the date the problem occurred." & vbCrLf & vbCrLf & _
        vbTab & "- Application version number." & vbCrLf & vbCrLf & _
        vbTab & "- Name of the form or report, if known." & vbCrLf & vbCrLf & _
        vbTab & "- Frequency of the problem (every time or just once?)." & vbCrLf & vbCrLf & _
        vbTab & "- Other details to help isolate the problem (e.g. 'Error #2001" _
        & vbCrLf & vbTab & "  message appears when entering the site code for a new record')."

    ' Change the instructions statement depending on whether or not the form will be opened
    If Nz(TempVars("Connected"), False) Then
        strMsg = "Please fill out the following form to describe the problem" & strMsg
    Else
        strMsg = "The database is not properly connected." & vbCrLf & vbCrLf & _
            "Please email the developer about the problem you're experiencing" & strMsg & _
            vbCrLf & vbCrLf & vbTab & _
            "- Send a screen capture of the problem if possible - hit" & _
            vbCrLf & vbTab & "  Alt-Print Scrn and paste the image into the email or save" _
            & vbCrLf & vbTab & "  it as a JPEG using a program such as Paint."
    End If
    MsgBox strMsg, , "Report a database problem or error"

    ' If connected, open the form along with instructions
    If Nz(TempVars("Connected"), False) Then DoCmd.OpenForm "BugReports", , , , acFormAdd, acDialog, 1

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReportBug_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' PAGE NAME:    Application settings (pgSettings)
' Description:  administrative application settings
' Unbound ctls: btnChangeDbInfo
' Subforms:     none
' ---------------------------------

' ---------------------------------
' SUB:          btnChangeDbInfo_Click
' Description:  Opens form for changing database info
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 6/19/2014 - Replaced Me.cAppMode with TempVars.Item("UserAccessLevel")
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub btnChangeDbInfo_Click()
On Error GoTo Err_Handler

    ' Update database version and contact info
    If TempVars("UserAccessLevel") = "admin" Then
        DoCmd.OpenForm "SetDbInfo"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnChangeDbInfo_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnManageLinks_Click
' Description:  Opens form for managing linked databases
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 6/19/2014 - Replaced Me.cAppMode with TempVars.Item("UserAccessLevel")
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub btnManageLinks_Click()
On Error GoTo Err_Handler

    ' Update database version and contact info
    If TempVars("UserAccessLevel") = "admin" Then
        DoCmd.OpenForm "ConnectDbs"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnManageLinks_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   DbAdmin: Data entry and edit functions
' =================================

' ---------------------------------
' SUB:          btnEnter_Click
' Description:  Reconnect to backend dbs
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub btnEnter_Click()
On Error GoTo Err_Handler
    
    ' Open the main data entry forms
    If TempVars("Connected") Then
        ' Prompt to make a backup, depending on application settings
        '   Note:  only relevant for Access back-end files
        If Me.chkBackupOnStartup And TempVars("HasAccessBE") Then MakeBackup
        DoCmd.OpenForm "SetDefaults", , , , , , 1
    Else
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEnter_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnTaskList_Click
' Description:  Open sample locations task list
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
'               BLC - 1/26/2017 - temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnTaskList_Click()
On Error GoTo Err_Handler

    ' View the list of tasks associated with sample locations
    If TempVars("Connected") Then
'        DoCmd.OpenForm "TaskList"
        DisplayMsg "undev"
    Else
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnTaskList_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'  DbAdmin: Management functions
' =================================

' ---------------------------------
' SUB:          btnBrowser_Click
' Description:  Open form for data browsing
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
'               BLC - 1/26/2017 - temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnBrowser_Click()
On Error GoTo Err_Handler

    ' Open the data browser
    If TempVars("Connected") Then
        'DoCmd.OpenForm "SetDefaults", , , , , , 2
        DisplayMsg "undev"
    Else
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnBrowser_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnLookups_Click
' Description:  Open form for viewing lookup tables
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
'               BLC - 1/26/2017 - temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnLookups_Click()
On Error GoTo Err_Handler

    ' Review and edit lookup tables
    If TempVars("Connected") = False Then
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    Else
'        DoCmd.OpenForm "Lookups"
        DisplayMsg "undev"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLookups_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnQA_Click
' Description:  Open the QA form tool
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
'               BLC - 1/26/2017 - temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnQA_Click()
On Error GoTo Err_Handler

    ' Open the data validation tool
    If TempVars("Connected") Then
'        DoCmd.OpenForm "QATool"
         DisplayMsg "dev"
    Else
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnQA_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnEditLog_Click
' Description:  Open edit log form
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
'               BLC - 1/26/2017 - temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnEditLog_Click()
On Error GoTo Err_Handler

    ' Open the edit log form
    If TempVars("Connected") Then
        'DoCmd.OpenForm "EditLog"
        DisplayMsg "undev"
    Else
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEditLog_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnViewTemplates_Click
' Description:  Open form for browsing templates
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, September 14 2014 for NCPN tools
' Adapted:      -
' Revisions:    BLC - 9/14/2016 - initial version
' ---------------------------------
Private Sub btnViewTemplates_Click()
On Error GoTo Err_Handler

    ' Open the template list
    If TempVars("Connected") Then
        DoCmd.OpenForm "TemplateList", , , , , , 2
    Else
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnViewTemplates_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnImportCSV_Click
' Description:  Import CSV button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 19, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/19/2016 - initial version
' ---------------------------------
Private Sub btnImportCSV_Click()
On Error GoTo Err_Handler
    
    'open import map form
    DoCmd.OpenForm "ImportMap", acNormal, , , , , "DbAdmin"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnImportCSV_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSOPs_Click
' Description:  SOPs button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 12, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/12/2017 - initial version
'   BLC - 1/19/2017 - revised to SOPs vs. Versions
' ---------------------------------
Private Sub btnSOPs_Click()
On Error GoTo Err_Handler
    
    'open import map form
    DoCmd.OpenForm "SOPVersion", acNormal, , , , , "DbAdmin"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSOPs_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'  DbAdmin: Summaries and db output
' =================================

' ---------------------------------
' SUB:          btnSummaries_Click
' Description:  Open the summaries tool
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
'               BLC - 1/26/2017 - temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnSummaries_Click()
On Error GoTo Err_Handler

    ' Open the data summary tool
    If TempVars("Connected") Then
'        DoCmd.OpenForm "SummaryTool"
        DisplayMsg "undev"
    Else
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSummaries_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnSpeciesListRpt_Click
' Description:  Open the species list tool
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell - January 26, 2017 for NCPN tools
' Adapted:      -
' Revisions:    BLC - 1/26/2017 - initial version & temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnSpeciesListRpt_Click()
On Error GoTo Err_Handler

    ' Open the species list tool
    If TempVars("Connected") Then
'        DoCmd.OpenForm "SummaryTool"
        DisplayMsg "undev"
    Else
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSpeciesListRpt_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnTaskListRpt_Click
' Description:  Open & generate task list report
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
'               BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
'               BLC - 1/26/2017 - temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnTaskListRpt_Click()
On Error GoTo Err_Handler

    ' Notify if not connected
    If TempVars("Connected") = False Then
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
        GoTo Exit_Handler
    End If

    DisplayMsg "undev"
    GoTo Exit_Handler
    
    ' Generate the task list report
    Dim strRptName As String
    Dim strMsg As String
    Dim strFilter As String
    Dim strTimeframe As String
    Dim bFilterOn As Boolean
    Dim strCaption As String
    Dim strPark As String
    Dim strSite As String
    Dim strInitFile As String
    Dim strSaveFile As String
    Dim varResponse As VbMsgBoxResult

    strRptName = "rpt_Task_List"

    strFilter = ""
    bFilterOn = False
    strTimeframe = TempVars("Timeframe")

    strMsg = "This will generate the task list report ..." & vbCrLf & vbCrLf & _
        "Would you like to limit task list output to " & vbCrLf & _
        "scheduled sampling locations for " & strTimeframe & "?" & vbCrLf & vbCrLf & _
        "Select NO to output all active task items ..."
    varResponse = MsgBox(strMsg, vbYesNoCancel, "Task list report")

    Select Case varResponse
      Case vbCancel
        GoTo Exit_Handler
      Case vbYes
        bFilterOn = True
        strFilter = "[Calendar_year]=""" & strTimeframe & """"
        strCaption = strTimeframe
      Case Else
        ' Do not filter by calendar year
        strCaption = ""
    End Select
    
    ' Get user input for the park and/or location to filter on
    strPark = Trim(InputBox("Enter the park code to filter by" & vbCrLf & _
        "(or leave blank to show all):", "Filter by park"))
    strSite = Trim(InputBox("Enter the site code" & vbCrLf & _
        "(or leave blank to show all):", "Filter by site code"))
    ' Create the filter string
    If strPark <> "" Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Park_code]=""" & strPark & """"
    End If
    If strSite <> "" Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Site_code]=""" & strSite & """"
    End If

    DoCmd.OpenReport strRptName, acViewPreview, , strFilter, , strCaption
    If MsgBox("Would you like to save this report?", vbYesNo + vbDefaultButton2, _
        "Save report to a file?") = vbYes Then
        If varResponse = vbYes And strTimeframe <> "" Then
            ' Add timeframe to file name
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                strTimeframe & "_" & CStr(Format(Now(), "yyyymmdd")) & ".snp"
        Else
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                CStr(Format(Now(), "yyyymmdd")) & ".snp"
        End If
        ' Open the save file dialog and update to the actual name given by the user
        strSaveFile = SaveFile(strInitFile, "Snapshot Viewer (*.snp)", "*.snp")
        DoCmd.OutputTo acOutputReport, strRptName, acFormatSNP, strSaveFile, True
        MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnTaskListRpt_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnVegWalk_Click
' Description:  Open & generate species list report
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
'               BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
' ---------------------------------
Private Sub btnVegWalkRpt_Click()
On Error GoTo Err_Handler

    ' Notify if not connected
    If TempVars("Connected") = False Then
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
        GoTo Exit_Handler
    End If

    ' Generate the transect species list report
    Dim strRptName As String
    Dim strMsg As String
    Dim strFilter As String
    Dim strTimeframe As String
    Dim bFilterOn As Boolean
    Dim strCaption As String
    Dim strPark As String
    Dim strSite As String
    Dim strInitFile As String
    Dim strSaveFile As String
    Dim varResponse As VbMsgBoxResult

    strRptName = "rpt_Transect_species_list"

    strFilter = ""
    bFilterOn = False
    strTimeframe = TempVars("Timeframe")

    strMsg = "This will generate the transect species report ..." & vbCrLf & vbCrLf & _
        "Would you like to show only scheduled sites for " & strTimeframe & "?" & vbCrLf & _
        vbCrLf & "Select NO to output species lists for all sites ..."
    varResponse = MsgBox(strMsg, vbYesNoCancel, "Transect species list report")

    Select Case varResponse
      Case vbCancel
        GoTo Exit_Handler
      Case vbYes
        bFilterOn = True
        strFilter = "[Calendar_year]=""" & strTimeframe & """"
        strCaption = strTimeframe
      Case Else
        strCaption = ""
    End Select

    ' Get user input for the park and/or location to filter on
    strPark = Trim(InputBox("Enter the park code to filter by" & vbCrLf & _
        "(or leave blank to show all):", "Filter by park"))
    strSite = Trim(InputBox("Enter the site code" & vbCrLf & _
        "(or leave blank to show all):", "Filter by site code"))
    ' Create the filter string
    If strPark <> "" Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Park_code]=""" & strPark & """"
    End If
    If strSite <> "" Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Site_code]=""" & strSite & """"
    End If

    DoCmd.OpenReport strRptName, acViewPreview, , strFilter, , strCaption
    If MsgBox("Would you like to save this report?", vbYesNo + vbDefaultButton2, _
        "Save report to a file?") = vbYes Then
        If strTimeframe <> "" Then
            ' Add timeframe to file name
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                strTimeframe & "_" & CStr(Format(Now(), "yyyymmdd")) & ".snp"
        Else
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                CStr(Format(Now(), "yyyymmdd")) & ".snp"
        End If
        ' Open the save file dialog and update to the actual name given by the user
        strSaveFile = SaveFile(strInitFile, "Snapshot Viewer (*.snp)", "*.snp")
        DoCmd.OutputTo acOutputReport, strRptName, acFormatSNP, strSaveFile, True
        MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnVegWalkRpt_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnNavReport_Click
' Description:  Open & generate transect navigation reports
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
'               BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
'               BLC - 1/26/2017 - temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnNavReport_Click()
On Error GoTo Err_Handler

    ' Notify if not connected
    If TempVars("Connected") = False Then
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
        GoTo Exit_Handler
    End If

    DisplayMsg "undev"
    GoTo Exit_Handler

    ' Generate the transect navigation report
    Dim strRptName As String
    Dim strMsg As String
    Dim strFilter As String
    Dim strTimeframe As String
    Dim bFilterOn As Boolean
    Dim strCaption As String
    Dim strPark As String
    Dim strSite As String
    Dim strInitFile As String
    Dim strSaveFile As String
    Dim varResponse As VbMsgBoxResult

    strRptName = "rpt_Navigation_report"

    strFilter = ""
    bFilterOn = False
    strTimeframe = TempVars("Timeframe")

    strMsg = "This will generate the navigation report ..." & vbCrLf & vbCrLf & _
        "Output will be limited to scheduled sampling locations for " & strTimeframe & "."
    varResponse = MsgBox(strMsg, vbOKCancel, "Navigation report")

    Select Case varResponse
      Case vbCancel
        GoTo Exit_Handler
      Case Else
        strCaption = strTimeframe
    End Select
    
    ' Get user input for the park and/or location to filter on
    strPark = Trim(InputBox("Enter the park code to filter by" & vbCrLf & _
        "(or leave blank to show all):", "Filter by park"))
    strSite = Trim(InputBox("Enter the site code" & vbCrLf & _
        "(or leave blank to show all):", "Filter by site code"))
    ' Create the filter string
    If strPark <> "" Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Park_code]=""" & strPark & """"
    End If
    If strSite <> "" Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Site_code]=""" & strSite & """"
    End If

    DoCmd.OpenReport strRptName, acViewPreview, , strFilter, , strCaption
    If MsgBox("Would you like to save this report?", vbYesNo + vbDefaultButton2, _
        "Save report to a file?") = vbYes Then
        If strTimeframe <> "" Then
            ' Add timeframe to file name
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                strTimeframe & "_" & CStr(Format(Now(), "yyyymmdd")) & ".snp"
        Else
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                CStr(Format(Now(), "yyyymmdd")) & ".snp"
        End If
        ' Open the save file dialog and update to the actual name given by the user
        strSaveFile = SaveFile(strInitFile, "Snapshot Viewer (*.snp)", "*.snp")
        DoCmd.OutputTo acOutputReport, strRptName, acFormatSNP, strSaveFile, True
        MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNavReport_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnQAReport_Click
' Description:  Open & generate QA reports
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
'               BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
'               BLC - 1/26/2017 - temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnQAReport_Click()
On Error GoTo Err_Handler

    ' Notify if not connected
    If TempVars("Connected") = False Then
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
        GoTo Exit_Handler
    End If

    DisplayMsg "undev"
    GoTo Exit_Handler

    ' Generate the QA report
    Dim strRptName As String
    Dim strMsg As String
    Dim strFilter As String
    Dim strTimeframe As String
    Dim strInitFile As String
    Dim strSaveFile As String
    Dim varResponse As VbMsgBoxResult

    strRptName = "rpt_QA_Results"

    strMsg = "This will open the quality assurance report ..." & vbCrLf & vbCrLf & _
        "Would you like to limit report results to " & TempVars("Timeframe") & "?"
    varResponse = MsgBox(strMsg, vbYesNoCancel, "Quality assurance report")

    Select Case varResponse
      Case vbCancel
        GoTo Exit_Handler
      Case vbYes
        strTimeframe = TempVars("Timeframe")
        strFilter = "[Time_frame]=""" & strTimeframe & """"
      Case Else
        strTimeframe = Trim(InputBox("Enter the time frame to filter by" & vbCrLf & _
            "(or leave blank to show all):", "Filter by data time frame", _
            TempVars("Timeframe")))
        If strTimeframe <> "" Then
            strFilter = "[Time_frame]=""" & strTimeframe & """"
        Else
            strFilter = ""
        End If
    End Select

    DoCmd.OpenReport strRptName, acViewPreview, , strFilter
    If MsgBox("Would you like to save this report?", vbYesNo + vbDefaultButton2, _
        "Save report to a file?") = vbYes Then
        If strTimeframe <> "" Then
            ' Add timeframe to file name
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                strTimeframe & "_" & CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".snp"
        Else
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".snp"
        End If
        ' Open the save file dialog and update to the actual name given by the user
        strSaveFile = SaveFile(strInitFile, "Snapshot Viewer (*.snp)", "*.snp")
        DoCmd.OutputTo acOutputReport, strRptName, acFormatSNP, strSaveFile, True
        MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnQAReport_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnNavCoords_Click
' Description:  Open & generate navigation coordinates listing
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
'               BLC - 7/31/2014 - changed gvars to TempVars
'               BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC - 6/12/2016 - adapted for big rivers
'               BLC - 1/26/2017 - temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnNavCoords_Click()
On Error GoTo Err_Handler

    ' Notify if not connected
    If TempVars("Connected") = False Then
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
        GoTo Exit_Handler
    End If

    DisplayMsg "undev"
    GoTo Exit_Handler

    ' Export navigation target coordinates for pre-season upload to GPS units
    Dim strQryName As String
    Dim strMsg As String
    Dim strInitFile As String
    Dim strSaveFile As String
    Dim varResponse As VbMsgBoxResult

    strMsg = "This will generate the navigation target coordinates ..." & vbCrLf & vbCrLf & _
        "Would you like to show only scheduled sites for " & TempVars("Timeframe") & "?" & vbCrLf & _
        vbCrLf & "Select NO to output navigation coordinates for all sites ..."
    varResponse = MsgBox(strMsg, vbYesNoCancel, "Navigation target coordinates")

    Select Case varResponse
      Case vbCancel
        GoTo Exit_Handler
      Case vbYes
        ' Verify that there are scheduled locations for the current year
        If DCount("*", "tbl_Schedule", "[Calendar_year]=""" & TempVars("Timeframe") & """") = 0 Then
            If MsgBox("There are no scheduled sites for " & TempVars("Timeframe") & "." & vbCrLf & _
                "Show all sites instead?", vbYesNo + vbDefaultButton2, _
                "No scheduled sites") = vbYes Then
                strQryName = "qrpt_Navigation_target_coordinates_all"
            Else
                GoTo Exit_Handler
            End If
        Else
            strQryName = "qrpt_Navigation_target_coordinates"
        End If
      Case vbNo
        strQryName = "qrpt_Navigation_target_coordinates_all"
      Case Else
        GoTo Exit_Handler
    End Select

    DoCmd.OpenQuery strQryName, acViewNormal, acReadOnly
    If MsgBox("Would you like to save this file?", vbYesNo, _
        "Save the export file?") = vbYes Then
        strInitFile = Application.CurrentProject.Path & "\" & _
            strQryName & "_" & CStr(Format(Now(), "yyyymmdd")) & ".xls"
        ' Open the save file dialog and update to the actual name given by the user
        strSaveFile = SaveFile(strInitFile, "Microsoft Excel (*.xls)", "*.xls")
        DoCmd.OutputTo acOutputQuery, strQryName, acFormatXLS, strSaveFile, True
        MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNavCoords_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'  DbAdmin: User Interface
' =================================

' ---------------------------------
' SUB:          btnUISetup_Click
' Description:  Open the UI setup tool
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db by DbAdmin control set
' Adapted:      Bonnie Campbell, June 30, 2016 for NCPN big rivers tool
' Revisions:    BLC - 6/30/2016 - initial version
'               BLC - 1/26/2017 - temporariliy redirected to undeveloped message
' ---------------------------------
Private Sub btnUISetup_Click()
On Error GoTo Err_Handler

    ' Open the UI setup tool
    If TempVars("Connected") Then
        'DoCmd.OpenForm "SummaryTool"
        DisplayMsg "undev"
    Else
        MsgBox "The back-end connections must be fixed first", vbOKOnly, _
            "Not connected to back-end database"
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUISetup_Click[DbAdmin form])"
    End Select
    Resume Exit_Handler
End Sub
