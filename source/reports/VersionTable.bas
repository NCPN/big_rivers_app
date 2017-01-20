Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =1
    GridX =24
    GridY =24
    Width =15480
    DatasheetFontHeight =11
    ItemSuffix =106
    Right =11955
    Bottom =9240
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0xe436f31936e0e440
    End
    RecordSource ="SELECT SOP_VersionTable.* FROM SOP_VersionTable;"
    Caption ="SOP Version Table"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006d01000000000000783c00009402000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
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
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =16777215
            GridlineColor =16777215
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BoundObjectFrame
            AddColon = NotDefault
            SizeMode =3
            BorderLineStyle =0
            LabelX =-1800
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
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin UnboundObjectFrame
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="=[SOPName]"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin PageHeader
            Height =480
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    Left =180
                    Top =60
                    Width =1740
                    Height =324
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =6447974
                    Name ="lblTitle"
                    Caption ="SOP Version Table"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =384
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            Name ="GroupHeader0"
            AutoHeight =255
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin Section
            KeepTogether = NotDefault
            Height =420
            Name ="Detail"
            AlternateBackColor =12632256
            Begin
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2340
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    ForeColor =4210752
                    Name ="tbxVersion2"
                    ControlSource ="2-Training observers"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =2340
                    LayoutCachedTop =60
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =1
                    ForeColor =4210752
                    Name ="tbxVersion1"
                    ControlSource ="1-Prior to field season/equip_ lists"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =840
                    Width =1019
                    Height =360
                    TabIndex =2
                    ForeColor =4210752
                    Name ="tbxStartYear"
                    ControlSource ="EffectiveDate"
                    GridlineColor =10921638

                    LayoutCachedLeft =840
                    LayoutCachedWidth =1859
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Width =780
                    Height =360
                    TabIndex =3
                    ForeColor =4210752
                    Name ="tbxVK"
                    ControlSource ="=\"VK\" & [tbxRowNum]"
                    GridlineColor =10921638

                    LayoutCachedWidth =780
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    HideDuplicates = NotDefault
                    RunningSum =1
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Top =60
                    Width =480
                    Height =360
                    TabIndex =4
                    ForeColor =4210752
                    Name ="tbxRowNum"
                    ControlSource ="=1"
                    GridlineColor =10921638

                    LayoutCachedTop =60
                    LayoutCachedWidth =480
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3180
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =5
                    ForeColor =4210752
                    Name ="tbxVersion4"
                    ControlSource ="4-Sentinel site set up"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =3180
                    LayoutCachedTop =60
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2760
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =6
                    ForeColor =4210752
                    Name ="tbxVersion3"
                    ControlSource ="3-GPS methods"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =2760
                    LayoutCachedTop =60
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4020
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =7
                    ForeColor =4210752
                    Name ="tbxVersion6"
                    ControlSource ="6-Measuring vegetation"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =4020
                    LayoutCachedTop =60
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3600
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =8
                    ForeColor =4210752
                    Name ="tbxVersion5"
                    ControlSource ="5-Photos"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =3600
                    LayoutCachedTop =60
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4860
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =9
                    ForeColor =4210752
                    Name ="tbxVersion7"
                    ControlSource ="7-Hydrologic measurements"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =4860
                    LayoutCachedTop =60
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4440
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =10
                    ForeColor =4210752
                    Name ="tbxVersion66"
                    ControlSource ="6-Vegetation"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =4440
                    LayoutCachedTop =60
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5700
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =11
                    ForeColor =4210752
                    Name ="tbxVersion88"
                    ControlSource ="8-Surveying"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =5700
                    LayoutCachedTop =60
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5280
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =12
                    ForeColor =4210752
                    Name ="tbxVersion8"
                    ControlSource ="8-Facies mapping & grain size dist_"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =5280
                    LayoutCachedTop =60
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6540
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =13
                    ForeColor =4210752
                    Name ="tbxVersion99"
                    ControlSource ="9-After each field visit"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =6540
                    LayoutCachedTop =60
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6120
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =14
                    ForeColor =4210752
                    Name ="tbxVersion9"
                    ControlSource ="9-After each field visit"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =6120
                    LayoutCachedTop =60
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7380
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =15
                    ForeColor =4210752
                    Name ="tbxVersion11"
                    ControlSource ="11-Data management"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =7380
                    LayoutCachedTop =60
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6960
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =16
                    ForeColor =4210752
                    Name ="tbxVersion10"
                    ControlSource ="10-After each field season"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =6960
                    LayoutCachedTop =60
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8220
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =17
                    ForeColor =4210752
                    Name ="tbxVersion14"
                    ControlSource ="14-BLCA field methods, DINO sentinel site set up"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =8220
                    LayoutCachedTop =60
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7800
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =18
                    ForeColor =4210752
                    Name ="tbxVersion13"
                    ControlSource ="13-Revising the protocol"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =7800
                    LayoutCachedTop =60
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9060
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =19
                    ForeColor =4210752
                    Name ="tbxVersion17"
                    ControlSource ="17-RTK surveying part 1"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =9060
                    LayoutCachedTop =60
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8640
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =20
                    ForeColor =4210752
                    Name ="tbxVersion15"
                    ControlSource ="15-CURE methods"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =8640
                    LayoutCachedTop =60
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9900
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =21
                    ForeColor =4210752
                    Name ="tbxVersion20"
                    ControlSource ="20-CURE field methods"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =9900
                    LayoutCachedTop =60
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9480
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =22
                    ForeColor =4210752
                    Name ="tbxVersion18"
                    ControlSource ="18-RTK surveying part 2"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =60
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10320
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =23
                    ForeColor =4210752
                    Name ="tbxVersion100"
                    ControlSource ="100-Rapid assessment"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =10320
                    LayoutCachedTop =60
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10740
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =24
                    ForeColor =4210752
                    Name ="tbxVersion101"
                    ControlSource ="101-BLCA field methods"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =10740
                    LayoutCachedTop =60
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11160
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =25
                    ForeColor =4210752
                    Name ="tbxVersion102"
                    ControlSource ="102-DINO field methods"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =11160
                    LayoutCachedTop =60
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11580
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =26
                    ForeColor =4210752
                    Name ="tbxVersion103"
                    ControlSource ="103-CANY field methods"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =11580
                    LayoutCachedTop =60
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12000
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =27
                    ForeColor =4210752
                    Name ="tbxVersion104"
                    ControlSource ="104-DINO measuring vegetation"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =12000
                    LayoutCachedTop =60
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12420
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =28
                    ForeColor =4210752
                    Name ="tbxVersion105"
                    ControlSource ="105-DINO facies mapping & grain size dist_"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =12420
                    LayoutCachedTop =60
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12840
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =29
                    ForeColor =4210752
                    Name ="tbxVersion106"
                    ControlSource ="106-CANY & DINO equip_ lists"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =12840
                    LayoutCachedTop =60
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13260
                    Top =60
                    Width =360
                    Height =360
                    FontSize =8
                    TabIndex =30
                    ForeColor =4210752
                    Name ="tbxVersion107"
                    ControlSource ="107-CANY field methods"
                    InputMask ="#.##"
                    GridlineColor =10921638

                    LayoutCachedLeft =13260
                    LayoutCachedTop =60
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            Name ="ReportFooter"
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
' Report:       VersionTable
' Level:        Application report
' Version:      1.00
'
' Description:  VersionTable report object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, January 18, 2015
' References:   -
' Revisions:    BLC - 1/18/2017 - 1.00 - initial version
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
        'Me.lblDirections.Caption = m_Directions
    Else
        RaiseEvent InvalidDirections(Value)
    End If
End Property

Public Property Get Directions() As String
    Directions = m_Directions
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
' Events
'---------------------
' ---------------------------------
' Sub:          Report_Open
' Description:  Report opening event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/18/2017 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "SOP"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize calling form
    ToggleForm Me.CallingForm, -1


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[VersionTable Report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Report_Close
' Description:  Closing event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/18/2017 - initial version
' ---------------------------------
Private Sub Report_Close()
On Error GoTo Err_Handler

    'restore calling form
    ToggleForm Me.CallingForm, 0
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Close[VersionTable Report])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------
