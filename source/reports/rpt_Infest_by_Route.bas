Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =126
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =9
    ItemSuffix =52
    Left =105
    Top =270
    Right =11160
    Bottom =7290
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7060c1afafcbe340
    End
    RecordSource ="tbl_wrk_Infest_Route"
    Caption ="rpt_Infest_by_Route"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xd0020000d0020000d0020000d002000000000000302a00008601000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnActivate ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =9
            FontWeight =700
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =18
            BorderLineStyle =0
            BackStyle =0
            FontName ="Times New Roman"
            AsianLineBreak =255
        End
        Begin ListBox
            TextFontFamily =18
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Times New Roman"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =18
            BorderLineStyle =0
            BackStyle =0
            FontName ="Times New Roman"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="RouteType"
        End
        Begin BreakLevel
            ControlSource ="PlotID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1500
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =3240
                    Top =240
                    Width =4320
                    Height =540
                    FontSize =24
                    FontWeight =400
                    Name ="Label28"
                    Caption ="Infestations by Route"
                End
                Begin Line
                    Top =60
                    Width =10800
                    Name ="Line31"
                End
                Begin Line
                    Top =90
                    Width =10800
                    Name ="Line32"
                End
                Begin Line
                    Top =1380
                    Width =10800
                    Name ="Line33"
                End
                Begin Line
                    Top =1410
                    Width =10800
                    Name ="Line34"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2880
                    Top =900
                    Width =3720
                    Height =360
                    FontSize =16
                    Name ="Park_Name"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6660
                    Top =900
                    Width =1260
                    Height =360
                    FontSize =16
                    TabIndex =1
                    Name ="Visit_Year"

                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1155
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =1440
                    Width =2310
                    Height =300
                    FontSize =10
                    Name ="RouteType"
                    ControlSource ="RouteType"
                    StatusBarText ="Route type for grouping"

                    Begin
                        Begin Label
                            Left =60
                            Width =1320
                            Height =300
                            FontSize =10
                            Name ="RouteType_Label"
                            Caption ="Type of Route"
                        End
                    End
                End
                Begin Label
                    Left =60
                    Top =540
                    Width =1440
                    Height =270
                    FontSize =10
                    Name ="PlotID_Label"
                    Caption ="Route"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =3420
                    Top =420
                    Width =1200
                    Height =540
                    FontSize =10
                    Name ="RouteLength_Label"
                    Caption ="Route Length (m)"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3420
                    LayoutCachedTop =420
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =960
                End
                Begin Label
                    TextAlign =2
                    Left =6120
                    Top =720
                    Width =900
                    Height =270
                    FontSize =10
                    Name ="InfestTot_Label"
                    Caption ="Total"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6120
                    LayoutCachedTop =720
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =990
                End
                Begin Label
                    TextAlign =2
                    Left =7080
                    Top =720
                    Width =1020
                    Height =270
                    FontSize =10
                    Name ="PriorityTot_Label"
                    Caption ="Priority 1"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7080
                    LayoutCachedTop =720
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =990
                End
                Begin Label
                    TextAlign =2
                    Left =8700
                    Top =720
                    Width =900
                    Height =270
                    FontSize =10
                    Name ="TotPct_Label"
                    Caption ="Total"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8700
                    LayoutCachedTop =720
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =990
                End
                Begin Label
                    TextAlign =2
                    Left =9660
                    Top =720
                    Width =1020
                    Height =270
                    FontSize =10
                    Name ="PriorityPct_Label"
                    Caption ="Priority 1"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9660
                    LayoutCachedTop =720
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =990
                End
                Begin Line
                    BorderWidth =1
                    Top =390
                    Width =10800
                    Name ="Line35"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Top =360
                    Width =10800
                    Name ="Line36"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Top =1020
                    Width =10800
                    Name ="Line37"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Top =1020
                    Width =10800
                    Name ="Line38"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =6120
                    Top =420
                    Width =1980
                    Height =270
                    FontSize =10
                    Name ="Label44"
                    Caption ="Infestations"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6120
                    LayoutCachedTop =420
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =690
                End
                Begin Line
                    BorderWidth =1
                    Left =6120
                    Top =720
                    Width =1980
                    Name ="Line45"
                    LayoutCachedLeft =6120
                    LayoutCachedTop =720
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =720
                End
                Begin Label
                    TextAlign =2
                    Left =8700
                    Top =420
                    Width =1995
                    Height =270
                    FontSize =10
                    Name ="Label46"
                    Caption ="Infestations/ha"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8700
                    LayoutCachedTop =420
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =690
                End
                Begin Label
                    TextAlign =2
                    Left =4800
                    Top =720
                    Width =1200
                    Height =255
                    Name ="Label47"
                    Caption ="Area (ha)"
                    LayoutCachedLeft =4800
                    LayoutCachedTop =720
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =975
                End
                Begin Line
                    BorderWidth =1
                    Left =8700
                    Top =720
                    Width =1980
                    Name ="Line50"
                    LayoutCachedLeft =8700
                    LayoutCachedTop =720
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =720
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =3240
                    Height =270
                    ColumnWidth =2505
                    Name ="PlotID"
                    ControlSource ="PlotID"
                    StatusBarText ="Route"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3420
                    Top =60
                    Width =1200
                    Height =270
                    ColumnWidth =1230
                    TabIndex =1
                    Name ="RouteLength"
                    ControlSource ="RouteLength"
                    StatusBarText ="Length of route in meters"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =60
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6120
                    Top =60
                    Width =900
                    Height =270
                    ColumnWidth =1905
                    TabIndex =2
                    Name ="InfestTot"
                    ControlSource ="InfestTot"
                    StatusBarText ="Total infestations detected"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =60
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Top =60
                    Width =1020
                    Height =270
                    TabIndex =3
                    Name ="PriorityTot"
                    ControlSource ="PriorityTot"
                    StatusBarText ="Total priority 1 infestations detected"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =60
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8700
                    Top =60
                    Width =900
                    Height =270
                    TabIndex =4
                    Name ="TotPct"
                    ControlSource ="TotPct"
                    StatusBarText ="Infestations per 100 m2"

                    LayoutCachedLeft =8700
                    LayoutCachedTop =60
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9660
                    Top =60
                    Width =990
                    Height =270
                    TabIndex =5
                    Name ="PriorityPct"
                    ControlSource ="PriorityPct"
                    StatusBarText ="Priority 1 infestations per 100 m2"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =60
                    LayoutCachedWidth =10650
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4800
                    Top =60
                    Width =1199
                    Height =270
                    TabIndex =6
                    Name ="RouteArea"
                    ControlSource ="RouteArea"
                    StatusBarText ="Area of route in hectares"

                    LayoutCachedLeft =4800
                    LayoutCachedTop =60
                    LayoutCachedWidth =5999
                    LayoutCachedHeight =330
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =780
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Width =10560
                    Name ="Text2"
                    ControlSource ="=\"Summary for \" & \"'RouteType' = \" & \" \" & [RouteType] & \" (\" & Count(*)"
                        " & \" \" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""

                    LayoutCachedLeft =60
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =240
                End
                Begin Label
                    Left =60
                    Top =240
                    Width =390
                    Height =240
                    FontSize =8
                    Name ="Label3"
                    Caption ="Sum"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3420
                    Top =240
                    Width =1200
                    TabIndex =1
                    Name ="Sum Of RouteLength"
                    ControlSource ="=Sum([RouteLength])"
                    EventProcPrefix ="Sum_Of_RouteLength"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =240
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6120
                    Top =240
                    Width =900
                    TabIndex =2
                    Name ="Sum Of InfestTot"
                    ControlSource ="=Sum([InfestTot])"
                    EventProcPrefix ="Sum_Of_InfestTot"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =240
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Top =240
                    Width =1020
                    TabIndex =3
                    Name ="Sum Of PriorityTot"
                    ControlSource ="=Sum([PriorityTot])"
                    EventProcPrefix ="Sum_Of_PriorityTot"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =240
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8700
                    Top =480
                    Width =900
                    TabIndex =4
                    Name ="Route Type Total Infestations/ha"
                    ControlSource ="=(Sum([InfestTot]))/Sum([RouteArea])"
                    EventProcPrefix ="Route_Type_Total_Infestations_ha"

                    LayoutCachedLeft =8700
                    LayoutCachedTop =480
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9660
                    Top =480
                    Width =1050
                    TabIndex =5
                    Name ="Route Type Priority Infestations/ha"
                    ControlSource ="=(Sum([PriorityTot]))/Sum([RouteArea])"
                    EventProcPrefix ="Route_Type_Priority_Infestations_ha"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =480
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =720
                End
                Begin Line
                    Left =60
                    Width =9240
                    Name ="Line39"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4800
                    Top =240
                    Width =1200
                    TabIndex =6
                    Name ="Text48"
                    ControlSource ="=Sum([RouteArea])"

                    LayoutCachedLeft =4800
                    LayoutCachedTop =240
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =480
                End
                Begin Line
                    BorderWidth =1
                    Width =10800
                    Name ="Line51"
                    Tag ="DetachedLabel"
                    LayoutCachedWidth =10800
                End
            End
        End
        Begin PageFooter
            Height =510
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =4560
                    Height =270
                    Name ="Text29"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4740
                    Top =240
                    Width =4560
                    Height =270
                    TabIndex =1
                    Name ="Text30"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                End
                Begin Line
                    Width =10800
                    Name ="Line40"
                End
                Begin Line
                    Top =30
                    Width =10800
                    Name ="Line41"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =330
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3420
                    Width =1200
                    Height =270
                    Name ="RouteLength Grand Total Sum"
                    ControlSource ="=Sum([RouteLength])"
                    EventProcPrefix ="RouteLength_Grand_Total_Sum"

                    LayoutCachedLeft =3420
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =270
                End
                Begin Label
                    Left =60
                    Width =960
                    Height =240
                    FontSize =8
                    Name ="Label18"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6120
                    Width =900
                    Height =270
                    TabIndex =1
                    Name ="InfestTot Grand Total Sum"
                    ControlSource ="=Sum([InfestTot])"
                    EventProcPrefix ="InfestTot_Grand_Total_Sum"

                    LayoutCachedLeft =6120
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =270
                End
                Begin Label
                    Left =60
                    Width =960
                    Height =240
                    FontSize =8
                    Name ="Label21"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Width =1020
                    Height =270
                    TabIndex =2
                    Name ="PriorityTot Grand Total Sum"
                    ControlSource ="=Sum([PriorityTot])"
                    EventProcPrefix ="PriorityTot_Grand_Total_Sum"

                    LayoutCachedLeft =7080
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =270
                End
                Begin Label
                    Left =60
                    Width =960
                    Height =240
                    FontSize =8
                    Name ="Label24"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8700
                    Width =900
                    Height =270
                    TabIndex =3
                    Name ="Overall Total Infestations/ha"
                    ControlSource ="=(Sum([InfestTot]))/Sum([RouteArea])"
                    EventProcPrefix ="Overall_Total_Infestations_ha"

                    LayoutCachedLeft =8700
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9660
                    Width =1050
                    Height =270
                    TabIndex =4
                    Name ="Overall Priority Infestations/ha"
                    ControlSource ="=(Sum([PriorityTot]))/Sum([RouteArea])"
                    EventProcPrefix ="Overall_Priority_Infestations_ha"

                    LayoutCachedLeft =9660
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4800
                    Width =1200
                    TabIndex =5
                    Name ="Text49"
                    ControlSource ="=Sum([RouteArea])"

                    LayoutCachedLeft =4800
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =240
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

Private Sub Report_Activate()
  Me!Park_Name = DLookup("[ParkName]", "tlu_Parks", "[ParkCode]= '" & Left(OpenArgs, 4) & "'")
  Me!Visit_Year = Right(OpenArgs, 4)
End Sub
