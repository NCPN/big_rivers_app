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
    ItemSuffix =62
    Left =270
    Top =210
    Right =12735
    Bottom =10590
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8e1406dc71cce340
    End
    RecordSource ="tbl_wrk_Infest_Size"
    Caption ="Infestations by Size"
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
            ControlSource ="UnitCode"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =1440
                    Top =120
                    Width =7920
                    Height =600
                    FontSize =24
                    FontWeight =400
                    Name ="Label38"
                    Caption ="Infestations by Size Class"
                End
                Begin Line
                    Top =1380
                    Width =10800
                    Name ="Line43"
                End
                Begin Line
                    Top =1410
                    Width =10800
                    Name ="Line44"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2520
                    Top =840
                    Width =4680
                    Height =360
                    FontSize =16
                    Name ="Park_Name"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7380
                    Top =840
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
            Height =825
            Name ="GroupHeader0"
            Begin
                Begin Label
                    Left =60
                    Top =540
                    Width =1260
                    Height =270
                    FontSize =10
                    Name ="Species_Label"
                    Caption ="Species"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =2280
                    Top =540
                    Width =1500
                    Height =270
                    FontSize =10
                    Name ="CommonName_Label"
                    Caption ="Common Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =4500
                    Top =300
                    Width =1020
                    Height =510
                    FontSize =10
                    Name ="InfestTot_Label"
                    Caption ="Total Infestations"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =7080
                    Top =60
                    Width =3540
                    Height =240
                    FontSize =10
                    Name ="Class5_Label"
                    Caption ="Infestations Size Class"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Top =780
                    Width =10800
                    Name ="Line47"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Top =810
                    Width =10800
                    Name ="Line48"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =7080
                    Top =300
                    Width =3540
                    Height =240
                    FontSize =10
                    Name ="Label52"
                    Caption ="Number of Infestations"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =7080
                    Top =540
                    Width =659
                    Height =239
                    FontSize =10
                    Name ="Label54"
                    Caption ="Class 1"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =7800
                    Top =540
                    Width =659
                    Height =239
                    FontSize =10
                    Name ="Label55"
                    Caption ="Class 2"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =8520
                    Top =540
                    Width =659
                    Height =239
                    FontSize =10
                    Name ="Label56"
                    Caption ="Class 3"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =9240
                    Top =540
                    Width =659
                    Height =239
                    FontSize =10
                    Name ="Label57"
                    Caption ="Class 4"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =9960
                    Top =540
                    Width =659
                    Height =239
                    FontSize =10
                    Name ="Label58"
                    Caption ="Class 5"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =5700
                    Top =300
                    Width =1020
                    Height =510
                    FontSize =10
                    Name ="Label60"
                    Caption ="Priority 1 Totals"
                    Tag ="DetachedLabel"
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
                    Width =2160
                    Height =270
                    ColumnWidth =2295
                    Name ="Species"
                    ControlSource ="Species"
                    StatusBarText ="Species name"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2280
                    Top =60
                    Width =2160
                    Height =270
                    ColumnWidth =1800
                    TabIndex =1
                    Name ="CommonName"
                    ControlSource ="CommonName"
                    StatusBarText ="Common name"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4500
                    Top =60
                    Width =1020
                    Height =270
                    ColumnWidth =885
                    TabIndex =2
                    Name ="InfestTot"
                    ControlSource ="InfestTot"
                    StatusBarText ="Total infestations detected"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Top =60
                    Width =659
                    Height =270
                    TabIndex =3
                    Name ="Class1"
                    ControlSource ="Class1"
                    StatusBarText ="Infestation"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7800
                    Top =60
                    Width =659
                    Height =270
                    TabIndex =4
                    Name ="Class2"
                    ControlSource ="Class2"
                    StatusBarText ="totals"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =60
                    Width =659
                    Height =270
                    TabIndex =5
                    Name ="Class3"
                    ControlSource ="Class3"
                    StatusBarText ="by size"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9240
                    Top =60
                    Width =659
                    Height =270
                    TabIndex =6
                    Name ="Class4"
                    ControlSource ="Class4"
                    StatusBarText ="class"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9960
                    Top =60
                    Width =659
                    Height =270
                    TabIndex =7
                    Name ="Class5"
                    ControlSource ="Class5"
                    StatusBarText ="class"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5700
                    Top =60
                    Width =1019
                    Height =270
                    TabIndex =8
                    Name ="PriorityTot"
                    ControlSource ="PriorityTot"
                    StatusBarText ="Total priority 1 infestations"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =15
            Name ="GroupFooter1"
            Begin
                Begin Line
                    Left =60
                    Width =9240
                    Name ="Line49"
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
                    Name ="Text39"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6060
                    Top =240
                    Width =4560
                    Height =270
                    TabIndex =1
                    Name ="Text40"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                End
                Begin Line
                    Width =10800
                    Name ="Line50"
                End
                Begin Line
                    Top =30
                    Width =10800
                    Name ="Line51"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =480
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4500
                    Top =180
                    Width =1019
                    Height =270
                    Name ="InfestTot Grand Total Sum"
                    ControlSource ="=Sum([InfestTot])"
                    EventProcPrefix ="InfestTot_Grand_Total_Sum"

                End
                Begin Label
                    Left =120
                    Top =180
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
                    Top =180
                    Width =659
                    Height =270
                    TabIndex =1
                    Name ="Class1 Grand Total Sum"
                    ControlSource ="=Sum([Class1])"
                    EventProcPrefix ="Class1_Grand_Total_Sum"

                End
                Begin Label
                    Left =120
                    Top =180
                    Width =960
                    Height =240
                    FontSize =8
                    Name ="Label24"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7800
                    Top =180
                    Width =659
                    Height =270
                    TabIndex =2
                    Name ="Class2 Grand Total Sum"
                    ControlSource ="=Sum([Class2])"
                    EventProcPrefix ="Class2_Grand_Total_Sum"

                End
                Begin Label
                    Left =120
                    Top =180
                    Width =960
                    Height =240
                    FontSize =8
                    Name ="Label27"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =180
                    Width =659
                    Height =270
                    TabIndex =3
                    Name ="Class3 Grand Total Sum"
                    ControlSource ="=Sum([Class3])"
                    EventProcPrefix ="Class3_Grand_Total_Sum"

                End
                Begin Label
                    Left =120
                    Top =180
                    Width =960
                    Height =240
                    FontSize =8
                    Name ="Label30"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9240
                    Top =180
                    Width =659
                    Height =270
                    TabIndex =4
                    Name ="Class4 Grand Total Sum"
                    ControlSource ="=Sum([Class4])"
                    EventProcPrefix ="Class4_Grand_Total_Sum"

                End
                Begin Label
                    Left =120
                    Top =180
                    Width =960
                    Height =240
                    FontSize =8
                    Name ="Label33"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9960
                    Top =180
                    Width =659
                    Height =270
                    TabIndex =5
                    Name ="Class5 Grand Total Sum"
                    ControlSource ="=Sum([Class5])"
                    EventProcPrefix ="Class5_Grand_Total_Sum"

                End
                Begin Label
                    Left =120
                    Top =180
                    Width =960
                    Height =240
                    FontSize =8
                    Name ="Label36"
                    Caption ="Grand Total"
                End
                Begin Line
                    Top =60
                    Width =10800
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5700
                    Top =180
                    Width =1019
                    Height =270
                    TabIndex =6
                    Name ="Text61"
                    ControlSource ="=Sum([PriorityTot])"

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
