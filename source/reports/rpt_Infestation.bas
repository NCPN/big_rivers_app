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
    ItemSuffix =38
    Left =270
    Top =210
    Right =13890
    Bottom =8190
    DatasheetGridlinesColor =12632256
    Filter ="([Unit_Code] = 'BLCA' AND Visit_Year = 2012)"
    RecSrcDt = Begin
        0xa95fc136d4cce340
    End
    RecordSource ="qry_Infest"
    Caption ="rpt_Infestation"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xd0020000d0020000d0020000d002000000000000302a00000e01000001000000 ,
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
            ControlSource ="Plot_ID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Species"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1200
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =2160
                    Width =5760
                    Height =540
                    FontSize =24
                    FontWeight =400
                    Name ="Label18"
                    Caption ="Exotic Plant Invasions"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Top =720
                    Width =3720
                    Height =360
                    FontSize =16
                    Name ="Park_Name"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4680
                    Top =720
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
            Height =510
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =9000
                    Height =390
                    ColumnWidth =3240
                    FontSize =14
                    FontWeight =700
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Route name"

                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =450
                End
                Begin Line
                    BorderWidth =2
                    Width =10800
                    Name ="Line34"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1020
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =1320
                    Top =60
                    Width =2310
                    Height =300
                    ColumnWidth =3225
                    FontSize =12
                    Name ="Species"
                    ControlSource ="Species"
                    StatusBarText ="Wyoming Species (Dorn 2001)"

                    Begin
                        Begin Label
                            Left =300
                            Top =60
                            Width =900
                            Height =300
                            FontSize =12
                            Name ="Species_Label"
                            Caption ="Species"
                        End
                    End
                End
                Begin Label
                    Left =3900
                    Top =60
                    Width =1680
                    Height =299
                    FontSize =12
                    Name ="Master_Common_Name_Label"
                    Caption ="Common Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =3780
                    Top =600
                    Width =720
                    Height =270
                    FontSize =10
                    Name ="Pulled_Label"
                    Caption ="Pulled?"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =4740
                    Top =600
                    Width =1215
                    Height =270
                    FontSize =10
                    Name ="Growth_Stage_Label"
                    Caption ="Growth Stage"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =6180
                    Top =600
                    Width =1275
                    Height =270
                    FontSize =10
                    Name ="N_Coord_Label"
                    Caption ="Northing"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =7680
                    Top =600
                    Width =1275
                    Height =270
                    FontSize =10
                    Name ="E_Coord_Label"
                    Caption ="Easting"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =2400
                    Top =420
                    Width =1140
                    Height =479
                    FontSize =10
                    Name ="Cover_Class_Label"
                    Caption ="Infestation Cover Class"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =1200
                    Top =420
                    Width =990
                    Height =480
                    FontSize =10
                    Name ="Size_Class_Label"
                    Caption ="InfestationSize Class"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Top =930
                    Width =10800
                    Name ="Line28"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5700
                    Top =60
                    Width =2460
                    Height =299
                    ColumnWidth =2115
                    FontSize =12
                    TabIndex =1
                    Name ="Master_Common_Name"
                    ControlSource ="Master_Common_Name"
                    StatusBarText ="Master Common Name"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =270
            Name ="Detail"
            Begin
                Begin CheckBox
                    Visible = NotDefault
                    Left =3780
                    Name ="Pulled"
                    ControlSource ="Pulled"
                    StatusBarText ="Was weed pulled"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4740
                    Width =1200
                    Height =270
                    TabIndex =1
                    Name ="Growth_Stage"
                    ControlSource ="Growth_Stage"
                    StatusBarText ="Growth stage of trees"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6180
                    Width =1260
                    Height =270
                    TabIndex =2
                    Name ="N_Coord"
                    ControlSource ="N_Coord"
                    StatusBarText ="UTM North of Infestation"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Width =1260
                    Height =270
                    TabIndex =3
                    Name ="E_Coord"
                    ControlSource ="E_Coord"
                    StatusBarText ="UTM East of Infestation"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2400
                    Width =1140
                    Height =270
                    TabIndex =4
                    Name ="Cover_Class"
                    ControlSource ="Cover_Class"
                    StatusBarText ="Cover class for reports"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1200
                    Width =990
                    Height =270
                    TabIndex =5
                    Name ="Size_Class"
                    ControlSource ="Size_Class"
                    StatusBarText ="Size class for report"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3780
                    Width =720
                    Height =270
                    TabIndex =6
                    Name ="txtPulled"
                    ControlSource ="=IIf([Pulled]=-1,\"Yes\",\"No\")"

                End
            End
        End
        Begin PageFooter
            Height =480
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =180
                    Width =4560
                    Height =270
                    Name ="Text19"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4740
                    Top =180
                    Width =4560
                    Height =270
                    TabIndex =1
                    Name ="Text20"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                End
                Begin Line
                    Width =9360
                    Name ="Line29"
                End
                Begin Line
                    Top =30
                    Width =9360
                    Name ="Line30"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
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

Private Sub Report_Open(Cancel As Integer)
'  If Left(OpenArgs, 4) = "FOBU" Then
'    Me.RecordSource = "qry_infest_wy"
'  ElseIf Left(OpenArgs, 4) = "COLM" Then
'    Me.RecordSource = "qry_infest_co"
'  Else
'    Me.RecordSource = "qry_infest_ut"
'  End If

End Sub
