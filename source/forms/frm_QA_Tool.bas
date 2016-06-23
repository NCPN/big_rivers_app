Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =2
    GridX =24
    GridY =24
    Width =14415
    DatasheetFontHeight =10
    ItemSuffix =677
    Right =7650
    Bottom =10995
    DatasheetGridlinesColor =12632256
    Filter ="[Query_name] = \"qa_a111_Overview_transect_pt_duplicates\" AND [Time_frame] = \""
        "2014\" AND [Data_scope] = 0"
    RecSrcDt = Begin
        0xdef19da9b06be340
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="tbl_QA_Results"
    Caption =" Data Validation and Quality Review Tool"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            OldBorderStyle =1
            TextAlign =1
            FontWeight =700
            BackColor =8388608
            BorderColor =8388608
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =2
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderWidth =2
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            BorderColor =16776960
        End
        Begin CommandButton
            FontItalic = NotDefault
            FontSize =8
            ForeColor =-2147483630
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =4
            BorderWidth =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =8388608
        End
        Begin CheckBox
            SpecialEffect =4
            BorderWidth =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =8388608
        End
        Begin OptionGroup
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
        End
        Begin BoundObjectFrame
            BorderLineStyle =0
            BackStyle =0
            BorderColor =16776960
        End
        Begin TextBox
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin ListBox
            BorderLineStyle =0
            BackColor =8421376
            ForeColor =16777215
            BorderColor =16776960
            FontName ="Arial"
        End
        Begin ComboBox
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =16776960
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =16776960
        End
        Begin ToggleButton
            FontItalic = NotDefault
            FontSize =8
            ForeColor =-2147483630
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin Tab
            FontItalic = NotDefault
            BackStyle =0
            FontWeight =700
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =11535
            BackColor =13025979
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Top =495
                    Width =14415
                    Height =11040
                    Name ="PageTabs"
                    OnChange ="[Event Procedure]"

                    LayoutCachedTop =495
                    LayoutCachedWidth =14415
                    LayoutCachedHeight =11535
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =120
                            Top =900
                            Width =14160
                            Height =10503
                            Name ="pgResults"
                            Caption =" Results summary"
                            LayoutCachedLeft =120
                            LayoutCachedTop =900
                            LayoutCachedWidth =14280
                            LayoutCachedHeight =11403
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextAlign =0
                                    Left =120
                                    Top =900
                                    Width =3300
                                    Height =423
                                    FontWeight =400
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labOverview"
                                    Caption ="* Double-click on the label to change sort order.  Click on a query name to open"
                                        "."
                                    ControlTipText ="View mode"
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =215
                                    Left =9900
                                    Top =960
                                    Width =1500
                                    Height =300
                                    Name ="cmdRefresh"
                                    Caption ="Refresh results"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Run the validation queries and refresh the results summary"

                                    LayoutCachedLeft =9900
                                    LayoutCachedTop =960
                                    LayoutCachedWidth =11400
                                    LayoutCachedHeight =1260
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =215
                                    Left =11640
                                    Top =960
                                    Width =2100
                                    Height =300
                                    TabIndex =1
                                    Name ="cmdViewReport"
                                    Caption ="View summary report"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="View the quality review results as a report"

                                    LayoutCachedLeft =11640
                                    LayoutCachedTop =960
                                    LayoutCachedWidth =13740
                                    LayoutCachedHeight =1260
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin Subform
                                    CanShrink = NotDefault
                                    OverlapFlags =247
                                    Left =120
                                    Top =1350
                                    Width =14160
                                    Height =10053
                                    TabIndex =2
                                    BorderColor =0
                                    Name ="subResults"
                                    SourceObject ="Form.fsub_QA_Results"
                                    LinkChildFields ="Time_frame;Data_scope"
                                    LinkMasterFields ="cmbTimeframe;optgScope"

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =1350
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =11403
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    RowSourceTypeInt =1
                                    SpecialEffect =2
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =5490
                                    Top =997
                                    Width =1170
                                    TabIndex =3
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    Name ="cmbTypeFilter"
                                    RowSourceType ="Value List"
                                    RowSource ="1;Critical;2;Warning;3;Information"
                                    ColumnWidths ="0;2160"
                                    StatusBarText ="Filter by query type"
                                    AfterUpdate ="[Event Procedure]"
                                    ControlTipText ="Filter by query type"

                                    LayoutCachedLeft =5490
                                    LayoutCachedTop =997
                                    LayoutCachedWidth =6660
                                    LayoutCachedHeight =1237
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =4260
                                            Top =990
                                            Width =1110
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labTypeFilter"
                                            Caption ="Query type:"
                                            LayoutCachedLeft =4260
                                            LayoutCachedTop =990
                                            LayoutCachedWidth =5370
                                            LayoutCachedHeight =1230
                                        End
                                    End
                                End
                                Begin ToggleButton
                                    FontItalic = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    Left =6780
                                    Top =960
                                    Width =480
                                    Height =300
                                    FontWeight =400
                                    TabIndex =4
                                    ForeColor =0
                                    Name ="togFilterByType"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    Caption ="Filter on"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                                        0xadadadadadadadad
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Turn the type filter on or off"

                                    LayoutCachedLeft =6780
                                    LayoutCachedTop =960
                                    LayoutCachedWidth =7260
                                    LayoutCachedHeight =1260
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    RowSourceTypeInt =1
                                    SpecialEffect =2
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8220
                                    Top =997
                                    Width =900
                                    TabIndex =5
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    Name ="cmbDoneFilter"
                                    RowSourceType ="Value List"
                                    RowSource ="True;False"
                                    StatusBarText ="Filter by the 'Done' flag"
                                    AfterUpdate ="[Event Procedure]"
                                    ControlTipText ="Filter by the 'Done' flag"

                                    LayoutCachedLeft =8220
                                    LayoutCachedTop =997
                                    LayoutCachedWidth =9120
                                    LayoutCachedHeight =1237
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =7500
                                            Top =997
                                            Width =600
                                            Height =228
                                            FontWeight =400
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labDoneFilter"
                                            Caption ="Done:"
                                            LayoutCachedLeft =7500
                                            LayoutCachedTop =997
                                            LayoutCachedWidth =8100
                                            LayoutCachedHeight =1225
                                        End
                                    End
                                End
                                Begin ToggleButton
                                    FontItalic = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    Left =9240
                                    Top =960
                                    Width =480
                                    Height =300
                                    FontWeight =400
                                    TabIndex =6
                                    ForeColor =0
                                    Name ="togFilterByDone"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    Caption ="Filter on"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                                        0xadadadadadadadad
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Turn the 'Done' filter on or off"

                                    LayoutCachedLeft =9240
                                    LayoutCachedTop =960
                                    LayoutCachedWidth =9720
                                    LayoutCachedHeight =1260
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =900
                            Width =14160
                            Height =10500
                            Name ="pgQueryViews"
                            Caption =" View and fix query results"
                            LayoutCachedLeft =120
                            LayoutCachedTop =900
                            LayoutCachedWidth =14280
                            LayoutCachedHeight =11400
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =8040
                                    Top =1035
                                    Width =1320
                                    Height =317
                                    Name ="cmdDesignView"
                                    Caption ="Design view"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open the selected query in design view"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =1260
                                    Top =1050
                                    Width =6660
                                    Height =252
                                    TabIndex =1
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                                    Name ="selObject"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT MSysObjects.Name AS Query_name FROM MSysObjects WHERE (((MSysObjects.Name"
                                        ") Like \"qa_*\") AND ((MSysObjects.Type)=5)) ORDER BY MSysObjects.Name; "
                                    AfterUpdate ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =120
                                            Top =1050
                                            Width =1110
                                            Height =270
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labObject"
                                            Caption ="Query name"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =10080
                                    Top =1050
                                    Width =2220
                                    Height =252
                                    TabIndex =2
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="txtUser"
                                    ControlSource ="QA_user"
                                    OnDirty ="[Event Procedure]"

                                    LayoutCachedLeft =10080
                                    LayoutCachedTop =1050
                                    LayoutCachedWidth =12300
                                    LayoutCachedHeight =1302
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =9480
                                            Top =1050
                                            Width =570
                                            Height =270
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labUser"
                                            Caption ="QA by"
                                            LayoutCachedLeft =9480
                                            LayoutCachedTop =1050
                                            LayoutCachedWidth =10050
                                            LayoutCachedHeight =1320
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =12360
                                    Top =1050
                                    Width =1920
                                    Height =252
                                    TabIndex =3
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="txtRemedy_date"
                                    ControlSource ="Remedy_date"
                                    Format ="mm/dd/yy"

                                    LayoutCachedLeft =12360
                                    LayoutCachedTop =1050
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =1302
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    EnterKeyBehavior = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    ScrollBars =2
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =1260
                                    Top =1410
                                    Width =13020
                                    Height =660
                                    TabIndex =4
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="txtQueryDesc"
                                    ControlSource ="Query_description"
                                    StatusBarText ="Description of the query"
                                    OnDirty ="[Event Procedure]"

                                    LayoutCachedLeft =1260
                                    LayoutCachedTop =1410
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =2070
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =120
                                            Top =1410
                                            Width =1035
                                            Height =495
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labQueryDesc"
                                            Caption ="Query description"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    EnterKeyBehavior = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    ScrollBars =2
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =1260
                                    Top =2190
                                    Width =13020
                                    Height =810
                                    TabIndex =5
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="txtRemedy"
                                    ControlSource ="Remedy_desc"
                                    StatusBarText ="Details about actions taken and/or not taken to resolve errors"

                                    LayoutCachedLeft =1260
                                    LayoutCachedTop =2190
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =3000
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =120
                                            Top =2190
                                            Width =810
                                            Height =495
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labRemedy"
                                            Caption ="Remedy details"
                                        End
                                    End
                                End
                                Begin Subform
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    SpecialEffect =2
                                    Left =120
                                    Top =3555
                                    Width =14160
                                    Height =7845
                                    TabIndex =6
                                    BorderColor =0
                                    Name ="subQueryResults"

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =3555
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =11400
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =255
                                            TextAlign =0
                                            Left =120
                                            Top =3315
                                            Width =1212
                                            Height =252
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labQueryResults"
                                            Caption ="Query results"
                                            LayoutCachedLeft =120
                                            LayoutCachedTop =3315
                                            LayoutCachedWidth =1332
                                            LayoutCachedHeight =3567
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    FELineBreak = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3270
                                    Top =3180
                                    Width =606
                                    Height =255
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =7
                                    BackColor =8454143
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="txtEditQuery"
                                    FontName ="Tahoma"
                                    AsianLineBreak =255

                                    LayoutCachedLeft =3270
                                    LayoutCachedTop =3180
                                    LayoutCachedWidth =3876
                                    LayoutCachedHeight =3435
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1440
                                            Top =3180
                                            Width =1770
                                            Height =255
                                            FontSize =9
                                            FontWeight =400
                                            BackColor =16777215
                                            BorderColor =0
                                            ForeColor =0
                                            Name ="labEditQuery"
                                            Caption ="Edit results directly?"
                                            FontName ="Tahoma"
                                            LayoutCachedLeft =1440
                                            LayoutCachedTop =3180
                                            LayoutCachedWidth =3210
                                            LayoutCachedHeight =3435
                                        End
                                    End
                                End
                                Begin CommandButton
                                    Enabled = NotDefault
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =4620
                                    Top =3120
                                    Width =1080
                                    Height =317
                                    TabIndex =8
                                    ForeColor =0
                                    Name ="cmdAutoFix"
                                    Caption ="Auto-fix"
                                    StatusBarText ="Run a pre-built query to automatically fix all the records"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Run a pre-built query to automatically fix all the records"

                                    LayoutCachedLeft =4620
                                    LayoutCachedTop =3120
                                    LayoutCachedWidth =5700
                                    LayoutCachedHeight =3437
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =6240
                                    Top =3120
                                    Width =2040
                                    Height =317
                                    TabIndex =9
                                    ForeColor =0
                                    Name ="cmdOpenRecord"
                                    Caption ="Open selected record"
                                    StatusBarText ="Open the form / query / table specified in the query to the record selected in t"
                                        "he subform"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open the form / query / table specified in the query to the record selected in t"
                                        "he subform"

                                    LayoutCachedLeft =6240
                                    LayoutCachedTop =3120
                                    LayoutCachedWidth =8280
                                    LayoutCachedHeight =3437
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =8580
                                    Top =3120
                                    Height =317
                                    TabIndex =10
                                    ForeColor =0
                                    Name ="cmdOpenBrowser"
                                    Caption ="Data browser"
                                    StatusBarText ="Open the project data browser"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open the project data browser"

                                    LayoutCachedLeft =8580
                                    LayoutCachedTop =3120
                                    LayoutCachedWidth =10020
                                    LayoutCachedHeight =3437
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =10500
                                    Top =3061
                                    Width =426
                                    Height =426
                                    FontWeight =400
                                    TabIndex =11
                                    Name ="cmdExport"
                                    Caption ="Export to Excel"
                                    StatusBarText ="Export the results of the selected query to Excel"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadada0000000dadadadadd00000dadadadada ,
                                        0xad000dadadadadaddad0dadadadadadaadadadadad72727ddada2727272f272a ,
                                        0xadad727272f272addada27272f2727daadada272f27272addadada2f2727dada ,
                                        0xadada2f272727daddada2f27272727daadad72727d7272addada2727dad727da ,
                                        0xadadadadadadadad
                                    End
                                    FontName ="Tahoma"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Export the results of the selected query to Excel"

                                    LayoutCachedLeft =10500
                                    LayoutCachedTop =3061
                                    LayoutCachedWidth =10926
                                    LayoutCachedHeight =3487
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =11040
                                    Top =3060
                                    Width =426
                                    Height =426
                                    FontWeight =400
                                    TabIndex =12
                                    Name ="cmdCloseup"
                                    Caption ="Zoom"
                                    StatusBarText ="Open the selected query in a new window"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadada00adadadadadadad000adadadadadada ,
                                        0xa000adadadadadadda000a700007dadaada0000888800daddada07ee888870da ,
                                        0xada708e88888807ddad08e888888880aada088888888880ddad088888888e80a ,
                                        0xada088888888e80ddad70888888ee07aadad07888eee70addadad00888800ada ,
                                        0xadadad700007adad
                                    End
                                    FontName ="Tahoma"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Open the selected query in a new window"

                                    LayoutCachedLeft =11040
                                    LayoutCachedTop =3060
                                    LayoutCachedWidth =11466
                                    LayoutCachedHeight =3486
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =12600
                                    Top =3120
                                    Width =1020
                                    Height =317
                                    TabIndex =13
                                    ForeColor =0
                                    Name ="cmdRequery"
                                    Caption ="Requery"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Requery the results set for the selected query"

                                    LayoutCachedLeft =12600
                                    LayoutCachedTop =3120
                                    LayoutCachedWidth =13620
                                    LayoutCachedHeight =3437
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =900
                            Width =14160
                            Height =10503
                            Name ="pgDataTables"
                            Caption =" Browse data tables"
                            LayoutCachedLeft =120
                            LayoutCachedTop =900
                            LayoutCachedWidth =14280
                            LayoutCachedHeight =11403
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListRows =20
                                    ListWidth =11520
                                    Left =840
                                    Top =1050
                                    Width =4320
                                    Height =252
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"0\";\"0\""
                                    Name ="selTable"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tsys_Link_Tables.Link_table, tsys_Link_Tables.Description_text FROM tsys_"
                                        "Link_Tables WHERE (((tsys_Link_Tables.Link_table) Like \"tbl_*\" And (tsys_Link_"
                                        "Tables.Link_table)<>\"tbl_QA_Results\")) OR (((tsys_Link_Tables.Link_table)=\"tl"
                                        "u_Project_Crew\")) OR (((tsys_Link_Tables.Link_table)=\"tlu_Project_Taxa\")) OR "
                                        "(((tsys_Link_Tables.Link_table)=\"tlu_Park_Taxa\")); "
                                    ColumnWidths ="4320;7200"
                                    AfterUpdate ="[Event Procedure]"
                                    OnEnter ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =180
                                            Top =1050
                                            Width =585
                                            Height =270
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labTable"
                                            Caption ="Table:"
                                        End
                                    End
                                End
                                Begin Subform
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    SpecialEffect =2
                                    Left =120
                                    Top =1698
                                    Width =14160
                                    Height =9705
                                    TabIndex =1
                                    BorderColor =0
                                    Name ="subDataTables"

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =1698
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =11403
                                End
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =0
                                    Left =5340
                                    Top =900
                                    Width =7716
                                    Height =699
                                    FontWeight =400
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labEditWarning"
                                    Caption =" Warning:  This is a last resort!  If possible, open the records needing fixes w"
                                        "ithin the data entry form.  Also, when making manual edits in data tables, pleas"
                                        "e be sure to update the updated_date and updated_by fields if they are present i"
                                        "n the table."
                                    ControlTipText ="View mode"
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =13320
                    Top =60
                    Width =720
                    Height =354
                    TabIndex =1
                    ForeColor =0
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close the form"

                    LayoutCachedLeft =13320
                    LayoutCachedTop =60
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =414
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =85
                    Left =10380
                    Top =60
                    Width =1914
                    Height =355
                    TabIndex =2
                    BackColor =16777215
                    BorderColor =0
                    Name ="optgMode"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="Change the form mode"

                    LayoutCachedLeft =10380
                    LayoutCachedTop =60
                    LayoutCachedWidth =12294
                    LayoutCachedHeight =415
                    Begin
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =11460
                            Top =144
                            OptionValue =1
                            BorderColor =0
                            Name ="optEditMode"

                            LayoutCachedLeft =11460
                            LayoutCachedTop =144
                            LayoutCachedWidth =11720
                            LayoutCachedHeight =384
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =11694
                                    Top =120
                                    Width =390
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labEditMode"
                                    Caption ="Edit"
                                    ControlTipText ="Edit mode"
                                    LayoutCachedLeft =11694
                                    LayoutCachedTop =120
                                    LayoutCachedWidth =12084
                                    LayoutCachedHeight =390
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =10500
                            Top =150
                            OptionValue =0
                            BorderColor =0
                            Name ="optViewMode"

                            LayoutCachedLeft =10500
                            LayoutCachedTop =150
                            LayoutCachedWidth =10760
                            LayoutCachedHeight =390
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =10734
                                    Top =120
                                    Width =495
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labViewMode"
                                    Caption ="View"
                                    ControlTipText ="View mode"
                                    LayoutCachedLeft =10734
                                    LayoutCachedTop =120
                                    LayoutCachedWidth =11229
                                    LayoutCachedHeight =390
                                End
                            End
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =85
                    Left =5100
                    Top =60
                    Width =4800
                    Height =355
                    TabIndex =4
                    BackColor =16777215
                    BorderColor =0
                    Name ="optgScope"
                    DefaultValue ="0"
                    ControlTipText ="Scope of the data included in the validation queries: uncertified events, certif"
                        "ied events, or both?"

                    LayoutCachedLeft =5100
                    LayoutCachedTop =60
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =415
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =0
                            OldBorderStyle =0
                            OverlapFlags =215
                            TextAlign =0
                            Left =5160
                            Top =120
                            Width =945
                            Height =255
                            FontWeight =400
                            BackColor =13025979
                            BorderColor =0
                            ForeColor =0
                            Name ="labIncludeCertified"
                            Caption ="Data scope:"
                            LayoutCachedLeft =5160
                            LayoutCachedTop =120
                            LayoutCachedWidth =6105
                            LayoutCachedHeight =375
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =6240
                            Top =144
                            OptionValue =0
                            BorderColor =0
                            Name ="optUncertOnly"

                            LayoutCachedLeft =6240
                            LayoutCachedTop =144
                            LayoutCachedWidth =6500
                            LayoutCachedHeight =384
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =6480
                                    Top =120
                                    Width =1050
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labUncertOnly"
                                    Caption ="Uncert. only"
                                    ControlTipText ="Run queries only on uncertified events"
                                    LayoutCachedLeft =6480
                                    LayoutCachedTop =120
                                    LayoutCachedWidth =7530
                                    LayoutCachedHeight =390
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =7740
                            Top =150
                            OptionValue =1
                            BorderColor =0
                            Name ="optBoth"

                            LayoutCachedLeft =7740
                            LayoutCachedTop =150
                            LayoutCachedWidth =8000
                            LayoutCachedHeight =390
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =7980
                                    Top =120
                                    Width =480
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labBoth"
                                    Caption ="Both"
                                    LayoutCachedLeft =7980
                                    LayoutCachedTop =120
                                    LayoutCachedWidth =8460
                                    LayoutCachedHeight =390
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =8700
                            Top =150
                            OptionValue =2
                            BorderColor =0
                            Name ="optCertOnly"

                            LayoutCachedLeft =8700
                            LayoutCachedTop =150
                            LayoutCachedWidth =8960
                            LayoutCachedHeight =390
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =8940
                                    Top =120
                                    Width =870
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labCertOnly"
                                    Caption ="Cert. only"
                                    ControlTipText ="Run queries only on certified events"
                                    LayoutCachedLeft =8940
                                    LayoutCachedTop =120
                                    LayoutCachedWidth =9810
                                    LayoutCachedHeight =390
                                End
                            End
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2820
                    Top =120
                    Width =1620
                    TabIndex =3
                    BackColor =16777215
                    BorderColor =0
                    ForeColor =0
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cmbTimeframe"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Forms]![frm_Switchboard]![cTimeframe] AS Timeframe FROM tbl_QA_Results  "
                        "UNION SELECT tbl_QA_Results.Time_frame FROM tbl_QA_Results GROUP BY tbl_QA_Resul"
                        "ts.Time_frame ORDER BY Timeframe DESC;"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =0
                            OldBorderStyle =0
                            OverlapFlags =85
                            TextAlign =0
                            Left =180
                            Top =120
                            Width =2520
                            Height =255
                            FontWeight =400
                            BackColor =13025979
                            BorderColor =0
                            ForeColor =0
                            Name ="labTime_frame"
                            Caption ="Time frame of data being certified:"
                        End
                    End
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
' FORM NAME:    frm_QA_Tool
' Description:  Standard form for data quality review and validation
' Data source:  tbl_QA_Results
' Data access:  edit only, no deletions; opens to allow additions until a query is
'               selected, at which time additions are disallowed (see code in the subform)
' Pages:        pgResults, pgQueryViews, pgDataTables
' Functions:    fxnUpdateQAResults, fxnFilterRecords, fxnSetQueryFlag
' References:   fxnChangeDelimiter, fxnSaveFile, fxnSwitchboardIsOpen, fxnTableExists
' Source/date:  John R. Boetsch, Jan 2006
' Adapted/date: Bonnie L. Campbell, June 3, 2014
' Revisions:    JRB, May 16, 2006 - updated to use a subform for results, added conditional
'                   formatting and sort capability, and improved documentation
'               JRB, June 20, 2006 - added a button on pgResults to open the selected record
'                   in the data entry forms to maximize quality control during record fixes
'               JRB, 8/2/2006 - added additional error trapping to cmdOpenRecord
'               JRB, 10/5/2006 - fixed a problem with the refresh button giving a copy/save
'                   error message by saving the current record and turning off the form filter;
'                   added timeframe to fxnUpdateQAResults, and updated to save record before
'                   running the qa report
'               JRB, 11/14/2007 - revised the description and code in fxnUpdateQAResults
'               JRB, 12/17/2007 - added selTable_Enter to restore table pick list functionality
'                   regardless of back end in Access or SQL Server; added PageTabs change
'                   code to update and bookmark the last-selected subform record upon
'                   moving back to the first page; added code to handle multiple possible
'                   data time frames by adding an unbound ctl and linking the subform to this;
'                   added code to the results set report to also filter on data time frame;
'                   also added code to allow the user to flag records using the Is_done field
'               JRB, May 2008 - updated documentation
'               JRB, 6/18/2008 - updated Form_Open to check switchboard and enable/disable
'                   functionality based on application mode
'               JRB, 7/1/2008 - updated by adding blnRunQueries; added filter capability for
'                   Is_done and query type; added fxnFilterRecords
'               JRB, 9/17/2008 - added ref to frm_Progress_Meter (progress meter popup) in
'                   fxnUpdateQAResults
'               JRB, 9/19/2008 - added optgScope; changed txtTime_frame to cmbTimeframe;
'                   updated fxnUpdateQAResults to reflect both changes; updated call to
'                   rpt_QA_Results
'               JRB, 11/21/2008 - added txtEditQuery and fxnSetQueryFlag; updated to lock
'                   subQueryResults except when the query is named in a way that indicates
'                   its results are editable; updated cmdOpenRecord; updated cmdViewReport;
'                   added error traps to selObject and cmdDesignView; fixed a bug with opening
'                   the report and changing the filter values
'               JRB, 1/13/2009 - added save record to PageTabs_Change (copy/edit error)
'               JRB, 2/23/2009 - added cmdOpenBrowser; fixed a bug in selObject_AfterUpdate and
'                   updated fxnUpdateQAResults
'               JRB, 3/27/2009 - added cmdExport to allow quick results export to Excel
'               JRB, 5/1/2009 - updated cmdOpenBrowser to turn browser filters off by default;
'                   updated cmdExport_Click to default to current application path
'               JRB, 5/22/2009 - updated fxnFilterRecords
'               JRB, 6/10/2009 - updated cmdViewReport, cmdExport, fxnUpdateQAResults
'               JRB, 7/9/2009 - updated selTable to rely on tsys_Link_Tables, if present
'               JRB, 11/3/2009 - added cmdAutoFix and fxnEnableAutoFix
'               JRB, 2/8/2010 - updated fxnSetQueryFlag
'               JRB, 6/6/2011 - fixed a minor glitch by adding a call to fxnFilterRecords
'                   within PageTabs_Change (the front page filters were being ignored)
'               JRB, 1/31/2013 - resized panes, added cmdCloseup; set to use login rather than
'                   rely on session default user for Form_Dirty
'               --------------------------------------------------------------------------------------
'               BLC, 6/3/2014 - Adapted for NCPN WQ Utilities tool
'               BLC, 6/16/2014 - Updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 6/16/2014 - Modified to pull queries from tsys_Db_Templates
'               BLC, 8/22/2014 - Shifted blnRunQueries to mod_User & extended to project scope
'                    since used in setUserAccess (Dim -> Public), shifted fxnUpdateQAResults to mod_QA &
'                    renamed UpdateQAResults
' =================================

' ---------------------------------
' SUB:     Form_Open
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/16/2014 - Updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
'               BLC, 8/22/2014 - Shifted user access level dictated field settings to setUserAccess
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Close the form if the switchboard is not open
    If SwitchboardIsOpen = False Then
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this form to function properly.", , "Cannot open the form ..."
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

    'set default app mode & initialize controls
    setUserAccess Me
    
    ' Initialize UI
    With Me
        ' Set form time frame to global time frame
        .cmbTimeframe = TempVars.item("Timeframe")
        
        .cmbDoneFilter = "False"
        .togFilterByDone = True
    End With
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_Load
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub Form_Load()
    On Error GoTo Err_Handler

    ' Requery the results subform to reflect updates if the user chose to run upon opening
    If blnRunQueries Then Me.subResults.Requery
    ' Turn off the form filter and move to a blank record so that no query record is visible
    Me.Filter = ""
    DoCmd.GoToRecord , , acNewRec

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2105   ' Someone saved the form as not allowing new records
        MsgBox "The form has been saved in a manner that does not permit new" & _
            vbCrLf & "records to be added. Contact the database administrator.", _
            vbOKOnly, "Form saved in wrong mode (QA Tool Load Error)"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbTimeframe_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/16/2014 - Updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
' ---------------------------------
Private Sub cmbTimeframe_AfterUpdate()
    On Error GoTo Err_Handler

    If Me.cmbTimeframe <> TempVars.item("Timeframe") Then
        Me.cmdRefresh.Enabled = False
        Me.optgMode.Enabled = False
    Else
        Select Case TempVars.item("UserAccessLevel")
          Case "admin", "power user"
            Me.cmdRefresh.Enabled = True
            Me.optgMode.Enabled = True
          Case "data entry"
            Me.cmdRefresh.Enabled = True
            Me.optgMode.Enabled = False
          Case Else
            ' leave them as is
        End Select
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          PageTabs_Change
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 8/22/2014 - updated UpdateQAResults function name
' ---------------------------------
Private Sub PageTabs_Change()
    On Error GoTo Err_Handler

    Dim rst As DAO.Recordset
    Dim strCriteria As String
    Dim varReturn As Variant

    ' Bail out if the refresh button is disabled (app mode or if selected timeframe <>
    '   db timeframe)
    If Me.cmdRefresh.Enabled = False Then GoTo Exit_Procedure

    ' If moving to the first page, and if a specific query record has been selected
    '   move the subform bookmark to the currently-selected record
    If Me.PageTabs = 0 And IsNull(Me.selObject) = False Then
        ' Save the current record, reset the form filter and query selector, reset the form
        '   to allow additions, and move to a blank record
        If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

        ' Run the function to update the current QA query record
        varReturn = UpdateQAResults(False, Me.selObject)
        Me.Requery
        fxnFilterRecords
        strCriteria = "[Query_name] = """ & Me.selObject.Value & _
            """ AND [Time_frame] = """ & Me.cmbTimeframe & _
            """ AND [Data_scope] = " & Me.optgScope

        Set rst = Me.subResults.Form.RecordsetClone
        rst.FindFirst strCriteria
        If rst.NoMatch Then
            'MsgBox "No entry found.", vbInformation
        Else
            Me.subResults.Form.Bookmark = rst.Bookmark
        End If
    ElseIf Me.PageTabs = 1 And IsNull(Me.selObject) = False Then
        ' Call the function to update the query flag
        fxnSetQueryFlag
        fxnEnableAutoFix
    End If

Exit_Procedure:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          optgMode_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub optgMode_AfterUpdate()
    On Error GoTo Err_Handler

    ' Change the subform data mode depending on the user choice
    If Me.optgMode = 0 Then
    ' View mode
        Me.subQueryResults.Locked = True
        Me.txtUser.Locked = True
        Me.txtQueryDesc.Locked = True
        Me.txtRemedy.Locked = True
        Me.subDataTables.Locked = True
        Me.Detail.backcolor = 13025979 ' steel blue (default)
    Else
    ' Edit mode
        ' Unlock the subform if an editable query
        If Me.txtEditQuery = "OK" Then Me.subQueryResults.Locked = False _
            Else Me.subQueryResults.Locked = True
        Me.txtUser.Locked = False
        Me.txtQueryDesc.Locked = False
        Me.txtRemedy.Locked = False
        Me.subDataTables.Locked = False
        Me.Detail.backcolor = 12574431 ' haystack
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbTypeFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbTypeFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByType = Not IsNull(Me.cmbTypeFilter)
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          togFilterByType_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByType_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbTypeFilter) = False Then fxnFilterRecords Else Me.togFilterByType = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbDoneFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbDoneFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByDone = Not IsNull(Me.cmbDoneFilter)
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          togFilterByDone_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByDone_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbDoneFilter) = False Then fxnFilterRecords Else Me.togFilterByDone = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          Form_Dirty
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub Form_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Note: this event is ignored on inserting a new record if BeforeInsert code exists

    ' Bail out if the refresh button is disabled (app mode or if selected timeframe <>
    '   db timeframe)
    If Me.cmdRefresh.Enabled = False Then GoTo Exit_Procedure

    ' Bail out if no object record is selected - keeps from adding bogus new records
    If IsNull(Me.selObject) Then
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

    ' Once a user starts to make edits in the record, update the user field
    '   on the results summary page
    If SwitchboardIsOpen Then Me.txtUser = Environ("Username")
    Me.txtRemedy_date = Now()

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdClose_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' PAGE NAME:    QA Results Summary Page (pgResults)
' Description:  shows an overview of validation query results
' Unbound ctls: none
' Subforms:     subResults - subform for showing the results summaries
' =================================

' ---------------------------------
' SUB:          cmdRefresh_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - initial version
'               BLC, 8/25/2014 - updated UpdateQAResults function name (dropped fxn prefix)
' ---------------------------------
Private Sub cmdRefresh_Click()
    On Error GoTo Err_Handler

    ' Save the current record, reset the form filter and query selector, reset the form
    '   to allow additions, and move to a blank record
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    Me.Filter = ""
    Me.FilterOn = False
    Me.selObject = Null
    Me.subQueryResults.SourceObject = ""
    Me.AllowAdditions = True
    DoCmd.GoToRecord , , acNewRec

    ' Set the form to view mode and call the event procedure for the form mode ctl
    Me.optgMode = 0
    optgMode_AfterUpdate
    Me.Repaint

    ' Refresh the validation query results (filtering requeries the subform)
    UpdateQAResults
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdViewReport_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdViewReport_Click()
    On Error GoTo Err_Handler

    ' Generate the QA report
    Dim strRptName As String
    Dim strMsg As String
    Dim strFilter As String
    Dim strTimeframe As String
    Dim strScope As String
    Dim strInitFile As String
    Dim strSaveFile As String
    Dim varResponse As VbMsgBoxResult

    strRptName = "rpt_QA_Results"

    strMsg = "This will open the quality assurance report ..." & vbCrLf & vbCrLf & _
        "Would you like to limit report results to " & Me.cmbTimeframe & "?"
    varResponse = MsgBox(strMsg, vbYesNoCancel, "Quality assurance report")

    Select Case varResponse
      Case vbCancel
        GoTo Exit_Procedure
      Case vbYes
        strTimeframe = Me.cmbTimeframe
        strFilter = "[Time_frame]=""" & strTimeframe & """"
      Case Else
        strTimeframe = Trim(InputBox("Enter the time frame to filter by" & vbCrLf & _
            "(or leave blank to show all):", "Filter by data time frame", _
            Me.cmbTimeframe))
        If strTimeframe <> "" Then
            strFilter = "[Time_frame]=""" & strTimeframe & """"
        Else
            strFilter = ""
        End If
    End Select

    ' Save the current record so that all changes are reflected in the report
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    Select Case Me.optgScope
      Case 0
        strScope = "Uncertifed event data only"
      Case 1
        strScope = "Both certified and uncertified events"
      Case 2
        strScope = "Certified event data only"
    End Select

    If MsgBox("Would you like to filter by the current data scope?" & _
        vbCrLf & vbCrLf & "   " & strScope, vbYesNo, "Filter by data scope?") = vbYes Then
        If strFilter <> "" Then strFilter = strFilter & " AND "
        strFilter = strFilter & "[Data_scope]=" & Me.optgScope
    End If

    ' Open the formatted report output, filtering on time frame
    DoCmd.OpenReport "rpt_QA_Results", acViewPreview, , strFilter
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

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' =================================
' PAGE NAME:    Query Results Page (pgQueryViews)
' Description:  shows records returned by individual QA queries, provides the
'               user the opportunity to fix these
' Unbound ctls: selObject - combo box for selecting the query object by name
' Subforms:     subQueryResults - subform showing results of the selected query
' =================================

' ---------------------------------
' SUB:          selObject_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 8/22/2014 - updated UpdateQAResults function name
' ---------------------------------
Private Sub selObject_AfterUpdate()
    On Error GoTo Err_Handler

    Dim strCriteria As String
    Dim varReturn As Variant

    ' Bail out if the refresh button is disabled (app mode or if selected timeframe <>
    '   db timeframe)
    If Me.cmdRefresh.Enabled = False Then GoTo Exit_Procedure

    ' Exit if no query selected
    If IsNull(Me.selObject) Then
        MsgBox "Please pick from the list", vbOKOnly, "No Query Selected"
        Me.AllowAdditions = True
        DoCmd.GoToRecord , , acNewRec
        Me.txtEditQuery = ""
        Me.txtEditQuery.forecolor = 0          'black
        Me.txtEditQuery.backcolor = 8454143    'yellow
        GoTo Exit_Procedure
    End If
    
    ' Bind the subform to the selected query
    Me.subQueryResults.SourceObject = "Query." & Me.selObject.Value
    ' Build the filter string and see if a record already exists
    strCriteria = "[Query_name] = """ & Me.selObject.Value & _
        """ AND [Time_frame] = """ & Me.cmbTimeframe & _
        """ AND [Data_scope] = " & Me.optgScope
    If DCount("*", "tbl_QA_Results", strCriteria) = 0 Then
        ' Run the function to update the current QA query record
        varReturn = UpdateQAResults(False, Me.selObject, True)
    End If
    ' Set the form to the selected record
    Me.Form.Filter = strCriteria
    Me.Form.FilterOn = True

    ' Call the function to update the query flag
    fxnSetQueryFlag
    fxnEnableAutoFix

    Dim qdf As DAO.QueryDef
    Dim qdfs As DAO.QueryDefs
    Set qdfs = DBEngine(0)(0).QueryDefs

    On Error Resume Next
    For Each qdf In qdfs
        If qdf.Name = Me.selObject.Value Then
            MsgBox ("This query returns (" & DCount("*", qdf.Name) & _
                ") records that meet the following criteria: " & _
                vbCrLf & vbCrLf & qdf.Properties("Description"))
        End If
    Next qdf

Exit_Procedure:
    On Error Resume Next
    Set qdfs = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is no longer available in the application." & _
            vbCrLf & """" & Me.selObject & """", , "Query not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdDesignView_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdDesignView_Click()
    On Error GoTo Err_Handler

    ' Open the selected query in design view after checking that a query is selected
    If IsNull(Me.selObject) = False Then _
        DoCmd.OpenQuery Me.selObject.Value, acViewDesign, acReadOnly

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is no longer available in the application." & _
            vbCrLf & """" & Me.selObject & """", , "Query not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdAutoFix_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdAutoFix_Click()
    On Error GoTo Err_Handler

    Dim ctlAutoFix As Control
    Dim varAutoFix As Variant

    varAutoFix = Null

    On Error Resume Next
    Set ctlAutoFix = Forms!frm_QA_Tool.subQueryResults!varAutoFix
    varAutoFix = ctlAutoFix.Value
    On Error GoTo Err_Handler

    If IsNull(varAutoFix) Then
        MsgBox "There are no records selected, or no query is specified to fix the results."
    ElseIf Left(varAutoFix, 1) = "t" Then
    ' Object is a table - open in the next tab
        MsgBox "Object is not labeled as a query:" & vbCrLf & vbCrLf & _
            "  " & varAutoFix, , "No action taken"
    ElseIf Left(varAutoFix, 1) = "q" Then
    ' Object is a query - open on its own
        Dim qdf As DAO.QueryDef
        Dim qdfs As DAO.QueryDefs
        Set qdfs = DBEngine(0)(0).QueryDefs
        On Error Resume Next
        For Each qdf In qdfs
            If qdf.Name = varAutoFix Then
                If MsgBox("This will open/run the following query:" & vbCrLf & vbCrLf & _
                    """" & varAutoFix & """" & vbCrLf & vbCrLf & qdf.Properties("Description"), _
                    vbOKCancel, "Open or run query ...") = vbCancel Then
                    GoTo Exit_Procedure
                End If
            End If
        Next qdf
        DoCmd.OpenQuery varAutoFix
        Me.subQueryResults.Requery
    End If

Exit_Procedure:
    On Error Resume Next
    Set ctlAutoFix = Nothing
    Set qdfs = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2427   ' No records in the subform
        ' Do nothing ...
      Case 2465   ' Needed field is not present in the record set
        MsgBox "No form is specified for fixing these results", , "Missing query field"
      Case 2467   ' No subform recordset
        MsgBox "No query result set"
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox Err.Number & ": " & Err.Description
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdOpenRecord_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdOpenRecord_Click()
    On Error GoTo Err_Handler

    ' Opens the selected subform record in the object specified in the query
    '   to make use of quality control features of the front end during edits

    Dim ctlObject As Control
    Dim ctlFilter As Control
    Dim ctlArgs As Control
    Dim varObject As Variant
    Dim varFilter As Variant
    Dim varArgs As Variant

    varObject = Null
    varFilter = Null
    varArgs = Null
    
    On Error Resume Next
    Set ctlObject = Forms!frm_QA_Tool.subQueryResults!varObject
    varObject = ctlObject.Value
    Set ctlFilter = Forms!frm_QA_Tool.subQueryResults!varFilter
    varFilter = ctlFilter.Value
    Set ctlArgs = Forms!frm_QA_Tool.subQueryResults!varArgs
    varArgs = ctlArgs.Value
    On Error GoTo Err_Handler

    If IsNull(varObject) Then
        MsgBox "There are no records selected, or no form is specified."
    ElseIf Left(varObject, 1) = "t" Then
    ' Object is a table - open in the next tab
        Me.subDataTables.SourceObject = "Table." & varObject
        Me.selTable = varObject
        Me.pgDataTables.SetFocus
    ElseIf Left(varObject, 1) = "q" Then
    ' Object is a query - open on its own
        Dim qdf As DAO.QueryDef
        Dim qdfs As DAO.QueryDefs
        Set qdfs = DBEngine(0)(0).QueryDefs
        On Error Resume Next
        For Each qdf In qdfs
            If qdf.Name = varObject Then
                If MsgBox("This will open/run the following query:" & vbCrLf & vbCrLf & _
                    """" & varObject & """" & vbCrLf & vbCrLf & qdf.Properties("Description"), _
                    vbOKCancel, "Open or run query ...") = vbCancel Then
                    GoTo Exit_Procedure
                End If
            End If
        Next qdf
        DoCmd.OpenQuery varObject
        Me.subQueryResults.Requery
    ElseIf IsNull(varFilter) Then
    ' Filter by form alone if no filter
        Select Case varObject
          Case "frm_Contacts"
            Set gvarRefContactCtl = Me.subQueryResults
          Case "fsub_Project_Taxa"
            Set gvarRefTaxonCtl = Me.subQueryResults
          Case Else
            Set gvarRefForm = Me.Form
            Set gvarRefCtl = Me.subQueryResults
        End Select
        DoCmd.OpenForm varObject, , , , , , varArgs
    Else
    ' Filter by form and filter
        Select Case varObject
          Case "frm_Contacts"
            Set gvarRefContactCtl = Me.subQueryResults
          Case "fsub_Project_Taxa"
            Set gvarRefTaxonCtl = Me.subQueryResults
          Case Else
            Set gvarRefForm = Me.Form
            Set gvarRefCtl = Me.subQueryResults
        End Select
        DoCmd.OpenForm varObject, , , varFilter, , , varArgs
    End If

Exit_Procedure:
    On Error Resume Next
    Set ctlArgs = Nothing
    Set ctlFilter = Nothing
    Set ctlObject = Nothing
    Set qdfs = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2427   ' No records in the subform
        ' Do nothing ...
      Case 2465   ' Needed field is not present in the record set
        MsgBox "No form is specified for fixing these results", , "Missing query field"
      Case 2467   ' No subform recordset
        MsgBox "No query result set"
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox Err.Number & ": " & Err.Description
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdOpenBrowser_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdOpenBrowser_Click()
    On Error GoTo Err_Handler

    Set gvarRefForm = Me.Form
    Set gvarRefCtl = Me.subQueryResults
    ' Open to a blank record - to distinguish from opening to the selected record in the subform
    DoCmd.OpenForm "frm_Data_Browser", , , , acFormAdd, , "off"

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox Err.Number & ": " & Err.Description
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdExport_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdExport_Click()
    On Error GoTo Err_Handler

    Dim strQName As String
    Dim strSaveFile As String

    ' Bail out if no query is currently selected
    If IsNull(Me.selObject) Then GoTo Exit_Procedure
    ' Requery the selected record in the recordset, and update the subform
    Me.subQueryResults.Requery
    strQName = Me.selObject
    strSaveFile = CurrentProject.Path & "\" & strQName & "_" & _
        CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".xls"
    DoCmd.OutputTo acOutputQuery, strQName, acFormatXLS, strSaveFile, True
    MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdCloseup_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdCloseup_Click()
    On Error GoTo Err_Handler

    ' Open the selected query in a new window after checking that a query is selected
    If IsNull(Me.selObject) = False Then
        If Me.txtEditQuery = "OK" Then
            DoCmd.OpenQuery Me.selObject.Value, acViewNormal, acEdit
        Else
            DoCmd.OpenQuery Me.selObject.Value, acViewNormal, acReadOnly
        End If
        DoCmd.Maximize
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.selObject & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdRequery_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmdRequery_Click()
    On Error GoTo Err_Handler

    'Dim varReturn As Variant

    ' Bail out if no query is currently selected
    If IsNull(Me.selObject) Then GoTo Exit_Procedure
    ' Requery the selected record in the recordset, and update the subform
    Me.subQueryResults.Requery
    ' Run the function to update the current QA query record - commented out because this
    '   is done upon changing page tabs
    'varReturn = fxnUpdateQAResults(False, Me.selObject)

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          txtUser_Dirty
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub txtUser_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Prompt user to confirm before allowing edits in the QA user control
    If MsgBox("Are you sure you want to change the user name?", _
        vbYesNo, "Please confirm ...") = vbNo Then
        DoCmd.CancelEvent
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          txtQueryDesc_Dirty
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub txtQueryDesc_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Prompt user to confirm before allowing edits in query definition control
    If MsgBox("Are you sure you want to change the query definition?", _
        vbYesNo, "Please confirm ...") = vbNo Then
        DoCmd.CancelEvent
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' PAGE NAME:    Data Tables Page (pgDataTables)
' Description:  allows the user to select and view the contents of individual data
'               tables to make data revisions
' Unbound ctls: selTable - combo box for selecting the table object by name
' Subforms:     subDataTables - subform showing the contents of the selected table
' =================================

' ---------------------------------
' SUB:          selTable_Enter
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub selTable_Enter()
     On Error GoTo Err_Handler

    Dim strSysTable As String

    strSysTable = "tsys_Link_Tables"     ' System table listing linked tables

    ' If the system table does not exist, replace the row source with one that doesn't use it
    If TableExists(strSysTable) = False Then
        Me.selTable.RowSource = "SELECT MSysObjects.Name " & _
            "FROM MSysObjects " & _
            "WHERE (((MSysObjects.Name) Like 'tbl_*' " & _
            "And (MSysObjects.Name)<>'tbl_QA_Results')) " & _
            "OR (((MSysObjects.Name)='tlu_Project_Crew')) " & _
            "OR (((MSysObjects.Name)='tlu_Project_Taxa')) " & _
            "OR (((MSysObjects.Name)='tlu_Park_Taxa'));"
        Me.selTable.ColumnCount = 1
        Me.selTable.ListWidth = Me.selTable.Width
        Me.selTable.Requery
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          selTable_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub selTable_AfterUpdate()
    On Error GoTo Err_Handler

    ' Once a table is selected, bind the subform to this table
    If IsNull(Me.selTable) Then
    ' If none selected ...
        Me.subDataTables.SourceObject = ""
    Else
    ' If a table is selected ...
        If TableExists(Me.selTable) Then
            Me.subDataTables.SourceObject = "Table." & Me.selTable.Value
        Else
            MsgBox "Unable to find the selected table in the database ...", , _
                "Table not found"
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' FUNCTION:     fxnFilterRecords
' Description:  Filter the records by the indicated field
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 5, 2006
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, May 2008 - made code more robust and error-proof
'               JRB, 7/1/2008 - updated by filtering on the subform rather than the form
'               JRB, 5/22/2009 - updated filter AND clauses
' ---------------------------------
Private Function fxnFilterRecords()
    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim bFilterOn As Boolean

    bFilterOn = False
    strFilter = ""

    ' Save the record (to trigger validation)
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    If Me.togFilterByType Then
        bFilterOn = True
        strFilter = strFilter & "[Query_type] = """ & Me.cmbTypeFilter & """"
    End If
    If Me.togFilterByDone Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Is_done] = " & Me.cmbDoneFilter & ""
    End If

    ' Apply the filter
    'Me.Filter = strFilter
    'Me.FilterOn = bFilterOn
    Me.subResults.Form.Filter = strFilter
    Me.subResults.Form.FilterOn = bFilterOn

    ' Make the labels bold or not depending on filter settings
    Me.labTypeFilter.fontBold = Me.togFilterByType
    Me.labDoneFilter.fontBold = Me.togFilterByDone

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 2001   ' Run time canceled event (validation error) - do nothing
        Me.togFilterByType = False
        Me.togFilterByDone = False
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnFilterRecords)"
    End Select
    Resume Exit_Procedure

End Function

' ---------------------------------
' FUNCTION:     fxnSetQueryFlag
' Description:  Updates the flag to indicate whether or not the query results are editable
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 10/7/2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, 2/8/2010 - updated flag from "X" to "_X" in of x as last letter in name
' ---------------------------------
Private Function fxnSetQueryFlag()
    On Error GoTo Err_Handler

    ' Update the visual flag to indicate whether or not the query results are editable
    '   Note: suffix of "_X" means that the query results may be edited
    If Right(Me.selObject.Value, 2) = "_X" Then
        Me.txtEditQuery = "OK"
        Me.txtEditQuery.forecolor = 16777215   'white
        Me.txtEditQuery.backcolor = 4227072    'green
        ' Unlock the subform if in edit mode
        If Me.optgMode = 1 Then Me.subQueryResults.Locked = False _
            Else Me.subQueryResults.Locked = True
    Else
        Me.txtEditQuery = "No"
        Me.txtEditQuery.forecolor = 16777215   'white
        Me.txtEditQuery.backcolor = 255        'red
        ' Lock the subform
        Me.subQueryResults.Locked = True
    End If

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnSetQueryFlag)"
    End Select
    Resume Exit_Procedure

End Function

' ---------------------------------
' FUNCTION:     fxnEnableAutoFix
' Description:  Enables or disables the control for running an action query to fix records
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 11/3/2009
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/16/2014 - Updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
' ---------------------------------
Private Function fxnEnableAutoFix()
    On Error GoTo Err_Handler

    Dim ctlAutoFix As Control

    Me.cmdAutoFix.Enabled = False

    ' The following looks for 'varAutoFix' field in the query results ...
    '   If it isn't there, it will throw a trapped error and the ctl will remain disabled
    Set ctlAutoFix = Forms!frm_QA_Tool.subQueryResults!varAutoFix

    ' If no error, the field is there ... enable the ctl if user has sufficient rights
    Select Case TempVars.item("UserAccessLevel")
      Case "admin", "power user"
        Me.cmdAutoFix.Enabled = True
    End Select

Exit_Procedure:
    On Error Resume Next
    Set ctlAutoFix = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 2465, 2467
        ' Do nothing ...
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnEnableAutoFix)"
    End Select
    Resume Exit_Procedure

End Function
