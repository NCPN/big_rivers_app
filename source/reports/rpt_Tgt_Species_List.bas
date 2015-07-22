Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11400
    DatasheetFontHeight =11
    ItemSuffix =51
    Left =105
    Top =1485
    Right =12120
    Bottom =10305
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x93926d5acb96e440
    End
    RecordSource ="qry_Park_Tgt_Species_Lists"
    Caption ="INVASIVE LIST"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006801000000000000882c0000a201000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =255
    FitToPage =1
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    RibbonName ="Export"
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
        Begin BreakLevel
            ControlSource ="Family"
        End
        Begin BreakLevel
            ControlSource ="utah_species"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =960
            BackColor =15849926
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =7140
                    Height =528
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblReportHdr"
                    Caption ="INVASIVES TARGET LIST"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =588
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9000
                    Width =2340
                    Height =540
                    ColumnOrder =1
                    FontSize =20
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxParkYear"
                    ControlSource ="=IIf([OpenArgs]=\"preview\",[TempVars]![park] & \" - \" & [TempVars]![TgtYear],["
                        "Park] & \" - \" & [TgtYear])"
                    StatusBarText ="Park and year for list"
                    GridlineColor =10921638

                    LayoutCachedLeft =9000
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =540
                    ForeTint =50.0
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1500
                    Top =600
                    Width =4140
                    Height =315
                    ColumnOrder =0
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModified"
                    ControlSource ="=Format([Last_Modified],\"mmmm d\"\", \"\"yyyy h:nn ampm\")"
                    Format ="General Number"
                    GridlineColor =10921638

                    LayoutCachedLeft =1500
                    LayoutCachedTop =600
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =915
                End
                Begin Label
                    Left =60
                    Top =612
                    Width =1320
                    Height =300
                    BorderColor =8355711
                    ForeColor =4210752
                    Name ="lblLastModified"
                    Caption ="Last Modified:"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =612
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =912
                    ForeTint =75.0
                End
            End
        End
        Begin PageHeader
            Height =1335
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    OldBorderStyle =0
                    Width =11400
                    Height =480
                    BackColor =15849926
                    BorderColor =10921638
                    Name ="rectPageHdr"
                    GridlineColor =10921638
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =480
                    BackThemeColorIndex =2
                    BackTint =20.0
                End
                Begin Label
                    TextAlign =2
                    Left =2160
                    Top =960
                    Width =1800
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNameUT"
                    Caption ="UT"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2160
                    LayoutCachedTop =960
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =4200
                    Top =960
                    Width =1980
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNameCO"
                    Caption ="CO"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4200
                    LayoutCachedTop =960
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =6360
                    Top =960
                    Width =1380
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPlantCode"
                    Caption ="Plant Code"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6360
                    LayoutCachedTop =960
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =240
                    Top =960
                    Width =1800
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFamily"
                    Caption ="Family"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =960
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =8100
                    Top =960
                    Width =1680
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCommonName"
                    Caption ="Common Name"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =8100
                    LayoutCachedTop =960
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =10080
                    Top =960
                    Width =1200
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPriority"
                    Caption ="Priority"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10080
                    LayoutCachedTop =960
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =2160
                    Top =600
                    Width =3720
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNames"
                    Caption ="Species Names"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2160
                    LayoutCachedTop =600
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =900
                End
                Begin Line
                    Left =2160
                    Top =924
                    Width =3720
                    Name ="lnSpecies"
                    GridlineColor =10921638
                    LayoutCachedLeft =2160
                    LayoutCachedTop =924
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =924
                End
                Begin Line
                    BorderWidth =2
                    Left =180
                    Top =1320
                    Width =11100
                    Name ="lnHeader"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =1320
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =1320
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1500
                    Top =60
                    Width =3300
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDate"
                    ControlSource ="=Format(Now(),\"mmmm d\"\", \"\"yyyy h:nn ampm\")"
                    Format ="Medium Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =1500
                    LayoutCachedTop =60
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7200
                    Top =60
                    Width =4080
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPage"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =7200
                    LayoutCachedTop =60
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4320
                    Top =60
                    Width =2880
                    Height =312
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxListName"
                    ControlSource ="=IIf([Page]>1,\"Invasives List for \" & [tbxParkYear],\"\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =4320
                    LayoutCachedTop =60
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =372
                End
                Begin Label
                    Left =60
                    Top =60
                    Width =1320
                    Height =300
                    BorderColor =8355711
                    ForeColor =4210752
                    Name ="Label49"
                    Caption ="Printed:"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =360
                    ForeTint =75.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =418
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Width =11400
                    Height =418
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDetail"
                    ConditionalFormat = Begin
                        0x0100000028010000020000000100000000000000000000001e00000001000000 ,
                        0x00000000ccff660001000000000000001f000000630000000100000000000000 ,
                        0xffff990000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078005000720069006f0072006900740079005d003d0022005400 ,
                        0x720061006e00730065006300740020004f006e006c0079002200000000002800 ,
                        0x4e006f0074002000490073004e0075006d00650072006900630028005b007400 ,
                        0x620078005000720069006f0072006900740079005d0029002900200041006e00 ,
                        0x6400200028005b007400620078005000720069006f0072006900740079005d00 ,
                        0x3c003e0022005400720061006e00730065006300740020004f006e006c007900 ,
                        0x2200290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedWidth =11400
                    LayoutCachedHeight =418
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000ccff66001d0000005b00 ,
                        0x7400620078005000720069006f0072006900740079005d003d00220054007200 ,
                        0x61006e00730065006300740020004f006e006c00790022000000000000000000 ,
                        0x0000000000000000000000000001000000000000000100000000000000ffff99 ,
                        0x004300000028004e006f0074002000490073004e0075006d0065007200690063 ,
                        0x0028005b007400620078005000720069006f0072006900740079005d00290029 ,
                        0x00200041006e006400200028005b007400620078005000720069006f00720069 ,
                        0x00740079005d003c003e0022005400720061006e00730065006300740020004f ,
                        0x006e006c00790022002900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            Width =705
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label34"
                            Caption ="Text33"
                            GridlineColor =10921638
                            LayoutCachedWidth =705
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9120
                    Top =60
                    Width =1140
                    Height =300
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPri1RunSum"
                    ControlSource ="=CDbl(Nz(IIf(Switch([Transect_Only]=1,0,Len([Tgt_Area])>0,0,[Priority]>-1,[Prior"
                        "ity])=1,1,0),0))"
                    GridlineColor =10921638

                    LayoutCachedLeft =9120
                    LayoutCachedTop =60
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =1800
                    Height =312
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFamily"
                    ControlSource ="Family"
                    StatusBarText ="Target Species name (ITIS species name from tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10080
                    Top =60
                    Width =1140
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPriority"
                    ControlSource ="=Switch([Transect_Only]=1,\"Transect Only\",Len([Tgt_Area])>0,[Tgt_Area],[Priori"
                        "ty]>-1,[Priority])"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10080
                    LayoutCachedTop =60
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8100
                    Top =60
                    Width =1680
                    Height =312
                    ColumnWidth =2400
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCommon"
                    ControlSource ="Master_Common_Name"
                    StatusBarText ="FK to plant master code (tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =8100
                    LayoutCachedTop =60
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6360
                    Top =60
                    Width =1380
                    Height =312
                    ColumnWidth =2655
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Species_Name"
                    ControlSource ="LU_code"
                    StatusBarText ="Target Species name (ITIS species name from tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =6360
                    LayoutCachedTop =60
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4260
                    Top =60
                    Width =1980
                    Height =312
                    ColumnWidth =1170
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbl_Target_Species.Park_Code"
                    ControlSource ="Co_Species"
                    StatusBarText ="Standard park code (CANY, FOBU, etc.)"
                    EventProcPrefix ="tbl_Target_Species_Park_Code"
                    GridlineColor =10921638

                    LayoutCachedLeft =4260
                    LayoutCachedTop =60
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2220
                    Top =60
                    Width =1800
                    Height =312
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbl_Target_Species.Target_Year"
                    ControlSource ="utah_species"
                    StatusBarText ="Year (4-digit)"
                    EventProcPrefix ="tbl_Target_Species_Target_Year"
                    GridlineColor =10921638

                    LayoutCachedLeft =2220
                    LayoutCachedTop =60
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =372
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
            Height =960
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10020
                    Top =120
                    Width =1140
                    Height =312
                    FontSize =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumPriority1"
                    ControlSource ="=[tbxPri1RunSum]"
                    GridlineColor =10921638

                    LayoutCachedLeft =10020
                    LayoutCachedTop =120
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =432
                End
                Begin Label
                    TextAlign =3
                    Left =7200
                    Top =120
                    Width =2700
                    Height =324
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTotalNum"
                    Caption ="Total # Priority 1 Species ="
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7200
                    LayoutCachedTop =120
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =444
                End
                Begin Line
                    BorderWidth =2
                    Left =180
                    Width =11100
                    Name ="lnPageFooter"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedWidth =11280
                End
                Begin TextBox
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10020
                    Top =540
                    Width =1140
                    Height =312
                    FontSize =12
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumSpecies"
                    ControlSource ="=CDbl(Count(*))"
                    GridlineColor =10921638

                    LayoutCachedLeft =10020
                    LayoutCachedTop =540
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =852
                End
                Begin Label
                    TextAlign =3
                    Left =7200
                    Top =540
                    Width =2700
                    Height =324
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label36"
                    Caption ="Total # Species ="
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7200
                    LayoutCachedTop =540
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =864
                End
                Begin Label
                    Visible = NotDefault
                    Left =1080
                    Top =120
                    Width =1350
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPunchMargin"
                    Caption ="|<< .75margin"
                    GridlineColor =10921638
                    LayoutCachedLeft =1080
                    LayoutCachedTop =120
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =420
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
' MODULE:       rpt_Tgt_Species_List
' Description:  Target species list crew report functions and routines
'
' Source/date:  Bonnie Campbell, 3/5/2015
' Revisions:    BLC - 3/5/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Report_Open
' Description:  Actions for when reports open
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:
'   http://stackoverflow.com/questions/11477297/giving-an-alias-to-a-subquery-containing-a-join-in-access
'   Bob Larsen, January 28, 2012
'   https://social.msdn.microsoft.com/Forums/office/en-US/3e126484-112f-4854-a5c0-2e9ef48e02bc/how-to-change-recordsource-for-a-report-with-vba?forum=accessdev
'       set recordset to passed in SQL via OpenArgs
'       If Me.OpenArgs <> vbNullString Then
'       Me.Recordset = Me.OpenArgs
'   dyDMA, Sept 8, 2008
'   http://www.utteraccess.com/forum/Run-time-error-32585-t1710296.html
'       Me.Recordset = qdf.OpenRecordset()
'       ==> Run-time Error 32585: This feature is only available in an ADP
'       ==> Only Access ADP's can use this method (assign report recordset @ run-time)
'       ==> Not available for *.mdb or *.accdb's
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/1/2015 - initial version
'   BLC - 6/3/2015 - added check for "preview" openarg to handle list previews
'   BLC - 6/11/2015 - added Last_Modified to qry_Tgt_Species_List_Preview to handle Last_Modified date
'                     in header vs. setting using IIF(Me.OpenArgs = "preview"... in vba or report tbx design
'                     which failed to display data
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)

On Error GoTo Err_Handler
   
    If Me.OpenArgs <> vbNullString Then
    
        Select Case Me.OpenArgs
            Case "preview"
                Dim qdf As DAO.QueryDef
                
                'delete table if exists
                If TableExists("temp_List_Preview") Then
                    DoCmd.SetWarnings False
                    DoCmd.DeleteObject acTable, "temp_List_Preview"
                    DoCmd.SetWarnings True
                End If
                
                'qry_Park_Tgt_Species_List_Preview --> creates table temp_List_Preview
                Set qdf = CurrentDb.QueryDefs("qry_Tgt_Species_List_Preview")
                qdf.Parameters("park") = TempVars("park")
                qdf.Parameters("TgtYear") = CInt(TempVars("TgtYear"))
                
                qdf.Execute
                
                'set report recordset
                Me.RecordSource = "temp_List_Preview"
                          
                'set headers
                '=IIf([OpenArgs]="preview",[TempVars]![park] & " - " & [TempVars]![TgtYear],[Park] & " - " & [TgtYear])
                '=IIf([OpenArgs] = "preview", "-", Format([Last_Modified], "mmmm d"", ""yyyy h:nn ampm"))
                'tbxLastModified.ControlSource = IIf(Me.OpenArgs = "preview", "-", Format([Last_Modified], "mmmm d"", ""yyyy h:nn ampm"))
                tbxListName.ControlSource = IIf([Page] > 1, "Invasives List for " & TempVars("park") & "-" & TempVars("TgtYear"), "")
                lblReportHdr.Caption = "INVASIVES TARGET LIST PREVIEW"
        End Select
        
    End If
        
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[Report_rptTgtSpeciesList])"
    End Select
    Resume Exit_Sub
End Sub
