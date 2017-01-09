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
    Width =11532
    DatasheetFontHeight =11
    ItemSuffix =18
    Right =21330
    Bottom =9645
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0x32f638d6d6a9e440
    End
    Caption ="VegPlot"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a010000680100006801000068010000000000000c2d00005820000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    FitToPage =1
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
        Begin Subform
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1380
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    Top =960
                    Width =492
                    Height =290
                    FontSize =9
                    FontWeight =700
                    TopMargin =14
                    BorderColor =8355711
                    Name ="lblRiver"
                    Caption ="River:"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedTop =960
                    LayoutCachedWidth =492
                    LayoutCachedHeight =1250
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =9180
                    Top =1020
                    Width =2220
                    Height =300
                    FontSize =10
                    FontWeight =600
                    BorderColor =8355711
                    Name ="lblSamplingDate"
                    Caption ="Date: ____/____/_____"
                    GridlineColor =10921638
                    LayoutCachedLeft =9180
                    LayoutCachedTop =1020
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =1320
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =600
                    Top =996
                    Width =1740
                    Height =240
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblRiverSegments"
                    Caption ="Green    CAC    CBC   "
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =600
                    LayoutCachedTop =996
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =1236
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =180
                    Top =120
                    Width =2700
                    Height =420
                    FontSize =14
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =6250335
                    Name ="lblTitle"
                    Caption ="Veg Plot"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =540
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    Top =600
                    Width =11520
                    Height =360
                    BackColor =8355711
                    BorderColor =10921638
                    Name ="rctUnderHeader"
                    GridlineColor =10921638
                    LayoutCachedTop =600
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =960
                    BackThemeColorIndex =0
                    BackTint =50.0
                End
                Begin Label
                    Left =4620
                    Top =960
                    Width =912
                    Height =276
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblObserver"
                    Caption ="Observer:"
                    GridlineColor =10921638
                    LayoutCachedLeft =4620
                    LayoutCachedTop =960
                    LayoutCachedWidth =5532
                    LayoutCachedHeight =1236
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =6960
                    Top =960
                    Width =912
                    Height =276
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblRecorder"
                    Caption ="Recorder:"
                    GridlineColor =10921638
                    LayoutCachedLeft =6960
                    LayoutCachedTop =960
                    LayoutCachedWidth =7872
                    LayoutCachedHeight =1236
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =3600
                    Top =60
                    Width =1620
                    Height =540
                    BorderColor =8355711
                    Name ="lblEntry"
                    Caption ="Data entered by: Date entered:"
                    GridlineColor =10921638
                    LayoutCachedLeft =3600
                    LayoutCachedTop =60
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =6780
                    Top =60
                    Width =1620
                    Height =540
                    BorderColor =8355711
                    Name ="lblVerify"
                    Caption ="Data verified by: Date verified:"
                    GridlineColor =10921638
                    LayoutCachedLeft =6780
                    LayoutCachedTop =60
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =5040
                    Top =300
                    Width =1620
                    Height =300
                    BorderColor =8355711
                    Name ="lblEntryDate"
                    Caption ="____/____/_____"
                    GridlineColor =10921638
                    LayoutCachedLeft =5040
                    LayoutCachedTop =300
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =8220
                    Top =300
                    Width =1620
                    Height =300
                    BorderColor =8355711
                    Name ="lblVerifyDate"
                    Caption ="____/____/_____"
                    GridlineColor =10921638
                    LayoutCachedLeft =8220
                    LayoutCachedTop =300
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    Left =9840
                    Width =1692
                    Height =600
                    FontSize =10
                    FontWeight =700
                    LeftMargin =72
                    TopMargin =144
                    RightMargin =72
                    BackColor =14540253
                    BorderColor =8355711
                    Name ="lblPageOf"
                    Caption ="Page ____ of ____"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =9840
                    LayoutCachedWidth =11532
                    LayoutCachedHeight =600
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =60
                    Top =636
                    Width =2460
                    Height =324
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblMonitoring"
                    Caption ="NCPN Big River Monitoring"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =636
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =960
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =3
                    Left =5100
                    Top =624
                    Width =6300
                    Height =324
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblProtocolVersion"
                    Caption ="Big River Monitoring - SOP #6 - Version 3.00 - Jan 2016"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =5100
                    LayoutCachedTop =624
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =948
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =2400
                    Top =960
                    Width =1020
                    Height =290
                    FontSize =9
                    FontWeight =700
                    TopMargin =14
                    BorderColor =8355711
                    Name ="lblSite"
                    Caption ="Sentinel Site:"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =2400
                    LayoutCachedTop =960
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =1250
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =3420
                    Top =960
                    Width =1068
                    Height =324
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblSiteName"
                    Caption ="_________"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =3420
                    LayoutCachedTop =960
                    LayoutCachedWidth =4488
                    LayoutCachedHeight =1284
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin PageHeader
            Height =2400
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Width =6599
                    Name ="rsubModWentworth"
                    SourceObject ="Report.ModWentworthKey"
                    GridlineColor =10921638
                    FilterOnEmptyMaster =0

                    LayoutCachedWidth =6599
                    LayoutCachedHeight =1440
                    ShowPageHeaderAndPageFooter =255
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =7830
                    Top =1432
                    Width =3615
                    Height =720
                    TabIndex =1
                    Name ="rsubPlotDimensionsKey"
                    SourceObject ="Report.PlotDimensionsKey"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =7830
                    LayoutCachedTop =1432
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =2152
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =8280
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Width =11477
                    Height =4320
                    Name ="TPctCover"
                    SourceObject ="Report.PercentCover"
                    GridlineColor =10921638
                    FilterOnEmptyMaster =0

                    LayoutCachedWidth =11477
                    LayoutCachedHeight =4320
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Top =4740
                    Width =11477
                    Height =2520
                    TabIndex =1
                    Name ="BPctCover"
                    SourceObject ="Report.PercentCover"
                    GridlineColor =10921638
                    FilterOnEmptyMaster =0

                    LayoutCachedTop =4740
                    LayoutCachedWidth =11477
                    LayoutCachedHeight =7260
                End
            End
        End
        Begin PageFooter
            Height =360
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    Left =5400
                    Top =60
                    Width =2460
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPages"
                    Caption ="=[Page] & \" | \" & [Pages]"
                    GridlineColor =10921638
                    LayoutCachedLeft =5400
                    LayoutCachedTop =60
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =300
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =360
            Name ="ReportFooter"
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
' Report:       VegPlot
' Level:        Application report
' Version:      1.00
'
' Description:  Vegetation plot report object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 25, 2016
' References:
'   Michael Lester, April 1, 2005
'   http://forums.aspfree.com/microsoft-access-help-18/changing-record-source-subreports-vba-53031.html
' Revisions:    BLC - 5/25/2016 - 1.00 - initial version
' =================================

'---------------------
' Global Declarations
'---------------------
'Public gSubReportCount As Integer --> in App settings

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Park As String
'Private WithEvents oTPctCover As Report_PercentCover
'Private WithEvents oBPctCover As Report_PercentCover

'---------------------
' Event Declarations
'---------------------
Public Event InvalidRow(value As Integer)
Public Event InvalidNumRows(value As Integer)

'---------------------
' Properties
'---------------------
Public Property Let Park(value As String)
    If Len(value) = 4 Then
        m_Park = value
    End If
End Property

Public Property Get Park() As String
    Park = m_Park
End Property

'---------------------
' Events
'---------------------

' ---------------------------------
' Sub:          XX
' Description:  XX event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 10, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/10/2015 - initial version
' ---------------------------------
Private Sub XX()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - XX[VegPlot Report])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Report_Open
' Description:  report opening actions
' Assumptions:  Two separate percent cover subreport objects are present to
'               handle WCC/ARS and URC
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 25, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/25/2016 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)
On Error GoTo Err_Handler
        
    Dim arySegments() As Variant, aryProtocol() As Variant
    Dim sopdata As Variant
    Dim i As Integer
    Dim ary() As String, strSegments As String
    Dim strSQL As String, strFamily As String, strSpecies As String
        
    'defaults
    strSegments = ""
    i = 0
        
    'set park
    If Len(Nz(OpenArgs, "")) > 0 Or IsNull(OpenArgs) Then
        Me.Park = Nz(TempVars("ParkCode"), "")
    Else
        ary = Split(OpenArgs, "|")
        Me.Park = UCase(ary(0))
    End If
        
    'page headers
    Me.lblTitle.Caption = UCase(Me.Park) & " Veg Walk"
    
    'protocol version
    aryProtocol = GetProtocolVersion
    Set sopdata = GetSOPMetadata("VegPlot") '0-code, 1-SOP#, 2-Version, 3-Effective Date
    i = CInt(sopdata(1))

    lblMonitoring.Caption = "NCPN " & aryProtocol(0, 0)
    lblProtocolVersion.Caption = aryProtocol(0, 0) & " - " & "SOP #" & i & " - Version " & Format(sopdata(2), "0.00") & " - " & Format(sopdata(3), "mmm yyyy")

    'set river segment(s)
    arySegments = GetRiverSegments(Me.Park)
    For i = 0 To UBound(arySegments, 2)
        strSegments = strSegments & arySegments(0, i) & Space(2)
    Next
    strSegments = Left(strSegments, Len(strSegments) - 1)
    
    Me.lblRiverSegments.Caption = strSegments
        
    'prepare interface
'    Dim oTPctCover As Report_PercentCover
'    Dim oBPctCover As Report_PercentCoverCOPY
    Dim oTPctCover As Report_PercentCover
    Dim oBPctCover As Report_PercentCover
    
    'VegPlot
    lblTitle.Caption = Nz(TempVars("ParkCode"), "") & Space(2) & "VegPlot"
    
    ' ------ Keys -------
    Me.rsubModWentworth.SourceObject = "Report.ModWentworthKey"
    
'    ' ------ Top -------
'    'Set oTPctCover = TPctCover
'    'TPctCover.SourceObject = "PercentCover"
'    Set oTPctCover = TPctCover '.Report '.Report
'    With oTPctCover
'
'        .Park = TempVars("ParkCode")
'        Select Case .Park
'            Case "BLCA"
'                .SetCoverType "WCC"
'            Case "CANY"
'                .SetCoverType "WCC"
'            Case "DINO"
'                .SetCoverType "ARS"
'        End Select
'        Debug.Print .Park & " - " & .CoverType & vbCrLf
'    End With
'
'    ' ------ Bottom -------
'    Set oBPctCover = BPctCover.Report
'    With oBPctCover
'        'default
'        .visible = True
'
'        .Park = TempVars("ParkCode")
'        .SetCoverType "URC"
'
'        Select Case .Park
'            Case "BLCA"
'            Case "CANY"
'            Case "DINO" 'DINO has only one percent cover displayed (ARS)
'                .visible = False
'        End Select
'        Debug.Print .Park & " - " & .CoverType & vbCrLf
'    End With
    
    'hide modal Main form
    Forms("Main").Visible = False
    
Exit_Handler:
    gSubReportCount = 0
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[VegPlot Report])"
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
' Source/date:  Bonnie Campbell, June 2, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/2/2016 - initial version
' ---------------------------------
Private Sub Report_Close()
On Error GoTo Err_Handler

    'unhide modal Main form
    Forms("Main").Visible = True

Exit_Handler:
    gSubReportCount = 0
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Close[VegPlot Report])"
    End Select
    Resume Exit_Handler
End Sub
