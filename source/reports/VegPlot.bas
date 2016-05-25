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
    Width =11491
    DatasheetFontHeight =11
    ItemSuffix =17
    Right =25650
    Bottom =12045
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0x32f638d6d6a9e440
    End
    Caption ="VegPlot"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006801000000000000e32c00005820000001000000 ,
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
            Height =360
            BackColor =15849926
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =840
                    Top =60
                    Width =4200
                    Height =300
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTitle"
                    Caption ="=[Report].[Caption]"
                    GridlineColor =10921638
                    LayoutCachedLeft =840
                    LayoutCachedTop =60
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =360
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
                    Left =14
                    Width =11477
                    Height =4320
                    Name ="oTPctCover"
                    SourceObject ="Report.PercentCover"
                    GridlineColor =10921638
                    FilterOnEmptyMaster =0

                    LayoutCachedLeft =14
                    LayoutCachedWidth =11491
                    LayoutCachedHeight =4320
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =14
                    Top =4680
                    Width =11477
                    Height =2520
                    TabIndex =1
                    Name ="oBPctCover"
                    SourceObject ="Report.PercentCover"
                    GridlineColor =10921638
                    FilterOnEmptyMaster =0

                    LayoutCachedLeft =14
                    LayoutCachedTop =4680
                    LayoutCachedWidth =11491
                    LayoutCachedHeight =7200
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
' References:   -
' Revisions:    BLC - 5/25/2016 - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private WithEvents oTPctCover As Report_PercentCover
Attribute oTPctCover.VB_VarHelpID = -1
Private WithEvents oBPctCover As Report_PercentCover
Attribute oBPctCover.VB_VarHelpID = -1

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------

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
            "Error encountered (#" & Err.Number & " - XX[VegPlot report])"
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
Private Sub Report_Open(cancel As Integer)
On Error GoTo Err_Handler
    
    'testing
    TempVars.Add "ParkCode", "CANY"
    
    'prepare interface
    Dim oTPctCover As New Report_PercentCover
    Dim oBPctCover As New Report_PercentCover
    
    'VegPlot
    lblTitle.Caption = "VegPlot"
    
    ' ------ Keys -------
    Me.rsubModWentworth.SourceObject = "Report.ModWentworthKey"
    
    ' ------ Top -------
    Set oTPctCover = oTPctCover.Report
    With oTPctCover
        .Park = TempVars("ParkCode")
        Select Case .Park
            Case "BLCA"
                .CoverType = "WCC"
            Case "CANY"
                .CoverType = "WCC"
            Case "DINO"
                .CoverType = "ARS"
        End Select
    
    End With
        
    ' ------ Bottom -------
    Set oBPctCover = oBPctCover.Report
    With oBPctCover
        'default
        .Visible = True
        
        .Park = TempVars("park")
        .CoverType = "URC"
        
        Select Case .Park
            Case "BLCA"
            Case "CANY"
            Case "DINO" 'DINO has only one percent cover displayed (ARS)
                .Visible = False
'               .Height = 0.1
        End Select

    End With
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[VegPlot report])"
    End Select
    Resume Exit_Handler
End Sub
