Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    ScrollBars =0
    BorderStyle =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =1
    GridX =24
    GridY =24
    Width =6480
    DatasheetFontHeight =11
    ItemSuffix =123
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0x662cc1c4d5dbe440
    End
    RecordSource ="PARAMETERS yr Long; SELECT Label AS ClassSize, Code, DiameterRange_mm, ActiveYea"
        "r, RetireYear, CategoryOrder, KeyOrder FROM ModWentworthScale WHERE ([yr] = [Act"
        "iveYear] AND ([yr]<[RetireYear] OR [RetireYear] IS NULL)) OR ([yr] = [RetireYear"
        "]) OR ([yr] > [ActiveYear] AND ([yr] <= [RetireYear] OR [RetireYear] IS NULL)) O"
        "RDER BY CategoryOrder; "
    Caption ="Modified Wentworth Key"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006801000000000000a80c00002001000000000000 ,
        0x020000000000000000000000a10700000100000001000000
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
            ControlSource ="KeyOrder"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =720
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Width =6480
                    Height =360
                    FontWeight =700
                    BackColor =11265523
                    BorderColor =8355711
                    Name ="lblTitle"
                    Caption ="NCPN Modified Wentworth Scale"
                    GridlineColor =10921638
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Top =360
                    Width =1368
                    Height =360
                    FontSize =8
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblClassSize1"
                    Caption ="Class Size"
                    GridlineColor =10921638
                    LayoutCachedTop =360
                    LayoutCachedWidth =1368
                    LayoutCachedHeight =720
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1368
                    Top =360
                    Width =432
                    Height =360
                    FontSize =8
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblCode1"
                    Caption ="Code"
                    GridlineColor =10921638
                    LayoutCachedLeft =1368
                    LayoutCachedTop =360
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =720
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1800
                    Top =360
                    Width =1440
                    Height =360
                    FontSize =8
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDiameter1"
                    Caption ="Diameter (mm)"
                    GridlineColor =10921638
                    LayoutCachedLeft =1800
                    LayoutCachedTop =360
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =720
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3240
                    Top =360
                    Width =1368
                    Height =360
                    FontSize =8
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblClassSize2"
                    Caption ="Class Size"
                    GridlineColor =10921638
                    LayoutCachedLeft =3240
                    LayoutCachedTop =360
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =720
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4608
                    Top =360
                    Width =432
                    Height =360
                    FontSize =8
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblCode2"
                    Caption ="Code"
                    GridlineColor =10921638
                    LayoutCachedLeft =4608
                    LayoutCachedTop =360
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =720
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =5040
                    Top =360
                    Width =1440
                    Height =360
                    FontSize =8
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDiameter2"
                    Caption ="Diameter (mm)"
                    GridlineColor =10921638
                    LayoutCachedLeft =5040
                    LayoutCachedTop =360
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =720
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =288
            OnFormat ="[Event Procedure]"
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            AlternateBackColor =12632256
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Width =3235
                    Height =288
                    FontSize =8
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblRow1"
                    GridlineColor =10921638
                    LayoutCachedWidth =3235
                    LayoutCachedHeight =288
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Width =1368
                    Height =288
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxClassSize"
                    ControlSource ="ClassSize"
                    LeftPadding =14
                    RightPadding =14
                    GridlineColor =10921638

                    LayoutCachedWidth =1368
                    LayoutCachedHeight =288
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1368
                    Width =432
                    Height =288
                    FontSize =8
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCode"
                    ControlSource ="Code"
                    GridlineColor =10921638

                    LayoutCachedLeft =1368
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =288
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1800
                    Height =288
                    FontSize =8
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDiameter"
                    ControlSource ="DiameterRange_mm"
                    GridlineColor =10921638

                    LayoutCachedLeft =1800
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =288
                End
            End
        End
        Begin PageFooter
            Height =0
            BackColor =8421504
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin FormFooter
            KeepTogether = NotDefault
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
' Report:       ModWentworthKey
' Level:        Application report
' Version:      1.01
'
' Description:  Modified Wentworth Key report object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 12, 2016
' References:
'  Allen Browne, April 2010
'  http://allenbrowne.com/ser-43.html
' Revisions:    BLC - 5/12/2016 - 1.00 - initial version
'               BLC - 12/14/2016 - 1.01 - revised to dynamic scale
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------

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
' Sub:          Report_Open
' Description:  Report opening event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 4, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/4/2016 - initial version
'   BLC - 12/14/2016 - revised to dynamic scale
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Dim ary() As String, strPark As String
    
    'defaults
    'header
    lblTitle.Caption = "NCPN Modified Wentworth Scale"
    lblClassSize1.Caption = "Class Size"
    lblClassSize2.Caption = "Class Size"
    lblCode1.Caption = "Code"
    lblCode2.Caption = "Code"
    lblDiameter1.Caption = "Diameter (mm)"
    lblDiameter2.Caption = "Diameter (mm)"
    
    'rows
'    lblRow1.Caption = "        Fines                              F        See SOP definition               Gravel                     GR                > 2 - 16"
'    lblRow2.Caption = "        Clay                               CL                                                       Pebble                     PB                 > 16 - 64"
'    lblRow3.Caption = "    Loam/Clay                        LC                                                       Cobble                    CO                > 64 - 256"
'    lblRow4.Caption = "       Loam                              LO                                                       Boulder                   BL                > 256 - 512"
'    lblRow5.Caption = "       Sand                               SA            0.5" & ChrW(uMu) & "m- 2mm              Bedrock, Hardpan       BR               > 512"

    'arrow --> from lines
    
    'data
    
    
    
    If Len(Nz(OpenArgs, "")) > 0 Or IsNull(OpenArgs) Then
        strPark = Nz(TempVars("ParkCode"), "")
    Else
        ary = Split(OpenArgs, "|")
        strPark = UCase(ary(0))
    End If
    
    'customizations, if any
    Select Case strPark
        Case "BLCA", "CANY", ""
        Case "DINO"
    End Select
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[ModWentworthKey Report])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------

' ---------------------------------
' Function:     Detail_Format
' Description:  report detail formatting actions
' Assumptions:  -
' Parameters:   Cancel - if format action should be cancelled (integer)
'               FormatCount - items to format (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 12, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/12/2016 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[ModWentworthKey Report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     Detail_Print
' Description:  report detail printing actions
' Assumptions:  -
' Parameters:   Cancel - if print action should be cancelled (integer)
'               PrintCount - items to print (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 12, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/12/2016 - initial version
' ---------------------------------
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Print[ModWentworthKey Report])"
    End Select
    Resume Exit_Handler
End Sub
