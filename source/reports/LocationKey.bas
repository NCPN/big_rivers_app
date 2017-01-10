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
    Width =6840
    DatasheetFontHeight =11
    ItemSuffix =108
    Right =25395
    Bottom =11790
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0x32f638d6d6a9e440
    End
    Caption ="Photo Key"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006d01000000000000b81a00000000000001000000 ,
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
            ControlSource ="=1"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin PageHeader
            Height =0
            BackColor =14540253
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1679
            BackColor =14540253
            Name ="GroupHeader0"
            AlternateBackColor =14540253
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =60
                    Top =96
                    Width =180
                    Height =1498
                    FontSize =7
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblLocation"
                    Caption ="LOCATION"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =96
                    LayoutCachedWidth =240
                    LayoutCachedHeight =1594
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2448
                    Top =60
                    Width =2160
                    Height =360
                    FontSize =6
                    FontWeight =700
                    TopMargin =29
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblPhotogFeature"
                    Caption ="F"
                    GridlineColor =10921638
                    LayoutCachedLeft =2448
                    LayoutCachedTop =60
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =420
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4632
                    Top =60
                    Width =2160
                    Height =360
                    FontSize =6
                    FontWeight =700
                    TopMargin =29
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblPhotogTransect"
                    Caption ="T"
                    GridlineColor =10921638
                    LayoutCachedLeft =4632
                    LayoutCachedTop =60
                    LayoutCachedWidth =6792
                    LayoutCachedHeight =420
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =270
                    Top =60
                    Width =2160
                    Height =749
                    FontSize =8
                    FontWeight =700
                    TopMargin =144
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblPhotographer"
                    Caption ="PHOTOGRAPHER"
                    GridlineColor =10921638
                    LayoutCachedLeft =270
                    LayoutCachedTop =60
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =809
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =270
                    Top =855
                    Width =2160
                    Height =749
                    FontSize =7
                    FontWeight =700
                    TopMargin =144
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblSubject"
                    Caption ="SUBJECT"
                    GridlineColor =10921638
                    LayoutCachedLeft =270
                    LayoutCachedTop =855
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =1604
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2448
                    Top =468
                    Width =2160
                    Height =360
                    FontSize =6
                    FontWeight =700
                    TopMargin =29
                    BorderColor =8355711
                    Name ="lblPhotogOverview"
                    Caption ="O"
                    GridlineColor =10921638
                    LayoutCachedLeft =2448
                    LayoutCachedTop =468
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =828
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4632
                    Top =468
                    Width =2160
                    Height =360
                    FontSize =6
                    FontWeight =700
                    TopMargin =29
                    BorderColor =8355711
                    Name ="lblPhotogRef"
                    Caption ="R"
                    GridlineColor =10921638
                    LayoutCachedLeft =4632
                    LayoutCachedTop =468
                    LayoutCachedWidth =6792
                    LayoutCachedHeight =828
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =306
                    Top =1275
                    Width =2016
                    Height =180
                    FontSize =6
                    BorderColor =8355711
                    Name ="lblSubjectHint"
                    Caption ="Location being photographed"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =306
                    LayoutCachedTop =1275
                    LayoutCachedWidth =2322
                    LayoutCachedHeight =1455
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2448
                    Top =864
                    Width =2160
                    Height =360
                    FontSize =6
                    FontWeight =700
                    TopMargin =29
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblSubjectFeature"
                    Caption ="F"
                    GridlineColor =10921638
                    LayoutCachedLeft =2448
                    LayoutCachedTop =864
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =1224
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4632
                    Top =864
                    Width =2160
                    Height =360
                    FontSize =6
                    FontWeight =700
                    TopMargin =29
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblSubjectTransect"
                    Caption ="T"
                    GridlineColor =10921638
                    LayoutCachedLeft =4632
                    LayoutCachedTop =864
                    LayoutCachedWidth =6792
                    LayoutCachedHeight =1224
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2448
                    Top =1248
                    Width =2160
                    Height =360
                    FontSize =6
                    FontWeight =700
                    TopMargin =29
                    BorderColor =8355711
                    Name ="lblSubjectOverview"
                    Caption ="O"
                    GridlineColor =10921638
                    LayoutCachedLeft =2448
                    LayoutCachedTop =1248
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =1608
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4632
                    Top =1248
                    Width =2160
                    Height =360
                    FontSize =6
                    FontWeight =700
                    TopMargin =29
                    BorderColor =8355711
                    Name ="lblSubjectRef"
                    Caption ="R"
                    GridlineColor =10921638
                    LayoutCachedLeft =4632
                    LayoutCachedTop =1248
                    LayoutCachedWidth =6792
                    LayoutCachedHeight =1608
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =0
            OnFormat ="[Event Procedure]"
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            AlternateBackColor =12632256
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
' Form:         LocationKey
' Level:        Application form
' Version:      1.00
'
' Description:  LocationKey form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 11, 2016
' References:
'  Allen Browne, April 2011
'  http://allenbrowne.com/ser-43.html
' Revisions:    BLC - 5/11/2016 - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Dim m_Park As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidPark(Park As String)

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let Park(Value As String)
    If Len(Value) = 4 Then
        m_Park = Value
    Else
        RaiseEvent InvalidPark(Value)
    End If
End Property

Public Property Get Park() As String
    Park = m_Park
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
' Source/date:  Bonnie Campbell, May 4, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/4/2016 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Dim strPF As String, strPT As String, strPO As String, strPR As String
    Dim strSF As String, strST As String, strSO As String, strSR As String
    Dim strArrow As String
    Dim ary() As String
    
    'default
    strArrow = " " & ChrW(uRArrow) & " " 'right arrow c.f. https://en.wikipedia.org/wiki/Arrow_(symbol)
    
    'location (vertical display)
    lblLocation.Caption = "L" & vbCrLf & "O" & vbCrLf & "C" & vbCrLf & "A" _
                        & vbCrLf & "T" & vbCrLf & "I" & vbCrLf & "O" & vbCrLf & "N"
    
    'set park
    If Len(Nz(OpenArgs, "")) > 0 Or IsNull(OpenArgs) Then
        Me.Park = Nz(TempVars("ParkCode"), "")
    Else
        ary = Split(OpenArgs, "|")
        Me.Park = UCase(ary(0))
    End If
        
    'setup key captions
    Select Case Me.Park
        Case "BLCA", ""
            strPF = "Feature letter transect #(s) @" _
                    & vbCrLf & "tagline distance in m (J1@20)"
            strPT = "Feature letter transect # - tagline" _
                    & vbCrLf & "dist. in m (G1LHP, G1LP@26)"
            strPO = "PP1, PP2, PP3"
            strPR = "from river, 10m upstream, etc."
            strSF = "Feature letter transect #(s) @" _
                    & vbCrLf & "tagline distance in m (J1@20)"
            strST = "Feature letter transect # - tagline" _
                    & vbCrLf & "dist. in m (G1LHP, G1LP@26)"
            strSO = "O # - feature(s) O1-GH"
            strSR = "CP1, RM2, etc."
        Case "CANY"
            strPF = "F transect #(s)-order # (F3/4-2)"
            strPT = "T transect # - order # (T2-1)"
            strPO = "PP1, PP2, PP3"
            strPR = "from river, 10m upstream, etc."
            strSF = "N/A"
            strST = "N/A"
            strSO = "O1, O2, O3"
            strSR = "CP1, RM2, etc."
        Case "DINO"
            strPF = "N/A"
            strPT = "N/A"
            strPO = "PP1, PP2, PP3"
            strPR = "from river, 10m upstream, etc."
            strSF = "N/A"
            strST = "N/A"
            strSO = "O1, O2, O3"
            strSR = "CP1, RM2, etc."
    End Select
    
    'iterate & position controls
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        With ctrl
            Select Case .Name
                '-- photographer --
                Case "lblPhotogTransect" ' Transect
                    .Caption = "T" & strArrow & strPT
                Case "lblPhotogFeature"  ' Feature
                    .Caption = "F" & strArrow & strPF
                Case "lblPhotogOverview" ' Overview
                    .Caption = "O" & strArrow & strPO
                Case "lblPhotogRef"      ' Reference
                    .Caption = "R" & strArrow & strPR
                '-- subject --
                Case "lblSubjectTransect" ' Transect
                    .Caption = "T" & strArrow & strST
                Case "lblSubjectFeature"  ' Feature
                    .Caption = "F" & strArrow & strSF
                Case "lblSubjectOverview" ' Overview
                    .Caption = "O" & strArrow & strSO
                Case "lblSubjectRef"      ' Reference
                    .Caption = "R" & strArrow & strSR
            End Select
        End With
    Next
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[LocationKey Report])"
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
' Source/date:  Bonnie Campbell, May 11, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/11/2016 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[LocationKey Report])"
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
' Source/date:  Bonnie Campbell, May 11, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/11/2016 - initial version
' ---------------------------------
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Print[LocationKey Report])"
    End Select
    Resume Exit_Handler
End Sub
