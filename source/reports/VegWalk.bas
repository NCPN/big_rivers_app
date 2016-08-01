Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
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
    Width =11532
    DatasheetFontHeight =11
    ItemSuffix =120
    Right =21330
    Bottom =9645
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0x8986edf6b5c0e440
    End
    Caption ="Species List"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006801000000000000100e0000a401000000000000 ,
        0x030000009000000000000000a20700000100000001000000
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
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="=[SeqNum]"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="=[Family]"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1320
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    Top =961
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
                    LayoutCachedTop =961
                    LayoutCachedWidth =492
                    LayoutCachedHeight =1251
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =9240
                    Top =1020
                    Width =2220
                    Height =300
                    FontSize =10
                    FontWeight =600
                    BorderColor =8355711
                    Name ="lblSamplingDate"
                    Caption ="Date: ____/____/_____"
                    GridlineColor =10921638
                    LayoutCachedLeft =9240
                    LayoutCachedTop =1020
                    LayoutCachedWidth =11460
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
                    Caption ="CANY Veg Walk"
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
                    Left =4560
                    Top =975
                    Width =912
                    Height =276
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblObserver"
                    Caption ="Observer:"
                    GridlineColor =10921638
                    LayoutCachedLeft =4560
                    LayoutCachedTop =975
                    LayoutCachedWidth =5472
                    LayoutCachedHeight =1251
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =6900
                    Top =975
                    Width =912
                    Height =276
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblRecorder"
                    Caption ="Recorder:"
                    GridlineColor =10921638
                    LayoutCachedLeft =6900
                    LayoutCachedTop =975
                    LayoutCachedWidth =7812
                    LayoutCachedHeight =1251
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
                Begin Label
                    Left =60
                    Top =612
                    Width =2460
                    Height =324
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblNCPNMonitoring"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =612
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =936
                    ThemeFontIndex =-1
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
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =396
            BackColor =0
            Name ="GroupHeader1"
            AutoHeight =255
            AlternateBackColor =0
            Begin
                Begin Label
                    TextAlign =2
                    Left =60
                    Width =3540
                    Height =360
                    FontSize =10
                    FontWeight =700
                    TopMargin =29
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblSpeciesFound"
                    Caption ="Species Found"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =240
                    Top =36
                    Width =720
                    Height =360
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblCheckMark"
                    Caption ="✔"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =36
                    LayoutCachedWidth =960
                    LayoutCachedHeight =396
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
            Height =0
            BreakLevel =1
            Name ="GroupHeader2"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =300
            BreakLevel =2
            BackColor =3223867
            Name ="GroupHeader0"
            AlternateBackColor =3223867
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Width =1500
                    Height =300
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BorderColor =10921638
                    ForeColor =16777215
                    Name ="tbxFamily"
                    ControlSource ="=IIf(Left([Species],3)=\"UNK\",\"UNKNOWN\",[Family])"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =315
            OnFormat ="[Event Procedure]"
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            AlternateBackColor =14869733
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =300
                    Width =2040
                    Height =300
                    FontSize =8
                    TopMargin =29
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSpecies"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2400
                    Width =1260
                    Height =300
                    FontSize =7
                    TabIndex =1
                    TopMargin =29
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLUCode"
                    ControlSource ="LU_Code"
                    GridlineColor =10921638

                    LayoutCachedLeft =2400
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =300
                End
                Begin Line
                    Top =300
                    Width =11520
                    Name ="lnSeparator"
                    GridlineColor =10921638
                    LayoutCachedTop =300
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =300
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
' Report:       VegWalk
' Level:        Application report
' Version:      1.00
'
' Description:  VegWalk report object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 11, 2016
' References:
'  Allen Browne, April 2011
'  http://allenbrowne.com/ser-43.html
' Revisions:    BLC - 5/11/2016 - 1.00 - initial version
'               BLC - 6/15/2016 - 1.01 - renamed from SpeciesList to VegWalk
' =================================

'---------------------
' NOTES:
'   VegWalk lists are specific to the combination of park, river segment & year
'   These lists must be present to generate the appropriate species recordset for display.
'---------------------

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Park As String
Private m_StartRow As Integer
Private m_EndRow As Integer
Private m_NumRows As Integer

'---------------------
' Event Declarations
'---------------------
Public Event InvalidRow(Value As Integer)
Public Event InvalidNumRows(Value As Integer)

'---------------------
' Properties
'---------------------
Public Property Let Park(Value As String)
    If Len(Value) = 4 Then
        m_Park = Value
    End If
End Property

Public Property Get Park() As String
    Park = m_Park
End Property

Public Property Let StartRow(Value As Integer)
    If Value > 0 Then
        m_StartRow = Value
    Else
        RaiseEvent InvalidRow(Value)
    End If
End Property

Public Property Get StartRow() As Integer
    StartRow = m_StartRow
End Property

Public Property Let EndRow(Value As Integer)
    If Value > 0 Then
        m_EndRow = Value
    Else
        RaiseEvent InvalidRow(Value)
    End If
End Property

Public Property Get EndRow() As Integer
    EndRow = m_EndRow
End Property

Public Property Let NumRows(Value As Integer)
    If Value > 0 Then
        m_NumRows = Value
    Else
        RaiseEvent InvalidNumRows(Value)
    End If
End Property

Public Property Get NumRows() As Integer
    NumRows = m_NumRows
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
    Set sopdata = GetSOPMetadata("VegWalk") '0-code, 1-SOP#, 2-Version, 3-Effective Date
    i = CInt(sopdata(1))

    lblMonitoring.Caption = "NCPN " & aryProtocol(0, 0)
    lblProtocolVersion.Caption = aryProtocol(0, 0) & " - " & "SOP #" & i & " - Version " & Format(sopdata(2), "0.00") & " - " & Format(sopdata(3), "mmm yyyy")

    'set river segment(s)
    arySegments = GetRiverSegments(Me.Park)
    For i = 0 To UBound(arySegments, 2)
        strSegments = strSegments & arySegments(0, i) & Space(4)
    Next
    strSegments = Left(strSegments, Len(strSegments) - 1)
    
    Me.lblRiverSegments.Caption = strSegments
    
    'column headers
    Me.lblSpeciesFound.Caption = "Species Found"
    Me.lblCheckMark.Caption = ChrW(uCheck)
    
    'setup data sources
    Select Case Me.Park
        Case "BLCA", ""
            strFamily = "Master_Family" '"CO_Family"
            strSpecies = "Co_species"
        Case "CANY"
            strFamily = "Master_Family" '"UT_Family"
            strSpecies = "Utah_species"
        Case "DINO"
            strFamily = "Master_Family" '"UT_Family"
            strSpecies = "Utah_species"
    End Select
    
    '###################################################################
    '# TODO: Adjust strSQL to filter by park, river segment & year     #
    '###################################################################

'    strSQL = "SELECT " & strFamily & " AS Family, " & strSpecies & " AS Species, LU_Code, " _
'            & "1 AS SeqNum " _
'            & "FROM tlu_NCPN_Plants " _
'            & "WHERE Master_Family NOT IN ('','Unknown') " _
'            & "UNION ALL " _
'            & "SELECT " & strFamily & " AS Family, " & strSpecies & " AS Species, LU_Code, " _
'            & "2 AS SeqNum " _
'            & "FROM tlu_NCPN_Plants " _
'            & "WHERE Master_Family = 'Unknown'" _
'            & "ORDER BY SeqNum ASC;"

    strSQL = GetTemplate("s_vegwalk", "strFamily" & PARAM_SEPARATOR & strFamily & "|strSpecies" & PARAM_SEPARATOR & strSpecies)
    
    Me.RecordSource = strSQL
    Debug.Print strSQL

    tbxSpecies.ControlSource = "Species"
    
    'hide modal Main form
    Forms("Main").visible = False
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[VegWalk Report])"
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
            "Error encountered (#" & Err.Number & " - Detail_Format[VegWalk Report])"
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
            "Error encountered (#" & Err.Number & " - Detail_Print[VegWalk Report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     Report_Close
' Description:  report closing actions
' Assumptions:  -
' Parameters:   Cancel - if print action should be cancelled (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 25, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/25/2016 - initial version
' ---------------------------------
Private Sub Report_Close()
On Error GoTo Err_Handler

    'unhide modal Main form
    Forms("Main").visible = True

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Close[VegWalk Report])"
    End Select
    Resume Exit_Handler
End Sub
