Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8280
    DatasheetFontHeight =11
    ItemSuffix =25
    Left =4365
    Top =3210
    Right =12645
    Bottom =11175
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x06dd372434a7e440
    End
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    SplitFormSplitterBar =0
    SaveSplitterBarPosition =0
    SplitFormSplitterBar =0
    SaveSplitterBarPosition =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
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
            SpecialEffect =3
            BackStyle =0
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
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
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =360
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =60
                    Top =60
                    Width =1980
                    Height =300
                    ForeColor =15921906
                    Name ="lblTitle"
                    Caption ="Data Entry"
                    ControlTipText ="Application role for logged in user."
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =360
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =95.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6645
                    Top =45
                    Height =299
                    ForeColor =15921906
                    Name ="btnAdmin"
                    Caption ="Admin"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Go to administration settings"
                    GridlineColor =10921638

                    LayoutCachedLeft =6645
                    LayoutCachedTop =45
                    LayoutCachedWidth =8085
                    LayoutCachedHeight =344
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =95.0
                    Gradient =0
                    BackColor =4144959
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderColor =14136213
                    HoverColor =9699294
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    Left =1380
                    Top =120
                    Width =915
                    Height =180
                    FontSize =9
                    ForeColor =15921906
                    Name ="lblAppUser"
                    Caption ="(app user)"
                    ControlTipText ="Logged in user"
                    GridlineColor =10921638
                    LayoutCachedLeft =1380
                    LayoutCachedTop =120
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =300
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =95.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7620
            BackColor =2634567
            Name ="Detail"
            AlternateBackColor =2634567
            AlternateBackShade =95.0
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =540
                    Width =2592
                    Height =3456
                    BorderColor =10921638
                    Name ="LTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =540
                    LayoutCachedWidth =2712
                    LayoutCachedHeight =3996
                End
                Begin Subform
                    OverlapFlags =85
                    Left =2850
                    Top =540
                    Width =2592
                    Height =3456
                    TabIndex =1
                    BorderColor =10921638
                    Name ="CTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =2850
                    LayoutCachedTop =540
                    LayoutCachedWidth =5442
                    LayoutCachedHeight =3996
                End
                Begin Subform
                    OverlapFlags =85
                    Left =5580
                    Top =540
                    Width =2592
                    Height =3456
                    TabIndex =2
                    BorderColor =10921638
                    Name ="RTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =5580
                    LayoutCachedTop =540
                    LayoutCachedWidth =8172
                    LayoutCachedHeight =3996
                End
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =4080
                    Width =2592
                    Height =3456
                    TabIndex =3
                    BorderColor =10921638
                    Name ="BLTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =4080
                    LayoutCachedWidth =2712
                    LayoutCachedHeight =7536
                End
                Begin Subform
                    OverlapFlags =85
                    Left =2850
                    Top =4080
                    Width =2592
                    Height =3456
                    TabIndex =4
                    BorderColor =10921638
                    Name ="BCTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =2850
                    LayoutCachedTop =4080
                    LayoutCachedWidth =5442
                    LayoutCachedHeight =7536
                End
                Begin Subform
                    OverlapFlags =85
                    Left =5580
                    Top =4080
                    Width =2592
                    Height =3456
                    TabIndex =5
                    BorderColor =10921638
                    Name ="BRTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =5580
                    LayoutCachedTop =4080
                    LayoutCachedWidth =8172
                    LayoutCachedHeight =7536
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =120
                    Top =60
                    Width =5832
                    Height =360
                    TabIndex =6
                    BorderColor =10921638
                    Name ="fsubBreadcrumb"
                    SourceObject ="Form._Level"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =5952
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Width =1176
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="breadcrumb Label"
                            Caption ="breadcrumb"
                            EventProcPrefix ="breadcrumb_Label"
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedWidth =1296
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =223
                    Width =8280
                    Height =480
                    BackColor =5540500
                    BorderColor =10921638
                    Name ="rctTop"
                    GridlineColor =10921638
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =480
                    BackThemeColorIndex =3
                    BackShade =50.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =6030
                    Top =60
                    Width =2100
                    Height =360
                    FontSize =8
                    ForeColor =15921906
                    Name ="lblNotice"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =6030
                    LayoutCachedTop =60
                    LayoutCachedWidth =8130
                    LayoutCachedHeight =420
                    ThemeFontIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =95.0
                End
            End
        End
        Begin FormFooter
            CanShrink = NotDefault
            Height =0
            Name ="FormFooter"
            AutoHeight =255
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
' Form:         Main
' Level:        Application form
' Version:      1.04
' Basis:        Main form
'
' Description:  Main switchboard form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, April 20, 2016
' References:   -
' Revisions:    BLC - 4/20/2016 - 1.00 - initial version
'               BLC - 6/28/2016 - 1.01 - added Form_Close event to clear breadcrumb
'               BLC - 9/6/2016  - 1.02 - added PrepareLinks() and updated UpdateBreadcrumb()
'               BLC - 9/8/2016  - 1.03 - code cleanup
'               BLC - 9/21/2016 - 1.04 - update PrepareLinks so BLCA & CANY enable transect links,
'                                        DINO does not
' =================================

'---------------------
' Declarations
'---------------------
Private WithEvents oTile As Form_Tile
Attribute oTile.VB_VarHelpID = -1

'declare tile form objects
Dim oLTile As Form_Tile
Dim oCTile As Form_Tile
Dim oRTile As Form_Tile
Dim oBLTile As Form_Tile
Dim oBCTile As Form_Tile
Dim oBRTile As Form_Tile

'---------------------
' Properties
'---------------------

'---------------------
' Events
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    'initialize app (mod_App_UI)
    Initialize
    
    'prepare interface
    Me.Detail.BackColor = lngNPSBrown
    Me.Detail.AlternateBackColor = lngNPSBrown

    Me.rctTop.Width = Me.Width
    
    'Main
    Me.lblTitle.Caption = IIf(TempVars("UserAccessLevel") = "admin", _
                            "Administrator", _
                            Nz(StrConv(TempVars("UserAccessLevel"), vbProperCase), _
                            "Data Entry"))
    lblTitle.ForeColor = lngWhite
    
    lblNotice.Caption = StringFromCodepoint(uPointerBlkL) & " R click to set values"
    lblNotice.FontSize = 9
    lblNotice.FontWeight = wtBold
    lblNotice.ForeColor = lngLtYellow
    
    'set app user
    lblAppUser.Caption = "(" & TempVars("AppUsername") & ")"
    lblAppUser.ForeColor = lngLtBlue
    
    'Admin --> filter actions when form opens (disable form buttons based on TempVars("AccessLevel")
    Me.btnAdmin.HoverColor = LINK_HIGHLIGHT_BKGD
        
    ' ------ Top -------
    'Left
    Set oLTile = LTile.Form
    oLTile.Title = "Where?"
    oLTile.BarColor = lngWhite
    oLTile.TileHeaderColor = lngLtSienna
    oLTile.TitleFontColor = lngWhite
    oLTile.Link1Caption = "Site"
    oLTile.Link2Caption = "Feature"
    oLTile.Link3Caption = "Transect"
    oLTile.Link4Caption = "Plot"
    oLTile.Link5Caption = ""
    oLTile.Link5Visible = 0
    oLTile.Link6Caption = "Location"
    oLTile.Link7Visible = 0
    oLTile.Link8Visible = 0
    
    'Center
    Set oCTile = CTile.Form
    oCTile.Title = "Sampling"
    oCTile.BarColor = lngWhite
    oCTile.TileHeaderColor = lngVanilla
    oCTile.Link1Caption = "Event"
    oCTile.Link2Caption = ""
    oCTile.Link2Visible = 0
    oCTile.Link3Caption = "VegPlots"
    oCTile.Link4Visible = 0
    oCTile.Link5Caption = "Locations"
    oCTile.Link6Caption = "People"
    oCTile.Link7Visible = 0
    oCTile.Link8Visible = 0

    'Right
    Set oRTile = RTile.Form
    oRTile.Title = "Vegetation"
    oRTile.BarColor = lngWhite
    oRTile.TileHeaderColor = lngSageGreen
    oRTile.TitleFontColor = lngWhite
    oRTile.Link1Caption = "Woody Canopy Cover"
    oRTile.Link2Caption = "Understory Cover"
    oRTile.Link3Caption = "Vegetation Walk"
    oRTile.Link4Visible = 0
    oRTile.Link5Visible = 0
    oRTile.Link6Caption = "Species"
    oRTile.Link7Caption = "Unknowns"
    oRTile.Link8Caption = "Species Search"

    ' ------ Bottom -------
    'Left
    Set oBLTile = BLTile.Form
    oBLTile.Title = "Observations"
    oBLTile.TileTag = "Obs-"
    oBLTile.BarColor = lngWhite
    oBLTile.TileHeaderColor = lngRobinEgg
    oBLTile.Link1Caption = "Photos"
    oBLTile.Link2Caption = "Transducers"
    oBLTile.Link3Visible = 0
    oBLTile.Link4Caption = "Survey Files"
    oBLTile.Link5Visible = 0
    oBLTile.Link6Visible = 0
    oBLTile.Link7Caption = Space(4) & "Upload Survey File"
    oBLTile.lblIcon7L.Caption = StringFromCodepoint(uSquareFoot)
    oBLTile.lblIcon7L.ForeColor = lngBlue
    oBLTile.lblIcon7L.FontWeight = wtMedium
    oBLTile.lblIcon8L.Caption = StringFromCodepoint(uPicFramed)
    oBLTile.lblIcon8L.ForeColor = lngBlue
    oBLTile.lblIcon8L.FontWeight = wtMedium
    oBLTile.Link8Caption = Space(4) & "Batch Upload Photos"
    
    'Center
    Set oBCTile = BCTile.Form
    oBCTile.Title = "Trip Prep"
    oBCTile.TileTag = "prep-"
    oBCTile.BarColor = lngWhite
    oBCTile.TileHeaderColor = lngLtSalmon
    oBCTile.Link1Caption = "VegPlot"
    oBCTile.Link2Caption = "VegWalk"
    oBCTile.Link3Caption = "Photo"
    oBCTile.Link4Caption = "Transducer"
    oBCTile.Link5Visible = 0
    oBCTile.Link6Caption = "Tasks"
    oBCTile.lblIcon7L.Caption = StringFromCodepoint(uMapLighthouse)
    oBCTile.lblIcon7L.ForeColor = lngBlue
    oBCTile.lblIcon7L.FontWeight = wtMedium
    oBCTile.Link7Caption = Space(4) & "Sediment Class Settings"
    oBCTile.lblIcon8L.Caption = StringFromCodepoint(uMapLighthouse)
    oBCTile.lblIcon8L.ForeColor = lngBlue
    oBCTile.lblIcon8L.FontWeight = wtMedium
    oBCTile.Link8Caption = Space(4) & "Sheet Settings"

    'Right
    Set oBRTile = BRTile.Form
    oBRTile.Title = "Reports"
    oBRTile.BarColor = lngWhite
    oBRTile.TileHeaderColor = lngBlue 'lngMimosa
    oBRTile.TitleFontColor = lngWhite
    oBRTile.Link1Caption = Space(4) & "# Plots"
    oBRTile.Link2Caption = Space(4) & "VegPlot - Species"
    oBRTile.Link3Caption = Space(4) & "VegPlot - Species #s"
    oBRTile.Link4Caption = Space(4) & "VegWalk - Species"
    oBRTile.Link5Caption = Space(4) & "VegWalk - Species #s"
    oBRTile.Link6Visible = 0
    oBRTile.Link7Visible = 0
    oBRTile.Link6Caption = ""
    oBRTile.Link7Caption = ""
    oBRTile.Link8Caption = "More..."
    oBRTile.lblLink8.TextAlign = aRight
    oBRTile.lblLink8.FontItalic = True
    
    Dim i As Integer
    Dim strControl As String
    
    For i = 1 To 5
        strControl = "lblIcon" & i & "L"
        With oBRTile.Controls(strControl)
            .Caption = StringFromCodepoint(uDocumentEmpty)
            .FontWeight = wtNormal
            .FontSize = 12
            .ForeColor = lngBlue
        End With
    Next
    
    HighlightBreadcrumb
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Main form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  Form actions when form the current form
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 18, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/18/2016 - initial version
'   BLC - 7/13/2016 - add exceptions for certain links (active w/o park selected)
'   BLC - 9/8/2016  - move enabling links to PrepareLinks() & code cleanup
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
    
    'enable links based on TempVar/breadcrumb settings
    PrepareLinks
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[Main form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAdmin_Click
' Description:  Admin button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 7, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/7/2016 - initial version
' ---------------------------------
Private Sub btnAdmin_Click()
On Error GoTo Err_Handler

    'open admin form
    DoCmd.OpenForm "DbAdmin", acNormal
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAdmin_Click[Main form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Close
' Description:  Form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 28, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/28/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'clear breadcrumb
    TempVars.Remove "ParkCode"
    TempVars.Remove "River"
    TempVars.Remove "SiteCode"
    TempVars.Remove "Feature"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[Main form])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          UpdateBreadcrumb
' Description:  Update captions w/in breadcrumb
' Assumptions:
' Parameters:   ClearValues - breadcrumb values to clear? (integer)
'                             0,1,2,3-highest level to clear, levels below this level are cleared
'                             (captions set to "Missing XX >")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 18, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/18/2016 - initial version
'   BLC - 7/1/2016 -  update indicator color & visibility
'   BLC - 9/6/2016 - added PrepareLinks() call to update links,
'                    shifted setting captions until AFTER TempVars clearing,
'                    fixed issue preventing Feature from clearing
'   BLC - 9/8/2016 - remove enable links which is addressed in PrepareLinks()
' ---------------------------------
Public Sub UpdateBreadcrumb(Optional ClearValues As Integer = 4)
On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim strLevel As String, strMore As String
    Dim strHierarchy() As Variant
    Dim frm As Form
    Dim bgcolor0 As Long, bgcolor1 As Long, bgcolor2 As Long, bgcolor3 As Long
        
    Set frm = Me!fsubBreadcrumb.Form
    
    With frm

        'clear
        strHierarchy() = Array("Park", "River", "Site", "Feature")
        
        'ClearValues depends on the level: 0-Park, 1-River, 2-Site, 3-Feature
        For i = ClearValues To 3
        
            'default
            strMore = ""
            
            If i < 4 Then strMore = Space(4) & ">"
        
            strLevel = "btnLevel" & i
            .Controls(strLevel).Caption = strHierarchy(i) & strMore
        
            Select Case i
                Case 0 'park
'                    TempVars.Remove "ParkCode" --> only removed when Main form closed
                    TempVars.Remove "River"
                    
                Case 1 'river
                    TempVars.Remove "SiteCode"
                    
                Case 2 'site
                   TempVars.Remove "Feature"
                
                Case 3 'feature
                   'nothing to remove
            
            End Select
        
        Next
        
        'update captions *AFTER* removing tempvars
        .btnLevel0.Caption = Nz(TempVars("ParkCode"), "Park") & Space(4) & ">"
        .btnLevel1.Caption = Nz(TempVars("River"), "River") & Space(4) & ">"
        .btnLevel2.Caption = Nz(TempVars("SiteCode"), "Site") & Space(4) & ">"
        .btnLevel3.Caption = Nz(TempVars("Feature"), "Feature")
            
            
        'if park --> enable links
        If Len(Nz(TempVars("ParkCode"), "")) > 0 Then
            
            'clear notice
            lblNotice.Caption = ""
        
            Dim ctrl As Control
            
            'iterate through the tiles - update indicator & enable links
            For Each ctrl In Me.Controls
            
                If Right(ctrl.Name, 4) = "Tile" Then
                    
                    With ctrl.Form
                        
                        '.EnableLinks .TileTag & strLinksToEnable <-- enabled via PrepareLinks
                        .IndicatorVisible = 1
                        .IndicatorColor = lngGreen
                        
                    End With
                End If
            
            Next
            
            'disable feature for non-feature parks
            If TempVars("ParkCode") <> "BLCA" Then
                .btnLevel2.Caption = Replace(.btnLevel2.Caption, ">", "")
                .btnLevel3.Caption = ""
            End If
        End If
        
        HighlightBreadcrumb

    End With
    
    'enable links based on TempVar/breadcrumb settings
    PrepareLinks
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateBreadcrumb[Main form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          HighlightUpdateBreadcrumb
' Description:  Highlight captions w/in breadcrumb
' Assumptions:
' Parameters:   ClearValues - breadcrumb values to clear? (integer)
'                             0,1,2,3-highest level to clear, levels below this level are cleared (captions set to "Missing XX >")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 18, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/18/2016 - initial version
' ---------------------------------
Public Sub HighlightBreadcrumb(Optional ClearValues As Integer = 4)
On Error GoTo Err_Handler
    
    Dim frm As Form
    Dim bgcolor0 As Long, bgcolor1 As Long, bgcolor2 As Long, bgcolor3 As Long
        
    Set frm = Me!fsubBreadcrumb.Form
    
    With frm
        'highlight buttons where Park/River/Site/Feature not set
        If Len(.btnLevel0.Caption) <> Len(Replace(.btnLevel0.Caption, "Park", "")) Then
            bgcolor0 = HIGHLIGHT_MISSING_VALUE
        Else
            bgcolor0 = lngWhite
        End If
        If Len(.btnLevel1.Caption) <> Len(Replace(.btnLevel1.Caption, "River", "")) Then
            bgcolor1 = HIGHLIGHT_MISSING_VALUE
        Else
            bgcolor1 = lngWhite
        End If
        If Len(.btnLevel2.Caption) <> Len(Replace(.btnLevel2.Caption, "Site", "")) Then
            bgcolor2 = HIGHLIGHT_MISSING_VALUE
        Else
            bgcolor2 = lngWhite
        End If
        If Len(.btnLevel3.Caption) <> Len(Replace(.btnLevel3.Caption, "Feature", "")) Then
            bgcolor3 = HIGHLIGHT_MISSING_VALUE
        Else
            bgcolor3 = lngWhite
        End If

        .btnLevel0.BackColor = bgcolor0
        .btnLevel1.BackColor = bgcolor1
        .btnLevel2.BackColor = bgcolor2
        .btnLevel3.BackColor = bgcolor3

    End With
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HighlightBreadcrumb[Main form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          PrepareLinks
' Description:  Identifies appropriate links to enable depending on breadcrumb settings
'               (via TempVars for ParkCode, River, Site, Feature)
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 6, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/6/2016  - initial version
'   BLC - 9/21/2016 - update so BLCA & CANY enable transect links, DINO does not
' ---------------------------------
Private Sub PrepareLinks()
On Error GoTo Err_Handler

    'T=top, B=bottom, L=left, C=center, R=right (TC=top center tile)
    Dim TL As String, TC As String, TR As String
    Dim BL As String, BC As String, BR As String

    'use set to allow use of tile methods
    Set oLTile = LTile.Form
    Set oCTile = CTile.Form
    Set oRTile = RTile.Form
    Set oBLTile = BLTile.Form
    Set oBCTile = BCTile.Form
    Set oBRTile = BRTile.Form
    
    'default sets
    TL = ""             'N/A
    TC = "6"            'People
    TR = "6,7,8"        'Species, Unknowns, Species Search
    BL = ""             'N/A
    BC = "7"            'Mod Wentworth Settings
    BR = ""             'N/A
    
    'if no park --> all links deactivated, EXCEPT these
    If Len(Nz(TempVars("ParkCode"), "")) > 0 Then
        
        'default sets
        TL = ""             'N/A
        TC = "6"            'People
        TR = "6,7,8"        'Species, Unknowns, Species Search
        BL = "4,7"          'Survey Files, Upload Survey File
        BC = "1,2,3,4,6,7,8"  'VegPlot, VegWalk, Photo, Transducer, Tasks, Mod Wentworth Settings, Sheet Settings
        BR = "1,2,3,4,5,8"  '#Plots, VegPlot-Species, VegPlot-#Species, VegWalk-Species, VegWalk-#Species
        
        'prepare park specific sets
        Select Case TempVars("ParkCode")

            Case "BLCA"
                If Len(Nz(TempVars("River"), "")) > 0 Then
                    
                    TL = "1"    'Site
                    
                    If Len(Nz(TempVars("SiteCode"), "")) > 0 Then
                        
                        TL = "1,2,6"     'Site, Feature, Location
                        TC = "1,5,6"     'Event, Locations, People
                        BL = "2,4,7,8"   'Transducers, Survey Files, Upload Survey File,
                                         ' Batch Upload Photos
                        BR = "1,2,3,4,5,8"  '#Plots, VegPlot-Species, VegPlot-#Species,
                                            ' VegWalk-Species, VegWalk-#Species, More
                        
                        If Len(Nz(TempVars("Feature"), "")) > 0 Then
                            TL = "1,2,3,4,6"    'Site, Feature, Transect, Plot, Location
                            TC = "1,3,5,6"      'Event, VegPlots, Locations, People
                            BL = "1,2,4,7,8"   'Photos, Transducers, Survey Files, Upload Survey File,
                                               ' Batch Upload Photos
                            BR = "1,2,3,4,5,8"  '#Plots, VegPlot-Species, VegPlot-#Species,
                                                ' VegWalk-Species, VegWalk-#Species, More
                        End If
                    
                    End If
                
                End If
            Case "CANY"
                If Len(Nz(TempVars("River"), "")) > 0 Then
                    
                    TL = "1"    'Site
                    
                    If Len(Nz(TempVars("SiteCode"), "")) > 0 Then
                        TL = "1,3,4,6"      'Site, Transect, Plot, Location
                        TC = "1,3,5,6"      'Event, VegPlots, Locations, People
                        BL = "1,2,4,7,8"    'Photos, Transducers, Survey Files, Upload Survey File,
                                            ' Batch Upload Photos
                        BR = "1,2,3,4,5,8"  '#Plots, VegPlot-Species, VegPlot-#Species,
                                            ' VegWalk-Species, VegWalk-#Species, More
                    End If
                
                End If
            Case "DINO"
                If Len(Nz(TempVars("River"), "")) > 0 Then
                    
                    TL = "1"    'Site
                    
                    If Len(Nz(TempVars("SiteCode"), "")) > 0 Then
                        TL = "1,4,6"        'Site, Plot, Location
                        TC = "1,3,5,6"      'Event, VegPlots, Locations, People
                        TR = "3,6,7,8"      'VegWalk, Species, Unknowns, Species Search
                        BL = "1,2,4,7,8"    'Photos, Transducers, Survey Files, Upload Survey File,
                                            ' Batch Upload Photos
                        BR = "1,2,3,4,5,8"  '#Plots, VegPlot-Species, VegPlot-#Species,
                                            ' VegWalk-Species, VegWalk-#Species, More
                    End If
                
                End If
        End Select

    End If

    'disable before selective re-enabling
    Dim ctrl As Control

    'iterate through the tiles - disable links
    For Each ctrl In Me.Controls

        If Right(ctrl.Name, 4) = "Tile" Then
    
            ctrl.Form.DisableLinks
    
        End If
        
    Next
    
    'selectively re-enable links
    oLTile.EnableLinks TL    'Site
    oCTile.EnableLinks TC    'People
    oRTile.EnableLinks TR    'Species, Unknowns, Species Search
    oBLTile.EnableLinks BL   'Batch Upload Photos
    oBCTile.EnableLinks BC   'VegPlot, VegWalk, Photo, Transducer, Tasks, Sheet Settings
    oBRTile.EnableLinks BR   '#Plots, VegPlot-Species, VegPlot-#Species, VegWalk-Species, VegWalk-#Species
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PrepareLinks[Main form])"
    End Select
    Resume Exit_Handler
End Sub
