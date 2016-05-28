Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8640
    DatasheetFontHeight =11
    ItemSuffix =21
    Left =3150
    Top =3105
    Right =12045
    Bottom =10770
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x06dd372434a7e440
    End
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
    FitToScreen =255
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
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =1980
                    Height =300
                    ForeColor =15921906
                    Name ="lblTitle"
                    Caption ="Data Entry"
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7080
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =540
                    Width =2592
                    Height =3168
                    BorderColor =10921638
                    Name ="LTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =540
                    LayoutCachedWidth =2712
                    LayoutCachedHeight =3708
                End
                Begin Subform
                    OverlapFlags =85
                    Left =2850
                    Top =540
                    Width =2592
                    Height =3168
                    TabIndex =1
                    BorderColor =10921638
                    Name ="CTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =2850
                    LayoutCachedTop =540
                    LayoutCachedWidth =5442
                    LayoutCachedHeight =3708
                End
                Begin Subform
                    OverlapFlags =85
                    Left =5580
                    Top =540
                    Width =2592
                    Height =3168
                    TabIndex =2
                    BorderColor =10921638
                    Name ="RTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =5580
                    LayoutCachedTop =540
                    LayoutCachedWidth =8172
                    LayoutCachedHeight =3708
                End
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =3840
                    Width =2592
                    Height =3168
                    TabIndex =3
                    BorderColor =10921638
                    Name ="BLTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =3840
                    LayoutCachedWidth =2712
                    LayoutCachedHeight =7008
                End
                Begin Subform
                    OverlapFlags =85
                    Left =2850
                    Top =3840
                    Width =2592
                    Height =3168
                    TabIndex =4
                    BorderColor =10921638
                    Name ="BCTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =2850
                    LayoutCachedTop =3840
                    LayoutCachedWidth =5442
                    LayoutCachedHeight =7008
                End
                Begin Subform
                    OverlapFlags =85
                    Left =5580
                    Top =3840
                    Width =2592
                    Height =3168
                    TabIndex =5
                    BorderColor =10921638
                    Name ="BRTile"
                    SourceObject ="Form.Tile"
                    GridlineColor =10921638

                    LayoutCachedLeft =5580
                    LayoutCachedTop =3840
                    LayoutCachedWidth =8172
                    LayoutCachedHeight =7008
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
                    Width =8640
                    Height =480
                    BackColor =5540500
                    BorderColor =10921638
                    Name ="rctTop"
                    GridlineColor =10921638
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =480
                    BackThemeColorIndex =3
                    BackShade =50.0
                End
            End
        End
        Begin FormFooter
            Height =240
            Name ="FormFooter"
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
' Version:      1.00
' Basis:        Main form
'
' Description:  Main switchboard form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, April 20, 2016
' References:   -
' Revisions:    BLC - 4/20/2016 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private WithEvents oTile As Form_Tile
Attribute oTile.VB_VarHelpID = -1

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
    Dim oLTile As Form_Tile
    Dim oCTile As Form_Tile
    Dim oRTile As Form_Tile
    Dim oBLTile As Form_Tile
    Dim oBCTile As Form_Tile
    Dim oBRTile As Form_Tile
    
    'Main
    Me.lblTitle.Caption = "Data Entry"
    
    ' ------ Top -------
    'Left
    Set oLTile = LTile.Form
    oLTile.Title = "Where?"
    oLTile.BarColor = vbGreen
    oLTile.TileHeaderColor = vbGreen
    oLTile.Link1Caption = "Site"
    oLTile.Link2Caption = "Feature"
    oLTile.Link3Caption = "Transect"
    oLTile.Link4Caption = "Plot"
    oLTile.Link5Caption = ""
    oLTile.Link6Caption = ""
    oLTile.Link5Visible = 0
    oLTile.Link6Visible = 0
    
    'Center
    Set oCTile = CTile.Form
    oCTile.Title = "Sampling"
    oCTile.BarColor = vbGreen
    oCTile.TileHeaderColor = vbGreen
    oCTile.Link1Caption = "Event"
    oCTile.Link2Caption = ""
    oCTile.Link2Visible = 0
    oCTile.Link3Caption = "Location"
    oCTile.Link4Visible = 0
    oCTile.Link5Caption = "People"
    oCTile.Link6Visible = 0

    'Right
    Set oRTile = RTile.Form
    oRTile.Title = "Vegetation"
    oRTile.BarColor = vbYellow
    oRTile.TileHeaderColor = vbBlue
    oRTile.Link1Caption = "Woody Canopy Cover"
    oRTile.Link2Caption = "Understory Cover"
    oRTile.Link3Caption = "Vegetation Walk"
    oRTile.Link4Visible = 0
    oRTile.Link5Visible = 0
    oRTile.Link6Caption = "Species"
    
    ' ------ Bottom -------
    'Left
    Set oBLTile = BLTile.Form
    oBLTile.Title = "Observations"
    oBLTile.TileTag = "Obs-"
    oBLTile.BarColor = vbGreen
    oBLTile.TileHeaderColor = vbGreen
    oBLTile.Link1Caption = "Photos"
    oBLTile.Link2Caption = "Transducers"
    oBLTile.Link3Visible = 0
    oBLTile.Link4Visible = 0
    oBLTile.Link5Visible = 0
    oBLTile.Link6Visible = 0
    
    'Center
    Set oBCTile = BCTile.Form
    oBCTile.Title = "Trip Prep"
    oBCTile.BarColor = vbGreen
    oBCTile.TileHeaderColor = vbGreen
    oBCTile.Link1Caption = "VegPlot"
    oBCTile.Link2Caption = "VegWalk"
    oBCTile.Link3Caption = "Photos"
    oBCTile.Link4Caption = "Transducer"
    oBCTile.Link5Visible = 0
    oBCTile.Link6Caption = "Tasks"

    'Right
    Set oBRTile = BRTile.Form
    oBRTile.Title = "Reports"
    oBRTile.BarColor = vbYellow
    oBRTile.TileHeaderColor = vbBlue
    oBRTile.Link1Caption = "rpt1"
    oBRTile.Link2Visible = 0
    oBRTile.Link3Visible = 0
    oBRTile.Link4Visible = 0
    oBRTile.Link5Visible = 0
    oBRTile.Link6Visible = 0
    oBRTile.Link2Caption = ""
    oBRTile.Link3Caption = ""
    oBRTile.Link4Caption = ""
    oBRTile.Link5Caption = ""
    oBRTile.Link6Caption = ""
    
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
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    'if no park --> deactivate links
    If Len(Nz(TempVars("ParkCode"), "")) = 0 Then
'        MsgBox "current" 'oBCTile.DisableLinks "1,2" '"1,2,3,4,5,6"
    End If
    
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

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          UpdateBreadcrumb
' Description:  Update captions w/in breadcrumb
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
Public Sub UpdateBreadcrumb(Optional ClearValues As Integer = 4)
On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim strLevel As String, strMore As String
    Dim strHierarchy() As Variant
    Dim frm As Form
    Dim bgcolor0 As Long, bgcolor1 As Long, bgcolor2 As Long, bgcolor3 As Long
        
    Set frm = Me!fsubBreadcrumb.Form
    
    With frm
        .btnLevel0.Caption = Nz(TempVars("ParkCode"), "Park") & Space(4) & ">"
        .btnLevel1.Caption = Nz(TempVars("River"), "River") & Space(4) & ">"
        .btnLevel2.Caption = Nz(TempVars("SiteCode"), "Site") & Space(4) & ">"
        .btnLevel3.Caption = Nz(TempVars("Feature"), "Feature")

        'clear
        strHierarchy() = Array("Park", "River", "Site", "Feature")
        
        For i = ClearValues + 1 To 3 - ClearValues
            
            'default
            strMore = ""
            
            If i < 4 Then strMore = Space(4) & ">"
        
            strLevel = "btnLevel" & i
            .Controls(strLevel).Caption = strHierarchy(i) & strMore
        
        Next
            
        'if park --> enable links
        If Len(Nz(TempVars("ParkCode"), "")) > 0 Then
            'enable links
            .Parent!LTile.Form.EnableLinks .Parent!LTile.Form.TileTag & ",1,2,3,4,5,6"
            .Parent!CTile.Form.EnableLinks .Parent!CTile.Form.TileTag & ",1,2,3,4,5,6"
            .Parent!RTile.Form.EnableLinks .Parent!RTile.Form.TileTag & ",1,2,3,4,5,6"
            .Parent!BLTile.Form.EnableLinks .Parent!BLTile.Form.TileTag & ",1,2,3,4,5,6"
            .Parent!BCTile.Form.EnableLinks .Parent!BCTile.Form.TileTag & ",1,2,3,4,5,6"
            .Parent!BRTile.Form.EnableLinks .Parent!BRTile.Form.TileTag & ",1,2,3,4,5,6"

            'disable feature for non-feature parks
            If TempVars("ParkCode") <> "BLCA" Then
                .btnLevel2.Caption = Replace(.btnLevel2.Caption, ">", "")
                .btnLevel3.Caption = ""
            End If
        End If
        
        HighlightBreadcrumb
'        'highlight buttons where Park/River/Site/Feature not set
'        If Len(.btnLevel0.Caption) <> Len(Replace(.btnLevel0.Caption, "Park", "")) Then
'            bgcolor0 = HIGHLIGHT_MISSING_VALUE
'        Else
'            bgcolor0 = lngWhite
'        End If
'        If Len(.btnLevel1.Caption) <> Len(Replace(.btnLevel1.Caption, "River", "")) Then
'            bgcolor1 = HIGHLIGHT_MISSING_VALUE
'        Else
'            bgcolor1 = lngWhite
'        End If
'        If Len(.btnLevel2.Caption) <> Len(Replace(.btnLevel2.Caption, "Site", "")) Then
'            bgcolor2 = HIGHLIGHT_MISSING_VALUE
'        Else
'            bgcolor2 = lngWhite
'        End If
'        If Len(.btnLevel3.Caption) <> Len(Replace(.btnLevel3.Caption, "Feature", "")) Then
'            bgcolor3 = HIGHLIGHT_MISSING_VALUE
'        Else
'            bgcolor3 = lngWhite
'        End If
'
'        .btnLevel0.backColor = bgcolor0
'        .btnLevel1.backColor = bgcolor1
'        .btnLevel2.backColor = bgcolor2
'        .btnLevel3.backColor = bgcolor3

    End With
    
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
' Sub:          UpdateBreadcrumb
' Description:  Update captions w/in breadcrumb
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
    
'    Dim i As Integer
'    Dim strLevel As String, strMore As String
'    Dim strHierarchy() As Variant
    Dim frm As Form
    Dim bgcolor0 As Long, bgcolor1 As Long, bgcolor2 As Long, bgcolor3 As Long
        
    Set frm = Me!fsubBreadcrumb.Form
    
    With frm
'        .btnLevel0.Caption = Nz(TempVars("ParkCode"), "Park") & Space(4) & ">"
'        .btnLevel1.Caption = Nz(TempVars("River"), "River") & Space(4) & ">"
'        .btnLevel2.Caption = Nz(TempVars("SiteCode"), "Site") & Space(4) & ">"
'        .btnLevel3.Caption = Nz(TempVars("Feature"), "Feature")
'
'        'clear
'        strHierarchy() = Array("Park", "River", "Site", "Feature")
'
'        For i = ClearValues + 1 To 3 - ClearValues
'
'            'default
'            strMore = ""
'
'            If i < 4 Then strMore = Space(4) & ">"
'
'            strLevel = "btnLevel" & i
'            .Controls(strLevel).Caption = strHierarchy(i) & strMore
'
'        Next
'
'        'if park --> enable links
'        If Len(Nz(TempVars("ParkCode"), "")) > 0 Then
'            'enable links
'            .Parent!LTile.Form.EnableLinks .Parent!LTile.Form.TileTag & ",1,2,3,4,5,6"
'            .Parent!CTile.Form.EnableLinks .Parent!CTile.Form.TileTag & ",1,2,3,4,5,6"
'            .Parent!RTile.Form.EnableLinks .Parent!RTile.Form.TileTag & ",1,2,3,4,5,6"
'            .Parent!BLTile.Form.EnableLinks .Parent!BLTile.Form.TileTag & ",1,2,3,4,5,6"
'            .Parent!BCTile.Form.EnableLinks .Parent!BCTile.Form.TileTag & ",1,2,3,4,5,6"
'            .Parent!BRTile.Form.EnableLinks .Parent!BRTile.Form.TileTag & ",1,2,3,4,5,6"
'
'            'disable feature for non-feature parks
'            If TempVars("ParkCode") <> "BLCA" Then
'                .btnLevel2.Caption = Replace(.btnLevel2.Caption, ">", "")
'                .btnLevel3.Caption = ""
'            End If
'        End If
'
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

        .btnLevel0.backColor = bgcolor0
        .btnLevel1.backColor = bgcolor1
        .btnLevel2.backColor = bgcolor2
        .btnLevel3.backColor = bgcolor3

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
