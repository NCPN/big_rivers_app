Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DataEntry = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5160
    DatasheetFontHeight =11
    ItemSuffix =17
    Left =2520
    Top =2400
    Right =19152
    Bottom =11808
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xa3116d04ebbee440
    End
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
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
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =840
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Width =4980
                    Height =300
                    ForeColor =15921906
                    Name ="lblTitle"
                    Caption ="Title"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =300
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =95.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =360
                    Width =5160
                    Height =480
                    BorderColor =10921638
                    Name ="rctHeader"
                    GridlineColor =10921638
                    LayoutCachedTop =360
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =215
                    Left =120
                    Top =480
                    Width =1680
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpecies"
                    Caption ="Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =480
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =720
                End
                Begin Label
                    OverlapFlags =215
                    Left =4080
                    Top =480
                    Width =924
                    Height =252
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCover"
                    Caption ="% Cover"
                    GridlineColor =10921638
                    LayoutCachedLeft =4080
                    LayoutCachedTop =480
                    LayoutCachedWidth =5004
                    LayoutCachedHeight =732
                End
            End
        End
        Begin Section
            Height =420
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4080
                    Top =60
                    Width =960
                    Height =300
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctCover"
                    GridlineColor =10921638

                    LayoutCachedLeft =4080
                    LayoutCachedTop =60
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =2520
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSpecies"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2760
                    Top =60
                    Width =1080
                    Height =300
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCode"
                    GridlineColor =10921638

                    LayoutCachedLeft =2760
                    LayoutCachedTop =60
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =360
                End
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    Width =4020
                    Height =420
                    BorderColor =10921638
                    Name ="rctOverlay"
                    GridlineColor =10921638
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =420
                End
            End
        End
        Begin FormFooter
            Height =360
            Name ="FormFooter"
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
' Form:         SpeciesList
' Level:        Framework form
' Version:      1.00
'
' Description:  Species listing form related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, April 27, 2016
' References:   -
' Revisions:    BLC - 4/27/2016 - 1.00 - initial version
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
' Sub:          XX
' Description:  XX event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 27, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub XX()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - XX[Default form])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 27, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    Dim ary() As String, CoverType As String, speciestype As String
    Dim RiverSegID As Integer, FieldSeason As Integer
    Dim strSQL As String
    Dim rs As dao.Recordset
        
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler
    
    ary() = Split(Me.OpenArgs, "|")

    RiverSegID = CInt(ary(0))
    FieldSeason = CInt(ary(1))
    CoverType = ary(2)
    
    'set title
    Select Case CoverType
        Case "ARS"
            speciestype = "All Rooted "
        Case "URC"
            speciestype = "Understory Rooted "
        Case "WCC"
            speciestype = "Woody Canopy "
    End Select
    
    Me.lblTitle.Caption = speciestype & "Species Cover"
    
    'check if table exists --> if so, delete
    If TableExists("tempSpeciesCover") Then CurrentDb.Execute "DROP TABLE tempSpeciesCover;"
    
    'populate temp table
    strSQL = "SELECT RiverSegment_ID,  CoverType, tlu_NCPN_Plants.LU_Code, tlu_NCPN_Plants.Utah_Species, " _
            & "NULL AS PercentCover " _
            & "INTO tempSpeciesCover " _
            & "FROM ListedSpecies " _
            & "LEFT JOIN tlu_NCPN_Plants ON tlu_NCPN_Plants.LU_Code = ListedSpecies.LU_Code " _
            & "WHERE CoverType = '" & CoverType & "' AND " _
            & "RiverSegment_ID = " & RiverSegID & " AND " _
            & "FieldSeason = " & FieldSeason & " " _
            & "ORDER BY ListedSpecies.LU_Code ASC"

    CurrentDb.Execute strSQL
    
    'setup form source
'    strSQL = "SELECT RiverSegment_ID,  CoverType, tlu_NCPN_Plants.LU_Code, tlu_NCPN_Plants.Utah_Species " _
'            & "FROM ListedSpecies " _
'            & "LEFT JOIN tlu_NCPN_Plants ON tlu_NCPN_Plants.LU_Code = ListedSpecies.LU_Code " _
'            & "WHERE CoverType = '" & CoverType & "' AND " _
'            & "RiverSegment_ID = " & RiverSegID & " AND " _
'            & "FieldSeason = " & FieldSeason & " " _
'            & "ORDER BY ListedSpecies.LU_Code ASC"
    strSQL = "SELECT RiverSegment_ID,  CoverType, LU_Code, Utah_Species, PercentCover " _
            & "FROM tempSpeciesCover;"
    
    'open DAO recordset & assign to form
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    Set Me.Form.Recordset = rs
    
    'assign field values
    Me.tbxCode.ControlSource = "LU_Code"
    Me.tbxSpecies.ControlSource = "Utah_Species"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Tagline form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxTaglineType_Change
' Description:  Tagline type change actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 27, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub cbxTaglineType_Change()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTaglineType_Change[Tagline form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxTaglineType_Click
' Description:  Tagline type click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 27, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub tbxTaglineType_Click()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxTaglineType_Click[Tagline form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxTaglineType_Click
' Description:  Tagline type click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 27, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub cbxTaglineType_Click()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTaglineType_Click[Tagline form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxTaglineType_AfterUpdate
' Description:  Tagline type actions after update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 27, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub cbxTaglineType_AfterUpdate()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTaglineType_AfterUpdate[Tagline form])"
    End Select
    Resume Exit_Handler
End Sub
