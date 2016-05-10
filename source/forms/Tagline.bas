Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DataEntry = NotDefault
    ScrollBars =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6600
    DatasheetFontHeight =11
    ItemSuffix =14
    Left =8208
    Top =3024
    Right =18804
    Bottom =11424
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
                    Width =1980
                    Height =300
                    ForeColor =15921906
                    Name ="lblTitle"
                    Caption ="Taglines"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedWidth =2100
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
                    Width =6600
                    Height =480
                    BorderColor =10921638
                    Name ="rctHeader"
                    GridlineColor =10921638
                    LayoutCachedTop =360
                    LayoutCachedWidth =6600
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
                    Name ="lblTagline"
                    Caption ="Measurement Type"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =480
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =720
                End
                Begin Label
                    OverlapFlags =215
                    Left =2760
                    Top =480
                    Width =924
                    Height =252
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblHeight"
                    Caption ="Height (cm)"
                    GridlineColor =10921638
                    LayoutCachedLeft =2760
                    LayoutCachedTop =480
                    LayoutCachedWidth =3684
                    LayoutCachedHeight =732
                End
                Begin Label
                    OverlapFlags =215
                    Left =3960
                    Top =480
                    Width =1344
                    Height =252
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblDistance"
                    Caption ="Line Distance (m)"
                    GridlineColor =10921638
                    LayoutCachedLeft =3960
                    LayoutCachedTop =480
                    LayoutCachedWidth =5304
                    LayoutCachedHeight =732
                End
            End
        End
        Begin Section
            Height =600
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2760
                    Top =120
                    Width =960
                    Height =300
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxHeight"
                    GridlineColor =10921638

                    LayoutCachedLeft =2760
                    LayoutCachedTop =120
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4080
                    Top =120
                    Width =960
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDistance"
                    GridlineColor =10921638

                    LayoutCachedLeft =4080
                    LayoutCachedTop =120
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =420
                End
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =120
                    Top =120
                    Width =2520
                    Height =300
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="cbxTaglineType"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5280
                    Top =120
                    Width =576
                    Height =300
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnAdd"
                    OnClick ="[Event Procedure]"
                    OnMouseDown ="[Event Procedure]"
                    ControlTipText ="Add"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b09880ff201010ff201010ff201010ff201010ff201010ff ,
                        0x201010ff201010ff201810ff201810ff201810ff201810ff201810ff00000000 ,
                        0x0000000000000000c0a090fffff8f0fffff8f0fffff0f0fffff0e0fff0e8e0ff ,
                        0xf0e8d0fff0e0d0fff0e0d0fff0e0d0fff0e0d0fff0e0d0ff403830ff00000000 ,
                        0x0000000000000000c0a090ffffffffffd07850ffd07840ffd07040ffc07040ff ,
                        0xc07040ffc07850ffd09070ffd0a890ffd0a890fff0f0f0ff909090ff00000000 ,
                        0x0000000000000000c0a890ffffffffffd07850fff0b8a0fff0b090fff0a880ff ,
                        0xf0a880fff0b090ffe0b0a0ff804040ff703840ff703840ff703840ff703840ff ,
                        0x703840ff703840ffc0a890ffffffffffd07850ffd07850ffd07840ffd07040ff ,
                        0xd08050ffe0a890ffa05850ffc07870ff604840ffd0d8d0ffd0d8d0ff605040ff ,
                        0xc06060ff703840ffc0a8a0fffffffffffffffffffffffffffffffffffff8f0ff ,
                        0xfff8f0fffff8f0ffb06060ffe09090ff605040ff605040ff605040ff605040ff ,
                        0xc07070ff703840ffc0a8a0ffc0a8a0ffc0a890ffc0a090ffc0a090ffc0a090ff ,
                        0xc0a8a0ffe0d0c0ffc07070fff0a8b0ffe0a0a0ffe098a0ffe09090ffe08890ff ,
                        0xd08080ff703840ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000d08080ffd07070ffd06860ffd06860ffc05850ffc05850ff ,
                        0xb05040ff804040ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000d08890ffe07070ffffffffffffffffffffffffffffffffff ,
                        0xc05850ff904850ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000d09090ffe07070ffffffffffffffffffffffffffffffffff ,
                        0xd06860ffa05860ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000e0a0a0ffd09090ffd08890ffd08080ffc07070ffc06870ff ,
                        0xc06870ffc06860ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5280
                    LayoutCachedTop =120
                    LayoutCachedWidth =5856
                    LayoutCachedHeight =420
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5940
                    Top =120
                    Width =576
                    Height =300
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnDelete"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b0a090ff302010ff302010ff302010ff302010ff302010ff ,
                        0x302010ff302010ff302010ff302010ff302010ff302010ff302010ff00000000 ,
                        0x0000000000000000b0a090fffff8f0fffff0f0ffffe8e0fff0e8e0fff0e0d0ff ,
                        0xf0d8d0fff0d8c0fff0d8c0fff0d8c0fff0d8c0fff0d8c0ff302010ff00000000 ,
                        0x0000000000000000b0a090ffffffffffe06830ffe06830ffe06830ffd06830ff ,
                        0xd06830ffd06830ffd06030ffc06030ff904820ffffe0d0ff302010ff00000000 ,
                        0x0000000000000000b0a090ffffffffffd06830ffffb080ffffa880ffffa070ff ,
                        0xf09870fff09060ffa0b0f0ff1020e0ffc0c8f0ffffe0d0ff302010ff00000000 ,
                        0x00000000a0a8f0ffb0a090ffffffffffe06830ffe06830ffe06830ffd06830ff ,
                        0xd06830ffe0e0f0ff0028ffff1028f0ff4050d0ffffe0d0ff302010ff00000000 ,
                        0x4050e0ff0010b0ffb0a090ffffffffffffffffffffffffffffffffffffffffff ,
                        0xfff8f0ffffe8e0ff2048ffff1038ffff1028ffffe0e8f0ff302010ff7088f0ff ,
                        0x0018c0ff6078f0ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ff ,
                        0xb0a090ffb0a090ffe0e0f0ff3050ffff2040ffff8090f0ffb0b8f0ff0028f0ff ,
                        0x4058f0ff00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000d0d8f0ff4060ffff3050ffff2040ffff3050ffff ,
                        0xe0e8f0ff00000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c0d0f0ff4068ffff4060ffffc0c8f0ff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000c0c8f0ff6078ffff6078ffff6080ffff5070ffff ,
                        0xe0e0f0ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b0b8f0ff6078ffff6078ffffb0c0f0fff0f0f0ff7088ffff ,
                        0x6078ffffc0d0f0ff000000000000000000000000000000000000000000000000 ,
                        0x0000000090a0ffff6078ffff6078ffffd0d8f0ff000000000000000000000000 ,
                        0xb0b8f0ff8098ffff000000000000000000000000000000000000000000000000 ,
                        0x000000008098ffff6080ffffd0d8f0ff00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5940
                    LayoutCachedTop =120
                    LayoutCachedWidth =6516
                    LayoutCachedHeight =420
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =120
                    Top =180
                    Width =2520
                    Height =300
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTaglineType"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =180
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =480
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
' Form:         Tagline
' Level:        Framework form
' Version:      1.00
'
' Description:  Tagline form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, April 26, 2016
' References:   -
' Revisions:    BLC - 4/26/2016 - 1.00 - initial version
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
' Source/date:  Bonnie Campbell, April 26, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/26/2016 - initial version
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
' Source/date:  Bonnie Campbell, April 26, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/26/2016 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    Dim ary() As String, SourceType As String
    Dim SourceID As Integer
    Dim strSQL As String, strSQL2 As String
    Dim rs As DAO.Recordset
        
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler
    
    ary() = Split(Me.OpenArgs, "|")

    SourceType = ary(0)
    SourceID = CInt(ary(1))
    
    'setup form source
    strSQL = "SELECT * FROM Tagline " _
            & "WHERE LineDistanceSource = '" & SourceType & "' AND " _
            & "LineDistanceSource_ID = " & SourceID & ""
    
    'open DAO recordset & assign to form
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    Set Me.Form.Recordset = rs
    
    'assign field values
    Me.tbxDistance.ControlSource = "LineDistance_m"
    Me.tbxHeight.ControlSource = "Height_cm"
    Me.tbxTaglineType.ControlSource = "TaglineType"

'    Me.Requery
    
    strSQL2 = "SELECT ID, Label, Summary FROM Enum WHERE EnumType = 'TaglineType';"

    'setup tagline types w/ bound column = ID
    With Me.cbxTaglineType
        .RowSource = strSQL2
        .ColumnCount = 3
        .ColumnWidths = "0, 0, 2"
        .BoundColumn = "1"

'        'set value of tagline type
'        Dim i As Integer
'
'        For i = 0 To (.ListCount - 1)
'            If .Column(1, i) = Me!HeightType Then
'                cbxTaglineType = .ItemData(i)
'            End If
'        Next

    End With

    'setup button face

    
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
' Source/date:  Bonnie Campbell, April 26, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/26/2016 - initial version
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
' Source/date:  Bonnie Campbell, April 26, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/26/2016 - initial version
' ---------------------------------
Private Sub tbxTaglineType_Click()
On Error GoTo Err_Handler

'        'set value of tagline type
'        Dim i As Integer
'
'        For i = 0 To (.ListCount - 1)
'            If .Column(1, i) = Me!HeightType Then
'                cbxTaglineType = .ItemData(i)
'            End If
'        Next


    'bring selection box to front
    cbxTaglineType.SetFocus

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
' Source/date:  Bonnie Campbell, April 26, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/26/2016 - initial version
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
' Source/date:  Bonnie Campbell, April 26, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/26/2016 - initial version
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
