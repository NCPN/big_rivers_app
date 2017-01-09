Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =11
    ItemSuffix =70
    Left =3840
    Top =2655
    Right =13920
    Bottom =11085
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x58bf7152d8c6e440
    End
    RecordSource ="tsys_App_Releases"
    Caption ="Big Rivers Database"
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
        Begin Image
            BackStyle =0
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin ToggleButton
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackColor =-1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-1
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverColor =0
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedColor =0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeColor =0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeColor =0
            PressedForeThemeColorIndex =1
        End
        Begin WebBrowser
            OldBorderStyle =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =8445
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =1980
                    Top =120
                    Width =2925
                    Height =585
                    FontSize =22
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="NCPN Big Rivers"
                    GridlineColor =10921638
                    LayoutCachedLeft =1980
                    LayoutCachedTop =120
                    LayoutCachedWidth =4905
                    LayoutCachedHeight =705
                    ForeThemeColorIndex =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =9240
                    Top =7980
                    Width =720
                    ForeColor =4210752
                    Name ="btnNext"
                    Caption ="Next"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Enter Application"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b0482050 ,
                        0xb048200000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e06830ff ,
                        0xb050205000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e06830ff ,
                        0x904820ffa0482040000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e07040ff ,
                        0xd07040ff904820ffb05020400000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e07840ff ,
                        0xe08850ffd05820ff904820ffb050205000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e07850ff ,
                        0xf0a070fff07830ffd05820ff904820ffb0502050000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e08050ff ,
                        0xf0b080ffff9860fff07830ffd05820ffa05830ffd0704060e090600000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e08860ff ,
                        0xffb890ffffa870ffff9860ffd07040ffe0906040e09060000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e09060ff ,
                        0xffc0a0ffffb890ffd07840ffe0906040e0906000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e09870ff ,
                        0xffc0a0ffe08050f0e0906040e090600000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e09870ff ,
                        0xe08860fff0906030e09060000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e09870ff ,
                        0xf0906020e0906000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e0906020 ,
                        0xe090600000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000e0906000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =9240
                    LayoutCachedTop =7980
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =8340
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Image
                    PictureType =2
                    Left =720
                    Top =1200
                    Width =8760
                    Height =6840
                    BorderColor =10921638
                    Name ="imgSplash"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to enter application"
                    Picture ="12346367_834204656702223_7626319009349163900_n"
                    GridlineColor =10921638

                    LayoutCachedLeft =720
                    LayoutCachedTop =1200
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =8040
                    TabIndex =4
                End
                Begin Image
                    PictureType =2
                    Left =120
                    Top =120
                    Width =2100
                    Height =2880
                    BorderColor =10921638
                    Name ="imgNPS"
                    ControlTipText ="Go to NPS website"
                    Picture ="200px_US_NationalParkService_ShadedLogo"
                    HyperlinkAddress ="https://www.nps.gov"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =3000
                    TabIndex =3
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =255
                    Left =900
                    Top =7260
                    Width =3900
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblQuote"
                    Caption ="A river is the report card for its watershed."
                    GridlineColor =10921638
                    LayoutCachedLeft =900
                    LayoutCachedTop =7260
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =7575
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =247
                    Left =1200
                    Top =7560
                    Width =6330
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblAttrib"
                    Caption ="— Alan Levere, Connecticut Department of Environmental Protection"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =7560
                    LayoutCachedWidth =7530
                    LayoutCachedHeight =7875
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1080
                    Top =8070
                    Width =1860
                    Height =315
                    ColumnOrder =0
                    TabIndex =1
                    BorderColor =8355711
                    ForeColor =15527148
                    Name ="tbxVersion"
                    ControlSource ="VersionNumber"
                    GridlineColor =10921638

                    LayoutCachedLeft =1080
                    LayoutCachedTop =8070
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =8385
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =2280
                    Top =780
                    Width =5370
                    Height =315
                    BorderColor =8355711
                    ForeColor =14277081
                    Name ="lblSubTitle"
                    Caption ="Providing park managers with science for decision making."
                    GridlineColor =10921638
                    LayoutCachedLeft =2280
                    LayoutCachedTop =780
                    LayoutCachedWidth =7650
                    LayoutCachedHeight =1095
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =85.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =8070
                    Width =795
                    Height =285
                    BorderColor =8355711
                    ForeColor =14277081
                    Name ="lblVersion"
                    Caption ="Version"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =8070
                    LayoutCachedWidth =975
                    LayoutCachedHeight =8355
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =85.0
                End
                Begin CommandButton
                    Transparent = NotDefault
                    FontUnderline = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =1800
                    Top =60
                    Width =3126
                    Height =726
                    TabIndex =2
                    ForeColor =16711680
                    Name ="btnNCPNBigRivers"
                    Caption ="Command69"
                    ControlTipText ="Go to NCPN Big Rivers Monitoring website"
                    HyperlinkAddress ="http://science.nature.nps.gov/im/units/ncpn/monitor/rivers.cfm"
                    GridlineColor =10921638

                    CursorOnHover =1
                    LayoutCachedLeft =1800
                    LayoutCachedTop =60
                    LayoutCachedWidth =4926
                    LayoutCachedHeight =786
                    ForeThemeColorIndex =10
                    ForeTint =100.0
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =0
            BackColor =4144959
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin FormFooter
            Height =0
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
' Form:         Splash
' Level:        Application form
' Version:      1.02
' Basis:        -
'
' Description:  Splash form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, June 29, 2016
' References:   -
' Revisions:    BLC - 6/29/2016 - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(value As String)

'---------------------
' Properties
'---------------------
Public Property Let Title(value As String)
    If Len(value) > 0 Then
        m_Title = value

        'set the form title & caption
        Me.lblTitle.Caption = m_Title
        Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(value)
    End If
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
'   BLC - 6/27/2016 - adjusted for ToggleForm()
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    lblTitle.Caption = "NCPN Big Rivers"
    lblSubTitle.Caption = "Providing park managers with science for decision making."
    lblQuote.Caption = "A river is the report card for its watershed."
    lblAttrib.Caption = "— Alan Levere, Connecticut Department of Environmental Protection"
    
    btnNext.ForeColor = lngBlue
    
    'set hover
    btnNext.HoverColor = lngGreen
            
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Splash form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          imgSplash_Click
' Description:  Image click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 29, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
' ---------------------------------
Private Sub imgSplash_Click()
On Error GoTo Err_Handler
    
    DoCmd.Close
    DoCmd.OpenForm "User", acNormal
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - imgSplash_Click[Splash form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnNext_Click
' Description:  Next button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 29, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
' ---------------------------------
Private Sub btnNext_Click()
On Error GoTo Err_Handler
    
    DoCmd.Close
    DoCmd.OpenForm "User", acNormal
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNext_Click[Splash form])"
    End Select
    Resume Exit_Handler
End Sub
