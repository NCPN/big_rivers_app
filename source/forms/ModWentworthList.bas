Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7860
    DatasheetFontHeight =11
    ItemSuffix =37
    Right =14235
    Bottom =11850
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x21398585f5d2e440
    End
    RecordSource ="ModWentworthScale"
    Caption ="_List"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            Height =1515
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    GridlineColor =10921638
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =180
                    Top =120
                    Width =5880
                    Height =585
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Edit or Delete Records using the buttons for the record at right.\015\012Icon co"
                        "des at left identify if record may be edited/deleted."
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =705
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1350
                    Top =1200
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblClass"
                    Caption ="Class"
                    GridlineColor =10921638
                    LayoutCachedLeft =1350
                    LayoutCachedTop =1200
                    LayoutCachedWidth =2595
                    LayoutCachedHeight =1515
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =960
                    Top =1200
                    Width =270
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblHdrID"
                    Caption ="ID"
                    GridlineColor =10921638
                    LayoutCachedLeft =960
                    LayoutCachedTop =1200
                    LayoutCachedWidth =1230
                    LayoutCachedHeight =1515
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =2700
                    Top =1200
                    Width =1380
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblCode"
                    Caption ="Code"
                    GridlineColor =10921638
                    LayoutCachedLeft =2700
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1515
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4020
                    Top =1200
                    Width =1140
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDiameter"
                    Caption ="Diameter"
                    GridlineColor =10921638
                    LayoutCachedLeft =4020
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =1515
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
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
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =2
                    Left =7080
                    Top =30
                    Width =720
                    Height =330
                    FontSize =14
                    ForeColor =255
                    Name ="btnDelete"
                    Caption ="í ½í·´"
                    OnClick ="[Event Procedure]"
                    FontName ="Academy Engraved LET"
                    ControlTipText ="Delete record"
                    GridlineColor =10921638

                    LayoutCachedLeft =7080
                    LayoutCachedTop =30
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =360
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =-1
                    BackColor =14136213
                    BorderColor =14136213
                    ThemeFontIndex =-1
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
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =45
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =45
                    LayoutCachedWidth =840
                    LayoutCachedHeight =345
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1320
                    Top =45
                    Width =1500
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxClass"
                    ControlSource ="Label"
                    ConditionalFormat = Begin
                        0x01000000f6000000020000000100000000000000000000002300000001000000 ,
                        0xed1c2400ffffff000100000000000000240000004a0000000100000022b14c00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b0052006500740069007200650059006500610072005d00 ,
                        0x3c005900650061007200280044006100740065002800290029002c0031002c00 ,
                        0x300029000000000049004900660028005b007400620078005200650074006900 ,
                        0x7200650059006500610072005d003e0059006500610072002800440061007400 ,
                        0x65002800290029002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1320
                    LayoutCachedTop =45
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =345
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010002000000010000000000000001000000ed1c2400ffffff00220000004900 ,
                        0x4900660028005b0052006500740069007200650059006500610072005d003c00 ,
                        0x5900650061007200280044006100740065002800290029002c0031002c003000 ,
                        0x2900000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000022b14c00ffffff002500000049004900660028005b0074006200780052 ,
                        0x006500740069007200650059006500610072005d003e00590065006100720028 ,
                        0x0044006100740065002800290029002c0031002c003000290000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =900
                    Top =30
                    Width =360
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =900
                    LayoutCachedTop =30
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2880
                    Top =45
                    Width =1020
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxCode"
                    ControlSource ="Code"
                    ConditionalFormat = Begin
                        0x01000000f6000000020000000100000000000000000000002300000001000000 ,
                        0xff000000ffffff000100000000000000240000004a0000000100000022b14c00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b0052006500740069007200650059006500610072005d00 ,
                        0x3c005900650061007200280044006100740065002800290029002c0031002c00 ,
                        0x300029000000000049004900660028005b007400620078005200650074006900 ,
                        0x7200650059006500610072005d003e0059006500610072002800440061007400 ,
                        0x65002800290029002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =45
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =345
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010002000000010000000000000001000000ff000000ffffff00220000004900 ,
                        0x4900660028005b0052006500740069007200650059006500610072005d003c00 ,
                        0x5900650061007200280044006100740065002800290029002c0031002c003000 ,
                        0x2900000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000022b14c00ffffff002500000049004900660028005b0074006200780052 ,
                        0x006500740069007200650059006500610072005d003e00590065006100720028 ,
                        0x0044006100740065002800290029002c0031002c003000290000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5520
                    Top =30
                    Width =720
                    Height =330
                    TabIndex =5
                    ForeColor =4210752
                    Name ="btnActive"
                    OnClick ="[Event Procedure]"
                    FontName ="Lucida Sans Unicode"
                    ControlTipText ="Toggle retired / retire date"
                    GridlineColor =10921638

                    LayoutCachedLeft =5520
                    LayoutCachedTop =30
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =360
                    BackColor =14136213
                    BorderColor =14136213
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3960
                    Top =45
                    Width =1500
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxDiameter"
                    ControlSource ="=Replace([DiameterRange_mm],\"micro\",\"\")"
                    ConditionalFormat = Begin
                        0x01000000f6000000020000000100000000000000000000002300000001000000 ,
                        0xff000000ffffff000100000000000000240000004a0000000100000022b14c00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b0052006500740069007200650059006500610072005d00 ,
                        0x3c005900650061007200280044006100740065002800290029002c0031002c00 ,
                        0x300029000000000049004900660028005b007400620078005200650074006900 ,
                        0x7200650059006500610072005d003e0059006500610072002800440061007400 ,
                        0x65002800290029002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3960
                    LayoutCachedTop =45
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =345
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010002000000010000000000000001000000ff000000ffffff00220000004900 ,
                        0x4900660028005b0052006500740069007200650059006500610072005d003c00 ,
                        0x5900650061007200280044006100740065002800290029002c0031002c003000 ,
                        0x2900000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000022b14c00ffffff002500000049004900660028005b0074006200780052 ,
                        0x006500740069007200650059006500610072005d003e00590065006100720028 ,
                        0x0044006100740065002800290029002c0031002c003000290000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2580
                    Width =1020
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxRetireYear"
                    ControlSource ="RetireYear"
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1560
                    Top =15
                    Width =1020
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxActiveYear"
                    ControlSource ="ActiveYear"
                    GridlineColor =10921638

                    LayoutCachedLeft =1560
                    LayoutCachedTop =15
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6300
                    Top =30
                    Width =720
                    Height =330
                    TabIndex =9
                    ForeColor =4210752
                    Name ="btnEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Edit record"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000303840ff404040ff505050ff504850f080686020 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000606060ff909890ffd0d0d0ffa0a8b0ff304850ff ,
                        0xa090905000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000a0a0a0fff0f0f0fff0f8ffffc0e0f0ff5090b0ff ,
                        0x204850ff80686020000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000080787080e0e0e0ffd0f0f0ff90e0f0ff50c0d0ff ,
                        0x4098b0ff204850ff806860200000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000006090a080c0e8f0ffa0f0f0ff70e0f0ff ,
                        0x50c0d0ff4098b0ff204850ff8068602000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000006090a090b0e8f0ffa0f0f0ff ,
                        0x70e0f0ff50c0d0ff4098b0ff204850ff80686020000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000006090a090b0e8f0ff ,
                        0xa0f0f0ff70e0f0ff50c0d0ff4098b0ff204850ff806860200000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006090a0a0 ,
                        0xb0e8f0ffa0f0f0ff70e0f0ff50c0d0ff4098b0ff204850ff8068602000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x6090a0a0b0e8f0ffa0f0f0ff70e0f0ff50c0d0ff4098b0ff204850ff80686020 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xd08060006090a0a0b0e8f0ffa0f0f0ff70e0f0ff50b8d0ff4098b0ff204850ff ,
                        0x8068602000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000d0d8e0006090a0b0b0e8f0ffa0f0f0ff70d0e0ff50a0b0ff808890ff ,
                        0x303870ff80686020000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000d0d8e0006090a0b0c0f0f0ffa0e0e0ffb0b0a0ff5058b0ff ,
                        0x303090ff505880ff000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000d0d8e0006090a0b0a0b8d0ff8088d0ff6070d0ff ,
                        0x303090ff202860ff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000d0d8e0006070b0b09098d0ff7078d0ff ,
                        0x4050a0ff9098b0ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000d0d8e000606090d05060a0ff ,
                        0x9090b0ff00000000
                    End

                    LayoutCachedLeft =6300
                    LayoutCachedTop =30
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =360
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
            End
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
' Form:         ModWentworthList
' Level:        Application form
' Version:      1.01
' Basis:        Dropdown form
'
' Description:  List form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
'               BLC - 10/20/2016 - 1.01 - removed ButtonCaption, SelectedID, SelectedValue properties
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let Title(Value As String)
    If Len(Value) > 0 Then
        m_Title = Value

        'set the form title & caption
        Me.lblTitle.Caption = m_Title
        Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(Value)
    End If
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Directions(Value As String)
    If Len(Value) > 0 Then
        m_Directions = Value

        'set the form directions
        Me.lblDirections.Caption = m_Directions
    Else
        RaiseEvent InvalidDirections(Value)
    End If
End Property

Public Property Get Directions() As String
    Directions = m_Directions
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
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    lblTitle.Caption = ""
    lblDirections.Caption = "Edit or Delete Records using the buttons for the record at right." _
                            & vbCrLf & "Icon codes at left identify if record may be edited/deleted."
    tbxIcon.Value = StringFromCodepoint(uLocked)
    tbxIcon.ForeColor = lngDkGreen
    lblDirections.ForeColor = lngLtBlue
    
    lblDiameter.Caption = "Diameter"
    
    'set hover
    btnEdit.HoverColor = lngGreen
    btnDelete.HoverColor = lngGreen
    btnActive.HoverColor = lngGreen
    
    btnDelete.Caption = StringFromCodepoint(uDelete)
    btnDelete.ForeColor = lngRed

    btnActive.Caption = StringFromCodepoint(uAlarmClock)
    btnActive.ForeColor = lngRed

    'set data source
    Me.RecordSource = "ModWentworthScale"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[ModWentworthList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'eliminate NULLs
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[ModWentworthList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[ModWentworthList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnActive_Click
' Description:  Active button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 10/20/2016 - revised to toggle retire date
' ---------------------------------
Private Sub btnActive_Click()
On Error GoTo Err_Handler
    
    Dim IsRetired As Byte
    
    'default
    IsRetired = 0
    
    Select Case Nz(tbxRetireYear, Year(Date) + 100)
        Case Is < Year(Date)   'past retire date
            IsRetired = 1
        Case Is = Year(Date) 'current year retire date
            IsRetired = 1
        Case Is > Year(Date) 'future retire date
            'default
    End Select
    
    'identify the record ID
    ToggleIsActive "ModWentworthScale", tbxID, IsRetired 'Abs(tbxIsActive.Value - 1)

    'clear the Activated record
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSensitive_Click[ModWentworthList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnEdit_Click
' Description:  Enter button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub btnEdit_Click()
On Error GoTo Err_Handler
    
    'populate the parent form
    PopulateForm Me.Parent, ID

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEdit_Click[ModWentworthList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnDelete_Click
' Description:  Delete button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub btnDelete_Click()
On Error GoTo Err_Handler
    
    Dim result As Integer
    
    'identify the record ID
     result = MsgBox("Delete Record this record: #" & tbxID & " ?" _
                        & vbCrLf & "This action cannot be undone.", vbYesNo, "Delete Record?")

    If result = vbYes Then DeleteRecord "Location", ID
    
    'clear the deleted record
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[ModWentworthList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Close
' Description:  form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[ModWentworthList form])"
    End Select
    Resume Exit_Handler
End Sub
