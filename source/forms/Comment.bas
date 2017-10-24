Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9660
    DatasheetFontHeight =11
    ItemSuffix =18
    Left =3375
    Top =2070
    Right =13035
    Bottom =10725
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xde3980067302e540
    End
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
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
            Height =447
            BackColor =13107
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
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="title"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =360
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =85
                    Top =432
                    Width =9660
                    BorderColor =6750105
                    Name ="lineIndicator"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedTop =432
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =432
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    Left =7380
                    Top =60
                    Width =1980
                    Height =300
                    ForeColor =12566463
                    Name ="lblContext"
                    Caption ="context"
                    GridlineColor =10921638
                    LayoutCachedLeft =7380
                    LayoutCachedTop =60
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =360
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =8220
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =60
                    Width =9120
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblInstructions"
                    Caption ="instructions"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =300
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8580
                    Top =2520
                    Width =780
                    ForeColor =4210752
                    Name ="btnCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000f0906060d0784080b0583010000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000e0785040f08850ffd07040ffa05830500000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000f0906020d0704060f08050ffd07050f0a050300000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c06840d0f08850ffc078508000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0c0b01000000000000000000000000090482040e07840ffe08860ffe0a08000 ,
                        0x00000000000000000000000000000000d07040ffd07040ffc06840ffb06030ff ,
                        0xb05830ff905030ff0000000000000000b0603020c06840ffe08050ffd0886080 ,
                        0x00000000000000000000000000000000d07850ffe07030fff08050fff09870ff ,
                        0xe09060fff0a08040000000000000000080402000c06840ffe07840f0e09870c0 ,
                        0x00000000000000000000000000000000d08050ffe08050fff09060fff0a070ff ,
                        0x904830b0b0603040000000000000000080402000c06840ffd07040f0e09870d0 ,
                        0x00000000000000000000000000000000d08860ffe09060fff09870fff08850f0 ,
                        0xb06040ffb06040ffb060307000000000b0805020a05830f0d07840f0e09070d0 ,
                        0x000000000000000000000000e0b09010c08060ffd09870e0d0886090d09070ff ,
                        0xd08050ffc07040ffc06840ffb06030c0b07040e0a06040ffe08050ffd0a080e0 ,
                        0x00000000000000000000000000000000c08860ffd0a0804000000000d08860c0 ,
                        0xd08860ffd08050f0c06840ffb06840ffb06030f0e07840f0e0a080f0d09880e0 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0a880c0e09880ffe09870f0e09070f0e09070e0e0a080f0e0a890f0f0b8a020 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000f0b89060f0b090c0f0b8a0e0f0c0a0c0f0c0a090f0c0b02000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =8580
                    LayoutCachedTop =2520
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =2880
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =7740
                    Top =2520
                    Width =720
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnAdd"
                    Caption ="Add"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000d0687050c06860ffb05850ffa05050ffa05050ff ,
                        0xa05050ff904850ff904840ff904840ff804040ff803840ff803840ff703840ff ,
                        0x703830ff0000000000000000d06870fff09090ffe08080ffb04820ff403020ff ,
                        0xc0b8b0ffc0b8b0ffd0c0c0ffd0c8c0ff505050ffa04030ffa04030ffa03830ff ,
                        0x703840ff0000000000000000d07070ffff98a0fff08880ffe08080ff705850ff ,
                        0x404030ff907870fff0e0e0fff0e8e0ff908070ffa04030ffa04040ffa04030ff ,
                        0x803840ff0000000000000000d07870ffffa0a0fff09090fff08880ff705850ff ,
                        0x000000ff404030fff0d8d0fff0e0d0ff807860ffb04840ffb04840ffa04040ff ,
                        0x804040ff0000000000000000d07880ffffa8b0ffffa0a0fff09090ff705850ff ,
                        0x705850ff705850ff705850ff706050ff806860ffc05850ffb05050ffb04840ff ,
                        0x804040ff0000000000000000e08080ffffb0b0ffffb0b0ffffa0a0fff09090ff ,
                        0xf08880ffe08080ffe07880ffd07070ffd06870ffc06060ffc05850ffb05050ff ,
                        0x904840ff0000000000000000e08890ffffb8c0ffffb8b0ffd06060ffc06050ff ,
                        0xc05850ffc05040ffb05030ffb04830ffa04020ffa03810ffc06060ffc05850ff ,
                        0x904840ff0000000000000000e09090ffffc0c0ffd06860ffffffffffffffffff ,
                        0xfff8f0fff0f0f0fff0e8e0fff0d8d0ffe0d0c0ffe0c8c0ffa03810ffc06060ff ,
                        0x904850ff0000000000000000e098a0ffffc0c0ffd07070ffffffffffffffffff ,
                        0xfffffffffff8f0fff0f0f0fff0e8e0fff0d8d0ffe0d0c0ffa04020ffd06860ff ,
                        0xa05050ff0000000000000000f0a0a0ffffc0c0ffe07870ffffffffffffffffff ,
                        0xfffffffffffffffffff8f0fff0f0f0fff0e8e0fff0d8d0ffb04830ffd07070ff ,
                        0xa05050ff0000000000000000f0a8a0ffffc0c0ffe08080ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffff8f0fff0f0f0fff0e8e0ffb05030ffe07880ff ,
                        0xa05050ff0000000000000000f0b0b0ffffc0c0fff08890ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffff8f0fff0f0f0ffc05040ff603030ff ,
                        0xb05850ff0000000000000000f0b0b0ffffc0c0ffff9090ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffff8f0ffc05850ffb05860ff ,
                        0xb05860ff0000000000000000f0b8b0fff0b8b0fff0b0b0fff0b0b0fff0a8b0ff ,
                        0xf0a0a0ffe098a0ffe09090ffe09090ffe08890ffe08080ffd07880ffd07870ff ,
                        0xd07070ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =7740
                    LayoutCachedTop =2520
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =2880
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
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =180
                    Top =720
                    Width =8940
                    Height =1680
                    TabIndex =2
                    BackColor =15921906
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxComment"
                    OnKeyPress ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =180
                    LayoutCachedTop =720
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =2400
                    BackShade =95.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =5760
                    Top =420
                    Width =1380
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCharacterCount"
                    Caption ="Character Count"
                    GridlineColor =10921638
                    LayoutCachedLeft =5760
                    LayoutCachedTop =420
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =660
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =7200
                    Top =420
                    Width =660
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    Name ="lblCount"
                    GridlineColor =10921638
                    LayoutCachedLeft =7200
                    LayoutCachedTop =420
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =-1
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =7800
                    Top =360
                    Width =1500
                    Height =360
                    BorderColor =10921638
                    Name ="rctAlert"
                    GridlineColor =10921638
                    LayoutCachedLeft =7800
                    LayoutCachedTop =360
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =720
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =7860
                    Top =420
                    Width =1380
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    Name ="lblMaxCount"
                    Caption ="maxcount"
                    GridlineColor =10921638
                    LayoutCachedLeft =7860
                    LayoutCachedTop =420
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =-1
                End
                Begin Subform
                    CanShrink = NotDefault
                    OverlapFlags =215
                    Left =105
                    Top =3720
                    Width =9435
                    Height =4380
                    TabIndex =3
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.CommentList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =3720
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =8100
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =3600
                    Width =9660
                    Height =4620
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =3600
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =8220
                    BackThemeColorIndex =-1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =3360
                    Width =9660
                    Height =315
                    FontSize =9
                    LeftMargin =360
                    TopMargin =36
                    RightMargin =360
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =65535
                    Name ="lblMsg"
                    Caption ="msg"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedTop =3360
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =3675
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =5760
                    Top =3180
                    Width =825
                    Height =600
                    FontSize =20
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =65535
                    Name ="lblMsgIcon"
                    Caption ="icon"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =5760
                    LayoutCachedTop =3180
                    LayoutCachedWidth =6585
                    LayoutCachedHeight =3780
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =2580
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxID"
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =2580
                    LayoutCachedWidth =480
                    LayoutCachedHeight =2880
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
            End
        End
        Begin FormFooter
            Visible = NotDefault
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
' Form:         Comment
' Level:        Framework form
' Version:      1.08
'
' Description:  Comment form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
'               BLC - 8/9/2016  - 1.01 - revised Comment to AppComment (comment reserved word)
'               BLC - 12/5/2016 - 1.02 - added instruction & max count
'               BLC - 9/25/2017 - 1.03 - revise for NCPN_framework.XX classes
'               BLC - 9/27/2017 - 1.04 - update to use Factory.NewClassXX() vs GetClass()
'               BLC - 10/16/2017 - 1.05 - remove Initialize() & Terminate()
'               BLC - 10/17/2017 - 1.06 - added calling form info
'               BLC - 10/18/2017 - 1.07 - added Form_Open() & Form_Close(), shifted calling form info,
'                                         set filter on comment list subform from Form_Load()
'               BLC 6- 10/19/2017 - 1.08 - added comment message
' =================================

'---------------------
' Declarations
'---------------------
Private m_oComment As AppComment

Private m_Title As String
Private m_Context As String
Private m_Instructions As String
Private m_CountLabel As String
Private m_CurrentCount As String
Private m_MaxCount As String
Private m_AlertCount As Integer
Private m_RemainingCount As String
Private m_Comment As String

Private m_CommentHeaderColor As Long
Private m_TitleFontColor As Long
Private m_InstructionFontColor As Long
Private m_CountLabelFontColor As Long
Private m_CurrentCountFontColor As Long
Private m_MaxCountFontColor As Long
Private m_RemainingCountFontColor As Long
Private m_AlertBoxBackgroundColor As Long

Private m_CommentVisible As Byte
Private m_ContextVisible As Byte
Private m_InstructionVisible As Byte
Private m_CountLabelVisible As Byte
Private m_CurrentCountVisible As Byte
Private m_MaxCountVisible As Byte
Private m_RemainingCountVisible As Byte
Private m_AlertCountVisible As Byte
Private m_AlertBoxVisible As Byte

Private m_AddButtonText As String
Private m_AddButtonForeColor As Long
Private m_AddButtonColor As Long

Private m_CancelButtonText As String
Private m_CancelButtonForeColor As Long
Private m_CancelButtonColor As Long

Private m_AddButtonVisible As Byte
Private m_CancelButtonVisible As Byte

Private m_AddAction As String
Private m_CancelAction As String
Private m_EditAction As String

Private m_CallingForm As String

'---------------------
' Event Declarations
'---------------------
Public Event Initialize()
Public Event Terminate()
Public Event InvalidCallingForm(Value As String)

'---------------------
' Properties
'---------------------

' ==== Values ====
Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Form Title"
    If ValidateString(Value, "alphanumdash") Then
        m_Title = Value
    End If
    lblTitle.Caption = m_Title
End Property

Public Property Get Context() As String
    Context = m_Context
End Property

Public Property Let Context(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Context"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_Context = Value
    End If
    lblContext.Caption = m_Context
End Property

Public Property Get Instructions() As String
    Instructions = m_Instructions
End Property

Public Property Let Instructions(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Instructions"
    If ValidateString(Value, "paragraph") Then
        m_Instructions = Value
    End If
    lblInstructions.Caption = m_Instructions
End Property

Public Property Get CountLabel() As String
    CountLabel = m_CountLabel
End Property

Public Property Let CountLabel(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Character Count"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_CountLabel = Value
    End If
    lblCharacterCount.Caption = m_CountLabel
End Property

Public Property Get CurrentCount() As String
    CurrentCount = m_CurrentCount
End Property

Public Property Let CurrentCount(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "1"
    If ValidateString(Value, "numeric") Then
        m_CurrentCount = Value
    End If
    lblCount.Caption = m_CurrentCount
End Property

Public Property Get MaxCount() As String
    MaxCount = m_MaxCount
End Property

Public Property Let MaxCount(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "/ XX characters"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_MaxCount = Value
    End If
    lblMaxCount.Caption = m_MaxCount
End Property

'set the value at which the count display changes color
Public Property Get AlertCount() As Integer
    AlertCount = m_AlertCount
End Property

Public Property Let AlertCount(Value As Integer)
    If Len(Trim(Value)) = 0 Then Value = 10
    m_AlertCount = Value
End Property

Public Property Get RemainingCount() As String
    RemainingCount = m_RemainingCount
End Property

Public Property Let RemainingCount(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "XX characters remain"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_RemainingCount = Value
    End If
    lblMaxCount.Caption = m_RemainingCount
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let Comment(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Comment"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_Comment = Value
    End If
    tbxComment.Value = m_Comment
End Property

' ==== Color ====
Public Property Get CommentHeaderColor() As Long
    CommentHeaderColor = m_CommentHeaderColor
End Property

Public Property Let CommentHeaderColor(Value As Long)
    m_CommentHeaderColor = Value
    FormHeader.BackColor = m_CommentHeaderColor
End Property

Public Property Get TitleFontColor() As Long
    TitleFontColor = m_TitleFontColor
End Property

Public Property Let TitleFontColor(Value As Long)
    m_TitleFontColor = Value
    lblTitle.ForeColor = m_TitleFontColor
End Property

Public Property Get InstructionFontColor() As Long
    InstructionFontColor = m_InstructionFontColor
End Property

Public Property Let InstructionFontColor(Value As Long)
    m_InstructionFontColor = Value
    lblInstructions.ForeColor = m_InstructionFontColor
End Property

Public Property Get CountLabelFontColor() As Long
    CountLabelFontColor = m_CountLabelFontColor
End Property

Public Property Let CountLabelFontColor(Value As Long)
    m_CountLabelFontColor = Value
    lblCount.ForeColor = m_CountLabelFontColor
End Property

Public Property Get CurrentCountFontColor() As Long
    CurrentCountFontColor = m_CurrentCountFontColor
End Property

Public Property Let CurrentCountFontColor(Value As Long)
    m_CurrentCountFontColor = Value
    lblCount.ForeColor = m_CurrentCountFontColor
End Property

Public Property Get MaxCountFontColor() As Long
    MaxCountFontColor = m_MaxCountFontColor
End Property

Public Property Let MaxCountFontColor(Value As Long)
    m_MaxCountFontColor = Value
    lblMaxCount.ForeColor = m_MaxCountFontColor
End Property

Public Property Get RemainingCountFontColor() As Long
    RemainingCountFontColor = m_RemainingCountFontColor
End Property

Public Property Let RemainingCountFontColor(Value As Long)
    m_RemainingCountFontColor = Value
    lblMaxCount.ForeColor = m_RemainingCountFontColor
End Property

Public Property Get AlertBoxBackgroundColor() As Long
    AlertBoxBackgroundColor = m_AlertBoxBackgroundColor
End Property

Public Property Let AlertBoxBackgroundColor(Value As Long)
    rctAlert.backstyle = 1 '1 = Normal, 0 = Transparent
    m_AlertBoxBackgroundColor = Value
    rctAlert.BackColor = m_AlertBoxBackgroundColor
End Property

' ==== Visibility ====
Public Property Get CommentVisible() As Byte
    CommentVisible = m_CommentVisible
End Property

Public Property Let CommentVisible(Value As Byte)
    m_CommentVisible = Value
    tbxComment.Visible = m_CommentVisible
End Property

Public Property Get InstructionVisible() As Byte
    InstructionVisible = m_InstructionVisible
End Property

Public Property Let InstructionVisible(Value As Byte)
    m_InstructionVisible = Value
    lblInstructions.Visible = m_InstructionVisible
End Property

Public Property Get CountLabelVisible() As Byte
    CountLabelVisible = m_CountLabelVisible
End Property

Public Property Let CountLabelVisible(Value As Byte)
    m_CountLabelVisible = Value
    lblCount.Visible = m_CountLabelVisible
End Property

Public Property Get CurrentCountVisible() As Byte
    CurrentCountVisible = m_CurrentCountVisible
End Property

Public Property Let CurrentCountVisible(Value As Byte)
    m_CurrentCountVisible = Value
    lblCount.Visible = m_CurrentCountVisible
End Property

Public Property Get MaxCountVisible() As Byte
    MaxCountVisible = m_MaxCountVisible
End Property

Public Property Let MaxCountVisible(Value As Byte)
    m_MaxCountVisible = Value
    lblMaxCount.Visible = m_MaxCountVisible
End Property

Public Property Get RemainingCountVisible() As Byte
    RemainingCountVisible = m_RemainingCountVisible
End Property

Public Property Let RemainingCountVisible(Value As Byte)
    m_RemainingCountVisible = Value
End Property

Public Property Get AlertBoxVisible() As Byte
    AlertBoxVisible = m_AlertBoxVisible
End Property

Public Property Let AlertBoxVisible(Value As Byte)
    m_AlertBoxVisible = Value
    Me.rctAlert.Visible = m_AlertBoxVisible
End Property

' ==== Buttons ====
Public Property Get AddButtonText() As String
    AddButtonText = m_AddButtonText
End Property

Public Property Let AddButtonText(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Add"
    If ValidateString(Value, "alphaspace") Then
        m_AddButtonText = Value
    End If
    btnAdd.Caption = m_AddButtonText
End Property

Public Property Get CancelButtonText() As String
    CancelButtonText = m_CancelButtonText
End Property

Public Property Let CancelButtonText(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Cancel"
    If ValidateString(Value, "alphaspace") Then
        m_CancelButtonText = Value
    End If
    btnCancel.Caption = m_CancelButtonText
End Property

Public Property Get AddButtonForeColor() As Long
    AddButtonForeColor = m_AddButtonForeColor
End Property

Public Property Let AddButtonForeColor(Value As Long)
    m_AddButtonForeColor = Value
    btnAdd.ForeColor = m_AddButtonForeColor
End Property

Public Property Get AddButtonColor() As Long
    AddButtonColor = m_AddButtonColor
End Property

Public Property Let AddButtonColor(Value As Long)
    m_AddButtonColor = Value
    btnAdd.BackColor = m_AddButtonColor
End Property

Public Property Get CancelButtonForeColor() As Long
    CancelButtonForeColor = m_CancelButtonForeColor
End Property

Public Property Let CancelButtonForeColor(Value As Long)
    m_CancelButtonForeColor = Value
    btnCancel.ForeColor = m_CancelButtonForeColor
End Property

Public Property Get CancelButtonColor() As Long
    CancelButtonColor = m_CancelButtonColor
End Property

Public Property Let CancelButtonColor(Value As Long)
    m_CancelButtonColor = Value
    btnCancel.BackColor = m_CancelButtonColor
End Property

Public Property Get AddButtonVisible() As Byte
    AddButtonVisible = m_AddButtonVisible
End Property

Public Property Let AddButtonVisible(Value As Byte)
    m_AddButtonVisible = Value
End Property

Public Property Get CancelButtonVisible() As Byte
    CancelButtonVisible = m_CancelButtonVisible
End Property

Public Property Let CancelButtonVisible(Value As Byte)
    m_CancelButtonVisible = Value
End Property

Public Property Get AddAction() As String
    AddAction = m_AddAction
End Property

Public Property Let AddAction(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "add"
    If ValidateString(Value, "alphanumdashunder") Then
        m_AddAction = Value
    End If
End Property

Public Property Get CancelAction() As String
    CancelAction = m_CancelAction
End Property

Public Property Let CancelAction(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "cancel"
    If ValidateString(Value, "alpha") Then
        m_CancelAction = Value
    End If
End Property
Public Property Get EditAction() As String
    EditAction = m_EditAction
End Property

Public Property Let EditAction(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "edit"
    If ValidateString(Value, "alpha") Then
        m_EditAction = Value
    End If
End Property

Public Property Let CallingForm(Value As String)
    If Len(Value) > 0 Then
        m_CallingForm = Value
    Else
        RaiseEvent InvalidCallingForm(Value)
    End If
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'---------------------
' Events
'---------------------
' ---------------------------------
' Sub:          Form_Open
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 4, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/18/2017 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    Dim ary() As String
    
    'default
    Me.CallingForm = "Main"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize calling form
    ToggleForm Me.CallingForm, -1
    
    Me.FormHeader.BackColor = lngBrown
    Me.TitleFontColor = lngWhite
    Me.Title = "Comment"
    
    Me.lineIndicator.Width = Me.Form.Width
    Me.lineIndicator.borderColor = lngLime
    
    'defaults
    Dim instruction As String
    Dim MaxCount As Integer
    
    instruction = "Enter your establishment comment."
    MaxCount = 50
    
    'set comment context
    ary = Split(Nz(Me.OpenArgs, ""), "|")
    If IsArray(ary) Then
        Me.Context = ary(0) & " - " & ary(1) '"Plot - 24"
        
        'set filter for subform
        Me.list.Form.Filter = "CommentType='" & ary(0) & "' AND CommentType_ID=" & ary(1)
        Me.list.Form.FilterOn = True
        
        'update subform
'        Me.list.Form.Requery
        
        MaxCount = ary(2)
        
        'set instructions based on calling form
        Select Case LCase(ary(0))
            Case "importeddata"
                instruction = "Enter your import comment."
        End Select
    Else
        GoTo Exit_Handler
    End If
    
    Me.Instructions = instruction
    Me.CountLabelVisible = False
    Me.CurrentCount = "Characters Remaining:"
    Me.lblCharacterCount.Visible = False
    Me.MaxCount = MaxCount
    Me.AlertCount = 10
   
    Me.AddAction = "add_"
    
    Me.Context = Me.OpenArgs

    'default
    btnAdd.Enabled = False
    lblMsgIcon.Caption = ""
    lblMsg.Caption = ""

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Comment form])"
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
' Source/date:  Bonnie Campbell, November 4, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/4/2015 - initial version
'   BLC - 12/5/2016 - added instruction and max count inputs
'   BLC - 10/18/2017 - shift to Form_Open()
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler
    
'    Dim ary() As String
'
'    'default
'    Me.CallingForm = "Main"
'
'    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs
'
'    'minimize calling form
'    ToggleForm Me.CallingForm, -1
'
'    Me.FormHeader.BackColor = lngBrown
'    Me.TitleFontColor = lngWhite
'
'    Me.lineIndicator.Width = Me.Form.Width
'    Me.lineIndicator.borderColor = lngLime
'
'    'defaults
'    Dim instruction As String
'    Dim MaxCount As Integer
'
'    instruction = "Enter your establishment comment."
'    MaxCount = 50
'
'    'set comment context
'    ary = Split(Nz(Me.OpenArgs, ""), "|")
'    If IsArray(ary) Then
'        Me.Context = ary(0) & " - " & ary(1) '"Plot - 24"
'
'        'set filter for subform
'        Me.list.Form.Filter = "CommentType=" & ary(0) & " AND CommentType_ID=" & ary(1)
'        'update subform
'        Me.list.Form.Requery
'
'        MaxCount = ary(2)
'
'        'set instructions based on calling form
'        Select Case LCase(ary(0))
'            Case "importeddata"
'                instruction = "Enter your import comment."
'        End Select
'    Else
'        GoTo Exit_Handler
'    End If
'
'    Me.Instructions = instruction
'    Me.CountLabelVisible = False
'    Me.CurrentCount = "Characters Remaining:"
'    Me.lblCharacterCount.Visible = False
'    Me.MaxCount = MaxCount
'    Me.AlertCount = 10
'
'    Me.AddAction = "add_"
'
'    Me.Context = Me.OpenArgs

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Comment form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_KeyPress
' Description:  Form keypress actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 19, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/19/2017 - initial version
' ---------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler

Debug.Print KeyAscii
    If KeyAscii = iTabKey Then
        lblMsgIcon.Caption = ""
        lblMsg.Caption = ""
        Me.Requery
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_KeyPress[Site form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxComment_Change
' Description:  tbxComment actions on change event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 4, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/4/2015 - initial version
'   BLC - 10/19/2017 - added comment message
' ---------------------------------
Private Sub tbxComment_Change()
On Error GoTo Err_Handler
    
    Dim CurrentCount As Integer
    
    CurrentCount = CInt(Me.MaxCount) - Len(tbxComment.Text)

    Me.lblMaxCount.Caption = CurrentCount & " remaining"
    
    Me.CurrentCountFontColor = vbBlack
    Me.AlertBoxVisible = False
    Me.MaxCountFontColor = vbBlack
    
    Select Case CurrentCount
        Case Is < Me.AlertCount
            Me.AlertBoxVisible = True
            Me.AlertBoxBackgroundColor = lngYellow
        Case Is = 0
            Me.CurrentCountFontColor = vbRed
        Case Else
    End Select
    
    If CurrentCount < 1 Then 'CInt(Me.MaxCount) Then
        Me.MaxCountFontColor = vbRed
    End If
    
'    If Len(tbxComment.Text) > CInt(Me.MaxCount) Then
'        Me.lblMaxCount.Caption = -CurrentCount & " over"
'        'disable add comment button until count is < or = MaxCount
'        Me.btnAdd.Enabled = False
'    ElseIf Len(tbxComment.Text) = 0 Then
'        'disable add comment button if count = 0
'        Me.btnAdd.Enabled = False
'    Else
'        're-enable add comment button
'        Me.btnAdd.Enabled = True
'    End If

    'enable add/save for new comments
    btnAdd.Enabled = False
    
    If Len(tbxComment.Text) > 0 And tbxID = 0 Then
        btnAdd.Enabled = True
    ElseIf Len(tbxComment.Text) < MaxCount + 1 And tbxID > 0 Then
        lblMsg.ForeColor = lngYellow
        lblMsgIcon.ForeColor = lngYellow
        lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
        lblMsg.Caption = "Tab to enter comment changes..."
    Else
        lblMsgIcon.Caption = ""
        lblMsg.Caption = ""
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxComment_Change[Comment form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxComment_KeyPress
' Description:  textbox keypress actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 19, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/19/2017 - initial version
' ---------------------------------
Private Sub tbxComment_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler

Debug.Print KeyAscii
    If KeyAscii = iTabKey Then
        lblMsgIcon.Caption = ""
        lblMsg.Caption = ""
        Me.Repaint
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxComment_KeyPress[Site tbxComment])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAdd_Click
' Description:  Add comment form entry
' Assumptions:  Person using the application is the "commentor"
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 12, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/12/2015 - initial version
'   BLC - 8/9/2016   - revised Comment > AppComment (comment reserved word)
'   BLC - 12/6/2016 - revise so comment type = context before "- ID#"
'   BLC - 9/25/2017 - revise for NCPN_framework.XX classes
'   BLC - 9/27/2017 - update to use Factory.NewClassXX() vs GetClass()
'   BLC - 10/17/2017 - added title to comment
'   BLC - 10/19/2017 - added comment message
' ---------------------------------
Private Sub btnAdd_Click()
On Error GoTo Err_Handler
    
    'Dim oComment As New AppComment
    Dim oComment As NCPN_framework.AppComment
    Set oComment = Factory.NewAppComment
    
    With oComment
        .CommentType = Left(lblContext.Caption, InStr(lblContext.Caption, " - "))
        .TypeID = RemoveChars(lblContext.Caption, True) 'return only numbers
        .Comment = tbxComment.Value
        .CommentorID = TempVars("AppUserID") '3 'Requestor
        '.RequestedByID = 3 'Requestor
        .AddComment
    
        If IsNumeric(.ID) Then
'            MsgBox "New Comment ID = " & .ID
            lblMsg.ForeColor = lngYellow
            lblMsgIcon.ForeColor = lngYellow
            lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
            lblMsg.Caption = "Comment added!"
    
            list.Requery

            'show added record message & clear
'            DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
'                        "msg" & PARAM_SEPARATOR & "Comment added (# " & .ID & " )" & _
'                        "|Type" & PARAM_SEPARATOR & "info" & _
'                        "|Title" & PARAM_SEPARATOR & "Comment Added!"

            
            'close comment form
'            DoCmd.Close acForm, "Comment"
            
        End If
    
    End With

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAdd_Click[Comment form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnCancel_Click
' Description:  Cancel comment form entry
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 4, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/4/2015 - initial version
' ---------------------------------
Private Sub btnCancel_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCancel_Click[Comment form])"
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
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
'   BLC - 10/20/2016 - revised to use callingform property
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore calling form
    ToggleForm Me.CallingForm, 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[Comment form])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          SetHeaderColor
' Description:  Set header color event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 28, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/3/2015 - initial version
' ---------------------------------
Private Sub SetHeaderColor(color As Long)
On Error GoTo Err_Handler
    
    MsgBox "SetHeaderColor...", vbOKOnly
    Me.CommentHeaderColor = color

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetHeaderColor[Comment form])"
    End Select
    Resume Exit_Handler
End Sub
