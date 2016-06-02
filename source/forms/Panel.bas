Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2592
    DatasheetFontHeight =11
    ItemSuffix =9
    Right =15552
    Bottom =9408
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x06dd372434a7e440
    End
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
        Begin FormHeader
            Height =444
            BackColor =65280
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
                    Name ="lblTitle"
                    Caption ="Left"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =360
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ForeShade =95.0
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =85
                    Top =432
                    Width =2592
                    BorderColor =65280
                    Name ="lineIndicator"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedTop =432
                    LayoutCachedWidth =2592
                    LayoutCachedHeight =432
                    BorderThemeColorIndex =-1
                End
            End
        End
        Begin Section
            Height =2280
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
                    Width =2160
                    Height =2040
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblMessage"
                    Caption ="Message"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =2100
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
' Form:         Panel
' Level:        Framework form
' Version:      1.00
'
' Description:  Panel form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 10/30/2015
' References:
' Revisions:    BLC - 10/30/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_MessageCaption As String
Private m_BarColor As Variant
Private m_PanelHeaderColor As Long
Private m_TitleFontColor As Long
Private m_MessageFontColor As Long
Private m_PanelVisible As Byte
Private m_MessageVisible As Byte

'---------------------
' Events
'---------------------
Public Event Selected()
Public Event CriticalState()
Public Event GoodState()
Public Event Initialize()
Public Event Terminate()

'---------------------
' Properties
'---------------------
Public Property Let Title(Value As String)
    m_Title = Value
    lblTitle.Caption = m_Title
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Get MessageCaption() As String
    MessageCaption = m_MessageCaption
End Property

Public Property Let MessageCaption(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Message"
    m_MessageCaption = Value
    lblMessage.Caption = m_MessageCaption
End Property

Public Property Let TitleFontColor(Value As Long)
    m_TitleFontColor = Value
    lblTitle.ForeColor = m_TitleFontColor
End Property

Public Property Get TitleFontColor() As Long
    TitleFontColor = m_TitleFontColor
End Property

Public Property Let PanelHeaderColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen '"#3F3F3F"
    m_PanelHeaderColor = Value
    FormHeader.backColor = m_PanelHeaderColor
    'set font color to match
    Select Case Value
        Case vbGreen
            Me.TitleFontColor = vbBlack
        Case vbRed, vbBlue
            Me.TitleFontColor = vbWhite
    End Select
End Property

Public Property Get PanelHeaderColor() As Long
    PanelHeaderColor = m_PanelHeaderColor 'FormHeader.BackColor
End Property

Public Property Let BarColor(Value As Variant)
    m_BarColor = Value
    Me.lineIndicator.BorderColor = m_BarColor
End Property

Public Property Get BarColor()
    BarColor = m_BarColor
End Property

Public Property Get MessageFontColor() As Long
    MessageFontColor = m_MessageFontColor
End Property

Public Property Let MessageFontColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen
    m_MessageFontColor = Value
End Property

Public Property Get PanelVisible() As Byte
    PanelVisible = m_PanelVisible
End Property

Public Property Let PanelVisible(Value As Byte)
    m_PanelVisible = Value
    Me.Visible = m_PanelVisible
End Property

Public Property Get MessageVisible() As Byte
    MessageVisible = m_MessageVisible
End Property

Public Property Let MessageVisible(Value As Byte)
    m_MessageVisible = Value
    Me.lblMessage.Visible = m_MessageVisible
End Property

'---------------------
' Events
'---------------------

' ---------------------------------
' Sub:          lblMessage_Click
' Description:  Link click event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 29, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/29/2015 - initial version
' ---------------------------------
Private Sub lblMessage_Click()
On Error GoTo Err_Handler

    MsgBox "You have not selected the number of images. Please do not delay!", vbOKOnly

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblMessage_Click[Panel form])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Class_Initialize
' Description:  Class initialization (starting) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 28, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/28/2015 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

    MsgBox "Initializing...", vbOKOnly

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[Panel form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Class_Terminate
' Description:  Class termination (closing) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 28, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/28/2015 - initial version
' ---------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler
    
    MsgBox "Terminating...", vbOKOnly

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[Panel form])"
    End Select
    Resume Exit_Handler
End Sub

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
'   BLC - 10/28/2015 - initial version
' ---------------------------------
Private Sub SetHeaderColor(color As Long)
On Error GoTo Err_Handler
    
    MsgBox "SetHeaderColor...", vbOKOnly
    Me.PanelHeaderColor = color

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[Panel form])"
    End Select
    Resume Exit_Handler
End Sub
