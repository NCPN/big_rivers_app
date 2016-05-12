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
    Width =2220
    DatasheetFontHeight =11
    ItemSuffix =10
    Right =20208
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
            Height =300
            BackColor =65280
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =1980
                    Height =300
                    Name ="lblCrumb"
                    Caption ="Left"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =300
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ForeShade =95.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =2040
                    Width =180
                    Height =300
                    Name ="lblSeparator"
                    Caption =">"
                    GridlineColor =10921638
                    LayoutCachedLeft =2040
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =300
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ForeShade =95.0
                End
            End
        End
        Begin Section
            Height =0
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' Form:         Breadcrumb
' Level:        Framework form
' Version:      1.00
'
' Description:  Breadcrumb form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 10/30/2015
' References:
' Revisions:    BLC - 10/30/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_Crumb As String
Private m_BreadcrumbHeaderColor As Long
Private m_CrumbFontColor As Long
Private m_BreadcrumbVisible As Byte
Private m_CrumbWidth As Integer
Private m_CrumbHeight As Integer

'Dim crumb As New link

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
Public Property Let crumb(Value As String)
    m_Crumb = Value
    lblCrumb.Caption = m_Crumb
End Property

Public Property Get crumb() As String
    crumb = m_Crumb
End Property

Public Property Let CrumbFontColor(Value As Long)
    m_CrumbFontColor = Value
    lblCrumb.ForeColor = m_CrumbFontColor
End Property

Public Property Get CrumbFontColor() As Long
    CrumbFontColor = m_CrumbFontColor
End Property

Public Property Let BreadcrumbHeaderColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen '"#3F3F3F"
    m_BreadcrumbHeaderColor = Value
    FormHeader.BackColor = m_BreadcrumbHeaderColor
    'set font color to match
    Select Case Value
        Case vbGreen, vbCyan, vbWhite
            Me.CrumbFontColor = vbBlack
        Case vbRed, vbBlue, vbMagenta, vbBlack
            Me.CrumbFontColor = vbWhite
    End Select
End Property

Public Property Get BreadcrumbHeaderColor() As Long
    BreadcrumbHeaderColor = m_BreadcrumbHeaderColor 'FormHeader.BackColor
End Property

Public Property Get BreadcrumbVisible() As Byte
    BreadcrumbVisible = m_BreadcrumbVisible
End Property

Public Property Let BreadcrumbVisible(Value As Byte)
    m_BreadcrumbVisible = Value
    Me.Visible = m_BreadcrumbVisible
End Property

'---------------------
' Events
'---------------------

' ---------------------------------
' Sub:          lblCrumb_Click
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
Private Sub lblCrumb_Click()
On Error GoTo Err_Handler

    MsgBox "You have not selected the number of images. Please do not delay!", vbOKOnly

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblCrumb_Click[Breadcrumb form])"
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
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
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
            "Error encountered (#" & Err.Number & " - Class_Initialize[Breadcrumb form])"
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
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler
Exit_Handler:
    Exit Sub
    
    MsgBox "Terminating...", vbOKOnly
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[Breadcrumb form])"
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
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub SetHeaderColor(Color As Long)
On Error GoTo Err_Handler
    
    MsgBox "SetHeaderColor...", vbOKOnly
    Me.BreadcrumbHeaderColor = Color

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[Breadcrumb form])"
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
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------


Private Sub Form_Load()
On Error GoTo Err_Handler
    
    MsgBox "SetHeaderColor...", vbOKOnly
    Me.BreadcrumbHeaderColor = vbCyan
    
    'Set crumb = lblCrumb
    'crumb.Action = "DoCmd.OpenForm('Tile', acNormal)"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Breadcrumb form])"
    End Select
    Resume Exit_Handler
End Sub
