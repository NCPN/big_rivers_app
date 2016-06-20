Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5820
    DatasheetFontHeight =11
    ItemSuffix =16
    Left =4230
    Top =3960
    Right =10065
    Bottom =4320
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x98f234fbd5a9e440
    End
    Caption ="Data Entry"
    OnCurrent ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnMouseDown ="[Event Procedure]"
    OnActivate ="[Event Procedure]"
    OnGotFocus ="[Event Procedure]"
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
            Height =0
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin Section
            Height =480
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =1380
                    ForeColor =4210752
                    Name ="btnLevel0"
                    Caption ="BLCA > "
                    OnClick ="[Event Procedure]"
                    OnMouseDown ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =420
                    Gradient =0
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderColor =14136213
                    HoverColor =9699294
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =16777164
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =9974127
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =60
                    Width =1500
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnLevel1"
                    Caption ="Gunnison > "
                    OnClick ="[Event Procedure]"
                    OnMouseDown ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedTop =60
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =420
                    Gradient =0
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderColor =14136213
                    HoverColor =9699294
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =16777164
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =9974127
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2940
                    Top =60
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnLevel2"
                    Caption ="Site > "
                    OnClick ="[Event Procedure]"
                    OnMouseDown ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2940
                    LayoutCachedTop =60
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =420
                    Gradient =0
                    BackColor =65535
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderColor =14136213
                    HoverColor =9699294
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =16777164
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =9974127
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4380
                    Top =60
                    Width =1380
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnLevel3"
                    Caption ="Feature"
                    OnClick ="[Event Procedure]"
                    OnMouseDown ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =60
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =420
                    Gradient =0
                    BackColor =65535
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderColor =14136213
                    HoverColor =9699294
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =16777164
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =9974127
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
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
' Form:         _Level
' Level:        Application form
' Version:      1.00
' Basis:        Level form
'
' Description:  Level form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, April 20, 2016
' References:   -
' Revisions:    BLC - 4/20/2016 - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Level0Color As Long
Private m_Level1Color As Long
Private m_Level2Color As Long
Private m_Level3Color As Long
Private m_Level0BgdColor As Long
Private m_Level1BgdColor As Long
Private m_Level2BgdColor As Long
Private m_Level3BgdColor As Long
Private m_Level0HoverColor As Long
Private m_Level1HoverColor As Long
Private m_Level2HoverColor As Long
Private m_Level3HoverColor As Long
Private m_Level0HoverBgdColor As Long
Private m_Level1HoverBgdColor As Long
Private m_Level2HoverBgdColor As Long
Private m_Level3HoverBgdColor As Long

'---------------------
' Event Declarations
'---------------------
Public Event InvalidLevel(value As String)
Public Event InvalidColor(value As Long)

'---------------------
' Properties
'---------------------

'-- std color/bgd color --
Public Property Let Level0Color(value As Long)
    If Len(value) > 0 Then
        m_Level0Color = value
        Me.btnLevel0.backcolor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level0Color() As Long
    Level0Color = m_Level0Color
End Property

Public Property Let Level1Color(value As Long)
    If Len(value) > 1 Then
        m_Level1Color = value
        Me.btnLevel1.backcolor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level1Color() As Long
    Level1Color = m_Level1Color
End Property

Public Property Let Level2Color(value As Long)
    If Len(value) > 2 Then
        m_Level2Color = value
        Me.btnLevel2.backcolor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level2Color() As Long
    Level2Color = m_Level2Color
End Property

Public Property Let Level3Color(value As Long)
    If Len(value) > 3 Then
        m_Level3Color = value
        Me.btnLevel3.backcolor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level3Color() As Long
    Level3Color = m_Level3Color
End Property

Public Property Let Level0BgdColor(value As Long)
    If Len(value) > 0 Then
        m_Level0BgdColor = value
        Me.btnLevel0.backcolor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level0BgdColor() As Long
    Level0BgdColor = m_Level0BgdColor
End Property

Public Property Let Level1BgdColor(value As Long)
    If Len(value) > 1 Then
        m_Level1BgdColor = value
        Me.btnLevel1.backcolor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level1BgdColor() As Long
    Level1BgdColor = m_Level1BgdColor
End Property

Public Property Let Level2BgdColor(value As Long)
    If Len(value) > 2 Then
        m_Level2BgdColor = value
        Me.btnLevel2.backcolor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level2BgdColor() As Long
    Level2BgdColor = m_Level2BgdColor
End Property

Public Property Let Level3BgdColor(value As Long)
    If Len(value) > 3 Then
        m_Level3BgdColor = value
        Me.btnLevel3.backcolor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level3BgdColor() As Long
    Level3BgdColor = m_Level3BgdColor
End Property

'-- on hover --
Public Property Let Level0HoverColor(value As Long)
    If Len(value) > 0 Then
        m_Level0HoverColor = value
        Me.btnLevel0.hoverForeColor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level0HoverColor() As Long
    Level0HoverColor = m_Level0HoverColor
End Property

Public Property Let Level1HoverColor(value As Long)
    If Len(value) > 1 Then
        m_Level1HoverColor = value
        Me.btnLevel1.hoverForeColor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level1HoverColor() As Long
    Level1HoverColor = m_Level1HoverColor
End Property

Public Property Let Level2HoverColor(value As Long)
    If Len(value) > 2 Then
        m_Level2HoverColor = value
        Me.btnLevel2.hoverForeColor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level2HoverColor() As Long
    Level2HoverColor = m_Level2HoverColor
End Property

Public Property Let Level3HoverColor(value As Long)
    If Len(value) > 3 Then
        m_Level3HoverColor = value
        Me.btnLevel3.hoverForeColor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level3HoverColor() As Long
    Level3HoverColor = m_Level3HoverColor
End Property

Public Property Let Level0HoverBgdColor(value As Long)
    If Len(value) > 0 Then
        m_Level0HoverBgdColor = value
        Me.btnLevel0.hoverColor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level0HoverBgdColor() As Long
    Level0HoverBgdColor = m_Level0HoverBgdColor
End Property

Public Property Let Level1HoverBgdColor(value As Long)
    If Len(value) > 1 Then
        m_Level1HoverBgdColor = value
        Me.btnLevel1.hoverColor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level1HoverBgdColor() As Long
    Level1HoverBgdColor = m_Level1HoverBgdColor
End Property

Public Property Let Level2HoverBgdColor(value As Long)
    If Len(value) > 2 Then
        m_Level2HoverBgdColor = value
        Me.btnLevel2.hoverColor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level2HoverBgdColor() As Long
    Level2HoverBgdColor = m_Level2HoverBgdColor
End Property

Public Property Let Level3HoverBgdColor(value As Long)
    If Len(value) > 3 Then
        m_Level3HoverBgdColor = value
        Me.btnLevel3.hoverColor = value
    Else
        RaiseEvent InvalidColor(value)
    End If
End Property

Public Property Get Level3HoverBgdColor() As Long
    Level3HoverBgdColor = m_Level3HoverBgdColor
End Property

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
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler
    
    'add right click menu
    CreateMenu "level"
    CreateDynamicMenu "park"
    CreateDynamicMenu "river"
    CreateDynamicMenu "site"
    CreateDynamicMenu "feature"
    
    Me.btnLevel0.Caption = Nz(TempVars("ParkCode"), "Park") & " > "
    Me.btnLevel1.Caption = Nz(TempVars("River"), "River") & " > "
    Me.btnLevel2.Caption = Nz(TempVars("SiteCode"), "Site") & " > "
    Me.btnLevel3.Caption = Nz(TempVars("Feature"), "Feature")
    
    'hide level3
'    Me.btnLevel3.Visible = False
'    Me.btnLevel3.Caption = " "

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[_Level form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_GotFocus
' Description:  Actions after returning to form
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub Form_GotFocus()
On Error GoTo Err_Handler
        
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_GotFocus[_Level form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Activate
' Description:  Actions after returning to form
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub Form_Activate()
On Error GoTo Err_Handler
        
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Activate[_Level form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  Actions after returning to form
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
        
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[_Level form])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------------------------
' Sub:          btnLevel0_Click
' Description:  btnLevel0 click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub btnLevel0_Click()
On Error GoTo Err_Handler

    SysCmd acSysCmdSetStatus, "lvl0"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLevel0_Click[_Level form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnLevel1_Click
' Description:  btnLevel1 click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub btnLevel1_Click()
On Error GoTo Err_Handler

    SysCmd acSysCmdSetStatus, "lvl1"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLevel1_Click[_Level form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnLevel2_Click
' Description:  btnLevel2 click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub btnLevel2_Click()
On Error GoTo Err_Handler

    SysCmd acSysCmdSetStatus, "lvl2"
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLevel2_Click[_Level form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnLevel3_Click
' Description:  btnLevel3 click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub btnLevel3_Click()
On Error GoTo Err_Handler

    SysCmd acSysCmdSetStatus, "lvl3"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLevel3_Click[_Level form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnLevel0_MouseDown
' Description:  btnLevel0 mouse actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub btnLevel0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    If Button = acRightButton Then
        CommandBars("park").ShowPopup
        DoCmd.CancelEvent
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLevel0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)[_Level form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnLevel1_MouseDown
' Description:  btnLevel1 mouse actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub btnLevel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    If Button = acRightButton Then
        CommandBars("river").ShowPopup
        DoCmd.CancelEvent
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLevel1_MouseDown[_Level form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnLevel2_MouseDown
' Description:  btnLevel2 mouse actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub btnLevel2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    If Button = acRightButton Then
        CommandBars("site").ShowPopup
        DoCmd.CancelEvent
    End If
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLevel2_MouseDown[_Level form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnLevel3_MouseDown
' Description:  btnLevel3 mouse actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub btnLevel3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    If Button = acRightButton Then
        CommandBars("feature").ShowPopup
        DoCmd.CancelEvent
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLevel3_MouseDown[_Level form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_MouseDown
' Description:  form mouse down keystroke actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Seth Schrock, May 23, 2013
'   https://bytes.com/topic/access/answers/949589-how-do-i-create-custom-right-click-menu
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    If Button = acRightButton Then
        CommandBars("setlevel").ShowPopup
        DoCmd.CancelEvent
    End If
    
Exit_Handler:
        'clear status bar
        SysCmd acSysCmdClearStatus
        
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_MouseDown[_Level form])"
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
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    
Exit_Handler:
        'clear status bar
        SysCmd acSysCmdClearStatus
        
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[_Level form])"
    End Select
    Resume Exit_Handler
End Sub
