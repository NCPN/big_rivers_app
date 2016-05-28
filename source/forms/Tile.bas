Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
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
    Left =3150
    Top =3105
    Right =23715
    Bottom =12735
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
            Height =447
            BackColor =16711680
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
                    Caption ="Reports"
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
                    BorderColor =65535
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
            OnMouseMove ="[Event Procedure]"
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
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLink1"
                    Caption ="rpt1"
                    OnClick ="[Event Procedure]"
                    OnMouseMove ="[Event Procedure]"
                    Tag ="DISABLED"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =-1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =900
                    Top =1920
                    Width =1200
                    Height =240
                    ForeColor =4210752
                    Name ="btnClick"
                    Caption ="Next >>"
                    GridlineColor =10921638

                    LayoutCachedLeft =900
                    LayoutCachedTop =1920
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =2160
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
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =180
                    Top =420
                    Width =2160
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLink2"
                    Caption ="Link2"
                    OnClick ="[Event Procedure]"
                    OnMouseMove ="[Event Procedure]"
                    Tag ="DISABLED"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =660
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =180
                    Top =780
                    Width =2160
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLink3"
                    Caption ="Link3"
                    OnClick ="[Event Procedure]"
                    OnMouseMove ="[Event Procedure]"
                    Tag ="DISABLED"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =780
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =1020
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =180
                    Top =1140
                    Width =2160
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLink4"
                    Caption ="Link4"
                    OnClick ="[Event Procedure]"
                    OnMouseMove ="[Event Procedure]"
                    Tag ="DISABLED"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =1140
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =1380
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =180
                    Top =1500
                    Width =2160
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLink5"
                    Caption ="Link5"
                    OnClick ="[Event Procedure]"
                    OnMouseMove ="[Event Procedure]"
                    Tag ="DISABLED"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =1500
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =1740
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =180
                    Top =1860
                    Width =2160
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLink6"
                    Caption ="Link6"
                    OnClick ="[Event Procedure]"
                    OnMouseMove ="[Event Procedure]"
                    Tag ="DISABLED"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =1860
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
' Form:         Tile
' Level:        Framework form
' Version:      1.00
'
' Description:  Tile form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:
'  SourceDaddy, unknown
'  http://sourcedaddy.com/ms-access/managing-class-interface.htm
'  Denise Gosnell, 2011
'  Beginning Access 2007 VBA
'  https://books.google.com/books?id=z2aoFGg1HFAC&pg=SA3-PA30&dq=access+vba+creating+custom+form+controls&hl=en&sa=X&ved=0CDAQ6AEwAGoVChMI6KblxdHoyAIVBcdjCh3Okw9V#v=onepage&q=access%20vba%20creating%20custom%20form%20controls&f=false
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 4/26/2016  - 1.01 - added tile tag property
' =================================

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_TileTag As String
Private m_Link1Caption As String
Private m_Link2Caption As String
Private m_Link3Caption As String
Private m_Link4Caption As String
Private m_Link5Caption As String
Private m_Link6Caption As String
Private m_BarColor As Variant
Private m_TileHeaderColor As Long
Private m_TitleFontColor As Long
Private m_Link1FontColor As Long
Private m_Link2FontColor As Long
Private m_Link3FontColor As Long
Private m_Link4FontColor As Long
Private m_Link5FontColor As Long
Private m_Link6FontColor As Long
Private m_TileVisible As Byte
Private m_Link1Visible As Byte
Private m_Link2Visible As Byte
Private m_Link3Visible As Byte
Private m_Link4Visible As Byte
Private m_Link5Visible As Byte
Private m_Link6Visible As Byte
Private m_Link1Action As String
Private m_Link2Action As String
Private m_Link3Action As String
Private m_Link4Action As String
Private m_Link5Action As String
Private m_Link6Action As String

'---------------------
' Events
'---------------------
Public Event Selected(Value As Boolean)
Public Event CriticalState(Value As Boolean)
Public Event GoodState(Value As Boolean)

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

Public Property Let TileTag(Value As String)
    m_TileTag = Value
    lblLink1.Tag = m_TileTag
    lblLink2.Tag = m_TileTag
    lblLink3.Tag = m_TileTag
    lblLink4.Tag = m_TileTag
    lblLink5.Tag = m_TileTag
    lblLink6.Tag = m_TileTag
End Property

Public Property Get TileTag() As String
    TileTag = m_TileTag
End Property

Public Property Get Link1Caption() As String
    Link1Caption = m_Link1Caption
End Property

Public Property Let Link1Caption(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link1"
    m_Link1Caption = Value
    lblLink1.Caption = m_Link1Caption
End Property

Public Property Get Link2Caption() As String
    Link2Caption = m_Link2Caption
End Property

Public Property Let Link2Caption(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link2"
    m_Link2Caption = Value
    lblLink2.Caption = m_Link2Caption
End Property

Public Property Get Link3Caption() As String
    Link3Caption = m_Link3Caption
End Property

Public Property Let Link3Caption(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link3"
    m_Link3Caption = Value
    lblLink3.Caption = m_Link3Caption
End Property

Public Property Get Link4Caption() As String
    Link4Caption = m_Link4Caption
End Property

Public Property Let Link4Caption(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link4"
    m_Link4Caption = Value
    lblLink4.Caption = m_Link4Caption
End Property

Public Property Get Link5Caption() As String
    Link5Caption = m_Link5Caption
End Property

Public Property Let Link5Caption(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link5"
    m_Link5Caption = Value
    lblLink5.Caption = m_Link5Caption
End Property

Public Property Get Link6Caption() As String
    Link6Caption = m_Link6Caption
End Property

Public Property Let Link6Caption(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link6"
    m_Link6Caption = Value
    lblLink6.Caption = m_Link6Caption
End Property

Public Property Get Link1Action() As String
    Link1Action = m_Link1Action
End Property

Public Property Let Link1Action(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link1"
    m_Link1Action = Value
End Property

Public Property Get Link2Action() As String
    Link2Action = m_Link2Action
End Property

Public Property Let Link2Action(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link2"
    m_Link2Action = Value
End Property

Public Property Get Link3Action() As String
    Link3Action = m_Link3Action
End Property

Public Property Let Link3Action(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link3"
    m_Link3Action = Value
End Property

Public Property Get Link4Action() As String
    Link4Action = m_Link4Action
End Property

Public Property Let Link4Action(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link4"
    m_Link4Action = Value
End Property

Public Property Get Link5Action() As String
    Link5Action = m_Link5Action
End Property

Public Property Let Link5Action(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link5"
    m_Link5Action = Value
End Property

Public Property Get Link6Action() As String
    Link6Action = m_Link6Action
End Property

Public Property Let Link6Action(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Link6"
    m_Link6Action = Value
End Property

Public Property Let TitleFontColor(Value As Long)
    m_TitleFontColor = Value
    lblTitle.ForeColor = m_TitleFontColor
End Property

Public Property Get TitleFontColor() As Long
    TitleFontColor = m_TitleFontColor
End Property

Public Property Let TileHeaderColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen '"#3F3F3F"
    m_TileHeaderColor = Value
    FormHeader.backColor = m_TileHeaderColor
    'set font color to match
    Select Case Value
        Case vbGreen
            Me.TitleFontColor = vbBlack
        Case vbRed, vbBlue
            Me.TitleFontColor = vbWhite
    End Select
End Property

Public Property Get TileHeaderColor() As Long
    TileHeaderColor = m_TileHeaderColor 'FormHeader.BackColor
End Property

Public Property Let BarColor(Value As Variant)
    m_BarColor = Value
    Me.lineIndicator.BorderColor = m_BarColor
End Property

Public Property Get BarColor()
    BarColor = m_BarColor
End Property

Public Property Get Link1FontColor() As Long
    Link1FontColor = m_Link1FontColor
End Property

Public Property Let Link1FontColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen
    m_Link1FontColor = Value
End Property

Public Property Get Link2FontColor() As Long
    Link2FontColor = m_Link2FontColor
End Property

Public Property Let Link2FontColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen
    m_Link2FontColor = Value
End Property

Public Property Get Link3FontColor() As Long
    Link3FontColor = m_Link3FontColor
End Property

Public Property Let Link3FontColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen
    m_Link3FontColor = Value
End Property

Public Property Get Link4FontColor() As Long
    Link4FontColor = m_Link4FontColor
End Property

Public Property Let Link4FontColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen '"#3F3F3F"
    m_Link4FontColor = Value
End Property

Public Property Get Link5FontColor() As Long
    Link5FontColor = m_Link5FontColor
End Property

Public Property Let Link5FontColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen '"#3F3F3F"
    m_Link5FontColor = Value
End Property

Public Property Get Link6FontColor() As Long
    Link6FontColor = m_Link6FontColor
End Property

Public Property Let Link6FontColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen '"#3F3F3F"
    m_Link6FontColor = Value
End Property

Public Property Get TileVisible() As Byte
    TileVisible = m_TileVisible
End Property

Public Property Let TileVisible(Value As Byte)
    m_TileVisible = Value
    Me.Visible = m_TileVisible
End Property

Public Property Get Link1Visible() As Byte
    Link1Visible = m_Link1Visible
End Property

Public Property Let Link1Visible(Value As Byte)
    m_Link1Visible = Value
    Me.lblLink1.Visible = m_Link1Visible
End Property

Public Property Get Link2Visible() As Byte
    Link2Visible = m_Link2Visible
End Property

Public Property Let Link2Visible(Value As Byte)
    m_Link2Visible = Value
    Me.lblLink2.Visible = m_Link2Visible
End Property

Public Property Get Link3Visible() As Byte
    Link3Visible = m_Link3Visible
End Property

Public Property Let Link3Visible(Value As Byte)
    m_Link3Visible = Value
    Me.lblLink3.Visible = m_Link3Visible
End Property

Public Property Get Link4Visible() As Byte
    Link4Visible = m_Link4Visible
End Property

Public Property Let Link4Visible(Value As Byte)
    m_Link4Visible = Value
    Me.lblLink4.Visible = m_Link4Visible
End Property

Public Property Get Link5Visible() As Byte
    Link5Visible = m_Link5Visible
End Property

Public Property Let Link5Visible(Value As Byte)
    m_Link5Visible = Value
    Me.lblLink5.Visible = m_Link5Visible
End Property

Public Property Get Link6Visible() As Byte
    Link6Visible = m_Link6Visible
End Property

Public Property Let Link6Visible(Value As Byte)
    m_Link6Visible = Value
    Me.lblLink6.Visible = m_Link6Visible
End Property

'---------------------
' Events
'---------------------

' ---------------------------------
' Sub:          Form_Load
' Description:  Form loading event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2016 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    Dim ctrl As Control
    Dim strLink As String
    Dim i As Integer
    
'    MsgBox "loading...", vbOKOnly
    
    'initialize labels (fixes if form inadvertently saved in non-standard state)
    lblTitle.Caption = "Title"
        
    For i = 1 To 6
        
        strLink = "lblLink" & i
    
        For Each ctrl In Me.Controls
        
            If ctrl.Name = strLink Then
                ctrl.Caption = Replace(strLink, "lbl", "")
                ctrl.Visible = True
                
                'disable clicks until value set
                ctrl.Tag = "DISABLED"
                
                Exit For
            End If
            
        Next
    Next
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLink1_Click
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
Private Sub lblLink1_Click()
On Error GoTo Err_Handler
    
    If lblLink1.Tag = "DISABLED" Then GoTo Exit_Handler
    
    With Me.lblLink1
        DoCmd.Minimize
        ClickAction .Tag & .Caption
    End With
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink1_Click[Tile form])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------------------------
' Sub:          lblLink2_Click
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
Private Sub lblLink2_Click()
On Error GoTo Err_Handler
    
    If lblLink2.Tag = "DISABLED" Then GoTo Exit_Handler
    
    With Me.lblLink2
        DoCmd.Minimize
        ClickAction .Tag & .Caption
    End With
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink2_Click[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLink3_Click
' Description:  Link click event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/26/2016 - initial version
' ---------------------------------
Private Sub lblLink3_Click()
On Error GoTo Err_Handler
    
    If lblLink3.Tag = "DISABLED" Then GoTo Exit_Handler

    With Me.lblLink3
        DoCmd.Minimize
        ClickAction .Tag & .Caption
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink3_Click[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLink4_Click
' Description:  Link click event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/26/2016 - initial version
' ---------------------------------
Private Sub lblLink4_Click()
On Error GoTo Err_Handler
    
    If lblLink4.Tag = "DISABLED" Then GoTo Exit_Handler
    
    With Me.lblLink4
        DoCmd.Minimize
        ClickAction .Tag & .Caption
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink4_Click[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLink5_Click
' Description:  Link click event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/26/2016 - initial version
' ---------------------------------
Private Sub lblLink5_Click()
On Error GoTo Err_Handler
    
    If lblLink5.Tag = "DISABLED" Then GoTo Exit_Handler
    
    With Me.lblLink5
        DoCmd.Minimize
        ClickAction .Tag & .Caption
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink5_Click[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLink6_Click
' Description:  Link click event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/26/2016 - initial version
' ---------------------------------
Private Sub lblLink6_Click()
On Error GoTo Err_Handler
    
    If lblLink6.Tag = "DISABLED" Then GoTo Exit_Handler
    
    With Me.lblLink6
        DoCmd.Minimize
        ClickAction .Tag & .Caption
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink6_Click[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLink1_MouseMove
' Description:  Link mouse over event actions
' Assumptions:  Requires similar mousemove in detail to reset link color
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Drew2010, Tony M, cheekybudda June 28, 2010
'   http://www.utteraccess.com/forum/change-text-color-mouse-t1947540.html
' Source/date:  Bonnie Campbell, May 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2016 - initial version
' ---------------------------------
Private Sub lblLink1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    With Me.lblLink1
        'avoid flicker w/ if statement
'        If Not .ForeColor = lngBlue Then .ForeColor = lngBlue
'        If Not .BackStyle = acNormal Then .BackStyle = acNormal
'        If Not .BackColor = lngYellow Then .BackColor = lngYellow
        Me.LinkHighlight lblLink1
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink1_MouseMove[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLink2_MouseMove
' Description:  Link mouse over event actions
' Assumptions:  Requires similar mousemove in detail to reset link color
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Drew2010, Tony M, cheekybudda June 28, 2010
'   http://www.utteraccess.com/forum/change-text-color-mouse-t1947540.html
' Source/date:  Bonnie Campbell, May 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2016 - initial version
' ---------------------------------
Private Sub lblLink2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    With Me.lblLink2
        Me.LinkHighlight lblLink2
'        'avoid flicker w/ if statement
'        If Not .ForeColor = lngBlue Then .ForeColor = lngBlue
'        If Not .BackStyle = acNormal Then .BackStyle = acNormal
'        If Not .backColor = lngYellow Then .backColor = lngYellow
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink2_MouseMove[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLink3_MouseMove
' Description:  Link mouse over event actions
' Assumptions:  Requires similar mousemove in detail to reset link color
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Drew2010, Tony M, cheekybudda June 28, 2010
'   http://www.utteraccess.com/forum/change-text-color-mouse-t1947540.html
' Source/date:  Bonnie Campbell, May 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2016 - initial version
' ---------------------------------
Private Sub lblLink3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    With Me.lblLink3
        Me.LinkHighlight lblLink3
'        'avoid flicker w/ if statement
'        If Not .ForeColor = lngBlue Then .ForeColor = lngBlue
'        If Not .BackStyle = acNormal Then .BackStyle = acNormal
'        If Not .backColor = lngYellow Then .backColor = lngYellow
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink3_MouseMove[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLink4_MouseMove
' Description:  Link mouse over event actions
' Assumptions:  Requires similar mousemove in detail to reset link color
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Drew2010, Tony M, cheekybudda June 28, 2010
'   http://www.utteraccess.com/forum/change-text-color-mouse-t1947540.html
' Source/date:  Bonnie Campbell, May 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2016 - initial version
' ---------------------------------
Private Sub lblLink4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    With Me.lblLink4
        Me.LinkHighlight lblLink4
'        'avoid flicker w/ if statement
'        If Not .ForeColor = lngBlue Then .ForeColor = lngBlue
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink4_MouseMove[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLink5_MouseMove
' Description:  Link mouse over event actions
' Assumptions:  Requires similar mousemove in detail to reset link color
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Drew2010, Tony M, cheekybudda June 28, 2010
'   http://www.utteraccess.com/forum/change-text-color-mouse-t1947540.html
' Source/date:  Bonnie Campbell, May 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2016 - initial version
' ---------------------------------
Private Sub lblLink5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    With Me.lblLink5
        Me.LinkHighlight lblLink5
'        'avoid flicker w/ if statement
'        If Not .ForeColor = lngBlue Then .ForeColor = lngBlue
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink5_MouseMove[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblLink6_MouseMove
' Description:  Link mouse over event actions
' Assumptions:  Requires similar mousemove in detail to reset link color
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Drew2010, Tony M, cheekybudda June 28, 2010
'   http://www.utteraccess.com/forum/change-text-color-mouse-t1947540.html
' Source/date:  Bonnie Campbell, May 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2016 - initial version
' ---------------------------------
Private Sub lblLink6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    With Me.lblLink6
        Me.LinkHighlight lblLink6
'        'avoid flicker w/ if statement
'        If Not .ForeColor = lngBlue Then .ForeColor = lngBlue
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblLink6_MouseMove[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Detail_MouseMove
' Description:  Detail mouse over event actions
' Assumptions:  Similar mousemove events exist for links setting colors
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Drew2010, Tony M, cheekybudda June 28, 2010
'   http://www.utteraccess.com/forum/change-text-color-mouse-t1947540.html
' Source/date:  Bonnie Campbell, May 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2016 - initial version
' ---------------------------------
Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    Dim ctrl As Control
    Dim strLink As String
    Dim i As Integer
    
    For i = 1 To 6
    
        strLink = "lblLink" & i
    
        For Each ctrl In Me.Controls
            
            If ctrl.Name = strLink Then
            
                With ctrl
                    'avoid flicker w/ if statement
                    If Not .ForeColor = lngLtGray2 Then .ForeColor = lngLtGray2
                    If Not .BackStyle = acTransparent Then .BackStyle = acTransparent
                End With
            
            End If
            
        Next
        
    Next
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_MouseMove[Tile form])"
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
            "Error encountered (#" & Err.Number & " - Class_Initialize[Tile form])"
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
            "Error encountered (#" & Err.Number & " - Class_Terminate[Tile form])"
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
Private Sub SetHeaderColor(Color As Long)
On Error GoTo Err_Handler
    
    MsgBox "SetHeaderColor...", vbOKOnly
    Me.TileHeaderColor = Color
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetHeaderColor[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          LinkHighlight
' Description:  Actions to highlight links within a tile
' Assumptions:  ctrl has .Fore & .Back colors as well as .BackStyle
' Parameters:   ctrl - control to highlight (control)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 27, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2016 - initial version
' ---------------------------------
Public Sub LinkHighlight(ctrl As Control)
On Error GoTo Err_Handler
    
    With ctrl
        'avoid flicker w/ if statement
        If Not .ForeColor = LINK_HIGHLIGHT_TEXT Then .ForeColor = LINK_HIGHLIGHT_TEXT
        If Not .BackStyle = acNormal Then .BackStyle = acNormal
        If Not .backColor = LINK_HIGHLIGHT_BKGD Then .backColor = LINK_HIGHLIGHT_BKGD
    End With

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LinkHighlight[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          DisableLinks
' Description:  Actions to disable links within a tile
' Assumptions:  -
' Parameters:   links - comma separated links to disable (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 27, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2016 - initial version
' ---------------------------------
Public Sub DisableLinks(links As String)
On Error GoTo Err_Handler

    Dim ctrl As Control
    Dim ary() As String
    Dim strLink As String
    Dim i As Integer
    
    ary = Split(links, ",")
    
    For i = 0 To UBound(ary)
        strLink = "lblLink" & ary(i)
    
        For Each ctrl In Me.Controls
            
            If ctrl.Name = strLink Then
                'ctrl.Tag = "Disabled"
            End If
        Next
        
    Next
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableLinks[Tile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          EnableLinks
' Description:  Actions to Enable links within a tile
' Assumptions:  TileTag is passed as the first value in the CSV string
' Parameters:   links - comma separated links to Enable (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 27, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2016 - initial version
' ---------------------------------
Public Sub EnableLinks(links As String)
On Error GoTo Err_Handler

    Dim ctrl As Control
    Dim ary() As String
    Dim strLink As String, strTileTag As String
    Dim i As Integer
    
    ary = Split(links, ",")
    
    strTileTag = ary(0)
    
    For i = 1 To UBound(ary)
        strLink = "lblLink" & ary(i)
    
        For Each ctrl In Me.Controls
            
            'set tag to overall tile tag
            If ctrl.Name = strLink Then
                ctrl.Tag = strTileTag
            End If
        Next
        
    Next

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - EnableLinks[Tile form])"
    End Select
    Resume Exit_Handler
End Sub
