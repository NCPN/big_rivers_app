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
    ScrollBars =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2232
    DatasheetFontHeight =11
    ItemSuffix =38
    Left =13095
    Top =18060
    Right =15315
    Bottom =20490
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x06dd372434a7e440
    End
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnClick ="[Event Procedure]"
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin WebBrowser
            OldBorderStyle =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            BackColor =5855577
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =0
            BackTint =65.0
        End
        Begin Section
            Height =2910
            BackColor =5855577
            Name ="Detail"
            AlternateBackColor =5855577
            AlternateBackThemeColorIndex =0
            AlternateBackTint =65.0
            BackThemeColorIndex =0
            BackTint =65.0
            Begin
                Begin Image
                    OldBorderStyle =1
                    Left =120
                    Top =300
                    Width =2016
                    Height =1728
                    BorderColor =12566463
                    Name ="imgPhoto"
                    OnClick ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =300
                    LayoutCachedWidth =2136
                    LayoutCachedHeight =2028
                    TabIndex =1
                    BorderThemeColorIndex =-1
                    BorderShade =75.0
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =300
                    BorderColor =10921638
                    Name ="chkSelect"
                    DefaultValue ="0"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Select photo"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =360
                    LayoutCachedHeight =300
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =2160
                    Width =2016
                    Height =315
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =14277081
                    Name ="lblName"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =2160
                    LayoutCachedWidth =2136
                    LayoutCachedHeight =2475
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ForeShade =85.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =2595
                    Width =2016
                    Height =315
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =14277081
                    Name ="lblFullPath"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =2595
                    LayoutCachedWidth =2136
                    LayoutCachedHeight =2910
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =85.0
                End
                Begin Label
                    OverlapFlags =247
                    Left =1260
                    Width =366
                    Height =315
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =14277081
                    Name ="lblPhotoType"
                    GridlineColor =10921638
                    LayoutCachedLeft =1260
                    LayoutCachedWidth =1626
                    LayoutCachedHeight =315
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ForeShade =85.0
                End
                Begin Label
                    OverlapFlags =247
                    Left =1680
                    Width =456
                    Height =315
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =14277081
                    Name ="lblID"
                    GridlineColor =10921638
                    LayoutCachedLeft =1680
                    LayoutCachedWidth =2136
                    LayoutCachedHeight =315
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ForeShade =85.0
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
' Form:         PicTile
' Level:        Framework form
' Version:      1.01
'
' Description:  PicTile form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 12/18/2017
' References:   -
' Revisions:    BLC - 12/18/2017 - 1.00 - initial version
'               BLC - 12/29/2017 - 1.01 - revise to accommodate PicPhotos subform
' =================================

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_TileTag As String
Private m_PicCaption As String
Private m_PicAction As String
Private m_TileHeaderColor As Long
Private m_TitleFontColor As Long
Private m_TileVisible As Byte

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
    'lblTitle.Caption = m_Title
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let TileTag(Value As String)
    m_TileTag = Value
'    lblLink1.Tag = m_TileTag
'    lblLink2.Tag = m_TileTag
'    lblLink3.Tag = m_TileTag
'    lblLink4.Tag = m_TileTag
'    lblLink5.Tag = m_TileTag
'    lblLink6.Tag = m_TileTag
End Property

Public Property Get TileTag() As String
    TileTag = m_TileTag
End Property

Public Property Get PicCaption() As String
    PicCaption = m_PicCaption
End Property

Public Property Let PicCaption(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Pic"
    m_PicCaption = Value
'    lblPic.Caption = m_PicCaption
End Property

Public Property Get PicAction() As String
    PicAction = m_PicAction
End Property

Public Property Let PicAction(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Pic"
    m_PicAction = Value
End Property

Public Property Let TitleFontColor(Value As Long)
    m_TitleFontColor = Value
    'lblTitle.ForeColor = m_TitleFontColor
End Property

Public Property Get TitleFontColor() As Long
    TitleFontColor = m_TitleFontColor
End Property

Public Property Let TileHeaderColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen '"#3F3F3F"
    m_TileHeaderColor = Value
    FormHeader.BackColor = m_TileHeaderColor
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

Public Property Get TileVisible() As Byte
    TileVisible = m_TileVisible
End Property

Public Property Let TileVisible(Value As Byte)
    m_TileVisible = Value
    Me.Visible = m_TileVisible
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

    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[PicTile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Click
' Description:  Form click event
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
Private Sub Form_Click()
On Error GoTo Err_Handler

    'Call chkSelect_Click
    'toggle the opposite of the current checkbox selection
    ToggleSelect Not chkSelect
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Click[PicTile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          imgPhoto_DblClick
' Description:  imgPhoto double click event actions
' Assumptions:  lblFullPath contains full image path
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/18/2017 - initial version
' ---------------------------------
Private Sub imgPhoto_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    If FileExists(lblFullPath.Caption) Then
        DoCmd.OpenForm "PhotoEnlarge", acNormal, , , , , lblFullPath.Caption 'TempVars("FullPhotoPath")
        Me.Parent!lblMsg.Caption = ""
        Me.Parent!lblMsg.ForeColor = lngRobinEgg
        Me.Parent!lblMsgIcon.Caption = ""
        Me.Parent!lblMsgIcon.ForeColor = lngRobinEgg
    Else
        Me.Parent!lblMsg.ForeColor = lngYellow
        Me.Parent!lblMsg.Caption = "Missing photo!"
        Me.Parent!lblMsgIcon.ForeColor = lngYellow
        Me.Parent!lblMsgIcon.Caption = StringFromCodepoint(uRTriangle) & StringFromCodepoint(uRTriangle)
    End If
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - imgPhoto_DblClick[PicPicTile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          chkSelect_Click
' Description:  checkbox click event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/18/2017 - initial version
' ---------------------------------
Private Sub chkSelect_Click()
On Error GoTo Err_Handler
    
    ToggleSelect chkSelect
    
'    If chkSelect = True Then
'        imgPhoto.BorderColor = lngGreen
'        lblName.ForeColor = lngGreen
'    Else
'        imgPhoto.BorderColor = lngLtBgdGray
'        lblName.ForeColor = lngLtTextGray
'    End If
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkSelect_Click[PicPicTile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          imgPhoto_Click
' Description:  imgPhoto click event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/18/2017 - initial version
' ---------------------------------
Private Sub imgPhoto_Click()
On Error GoTo Err_Handler
    
    'toggle the opposite of the current checkbox selection
    ToggleSelect Not chkSelect
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - imgPhoto_Click[PicPicTile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblName_Click
' Description:  lblName click event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/18/2017 - initial version
' ---------------------------------
Private Sub lblName_Click()
On Error GoTo Err_Handler
    
    'toggle the opposite of the current checkbox selection
    ToggleSelect Not chkSelect
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblName_Click[PicPicTile form])"
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

    'MsgBox "Initializing...", vbOKOnly
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[PicTile form])"
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
    
    'MsgBox "Terminating...", vbOKOnly
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[PicTile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ToggleSelect
' Description:  Toggle photo selection
' Assumptions:  -
' Parameters:   selection - whether or not item is selected (boolean)
'                           true = selected, false = not selected
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 28, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/28/2015 - initial version
'   BLC - 12/29/2017 - revise to accommodate PicPhotos subform
' ---------------------------------
Private Sub ToggleSelect(selection As Boolean)
On Error GoTo Err_Handler
    
    'set checkbox
    chkSelect = selection
    
    'define grandparent form
    Dim frm As Form
    Set frm = Me.Parent.Parent '.SelPhotos 'Form.SelPhotos
    
    If selection = True Then
        imgPhoto.BorderColor = lngGreen
        lblName.ForeColor = lngGreen
        
        'add to PicCatalog form's collection
        frm.SelPhoto = lblID.Caption
        
        'set tempvars
'        SetTempVar "SelectedPhotos", P
        
'        Me.Parent!tbxIDs = IIf(Len(Me.Parent!tbxIDs) > 0, Me.Parent!tbxIDs & "," & lblID.Caption, lblID.Caption)
    Else
        imgPhoto.BorderColor = lngLtBgdGray
        lblName.ForeColor = lngLtTextGray
        
        'remove from list > remove single comma, replace double comma with single, remove #
'        Me.Parent!tbxIDs = _
'        IIf(Left(Me.Parent!tbxIDs, 1) = ",", "", _
'        IIf(Me.Parent!tbxIDs = _
'        Replace(Replace(Me.Parent!tbxIDs, lblID.Caption, ""), ",,", ","), _
'        "", Replace(Replace(Me.Parent!tbxIDs, lblID.Caption, ""), ",,", ",")))
        Dim i As Long
        
        If frm.SelPhotos.Count > 0 Then
            For i = 1 To frm.SelPhotos.Count
                If frm.SelPhotos.Item(i) = lblID.Caption Then
                    'remove from PicCatalog form's collection (must use index since collection is unkeyed)
                    frm.SelPhotos.Remove i 'lblID.Caption
                    Exit For
                End If
            Next
        End If
    End If
    
    'print the collection
    Debug.Print "Selected Photos: "
    If frm.SelPhotos.Count > 0 Then
        Dim pics As String
        For i = 1 To frm.SelPhotos.Count
            pics = pics & frm.SelPhotos.Item(i) & Space(2)
        Next
        Debug.Print pics
    Else
        Debug.Print 0
    End If
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleSelect[PicTile form])"
    End Select
    Resume Exit_Handler
End Sub
