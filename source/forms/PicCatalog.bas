﻿Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11880
    DatasheetFontHeight =11
    ItemSuffix =30
    Right =12750
    Bottom =11385
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x088d9b1cf009e540
    End
    Caption ="Photo Binder Photos"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
            Height =3780
            BackColor =4210752
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =0
            BackTint =75.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =900
                    Width =960
                    Height =315
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =16777215
                    Name ="tbxNumPix"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =900
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =1215
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Top =60
                    Width =7500
                    Height =615
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Select the desired photos"
                    GridlineColor =10921638
                    LayoutCachedTop =60
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =675
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =10800
                    Top =1560
                    Width =720
                    TabIndex =1
                    ForeColor =16711680
                    Name ="btnComment"
                    Caption ="������"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =10800
                    LayoutCachedTop =1560
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1920
                    ForeThemeColorIndex =-1
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
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =7620
                    Width =4140
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblContext"
                    Caption ="  >  -  >  -"
                    GridlineColor =10921638
                    LayoutCachedLeft =7620
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =1620
                    Width =1080
                    TabIndex =2
                    ForeColor =16711680
                    Name ="btnClearAll"
                    Caption ="Clear All"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Uncheck all photos"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =1620
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =1980
                    ForeThemeColorIndex =-1
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
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1260
                    Top =1620
                    Width =1080
                    TabIndex =3
                    ForeColor =16711680
                    Name ="btnSelectAll"
                    Caption ="Select All"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Check all photos"
                    GridlineColor =10921638

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1620
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =1980
                    ForeThemeColorIndex =-1
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
                End
                Begin Label
                    OverlapFlags =215
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
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    Left =60
                    Top =3255
                    Width =7860
                    Height =315
                    FontSize =9
                    LeftMargin =360
                    TopMargin =36
                    RightMargin =360
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =65535
                    Name ="lblMsg"
                    Caption ="Missing photo!"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =3255
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =3570
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4020
                    Top =3120
                    Width =825
                    Height =600
                    FontSize =20
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =65535
                    Name ="lblMsgIcon"
                    Caption ="⯈⯈"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =4020
                    LayoutCachedTop =3120
                    LayoutCachedWidth =4845
                    LayoutCachedHeight =3720
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =12960
            BackColor =4210752
            Name ="Detail"
            AlternateBackColor =4210752
            AlternateBackThemeColorIndex =0
            AlternateBackTint =75.0
            BackThemeColorIndex =0
            BackTint =75.0
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Width =2232
                    Height =2448
                    BorderColor =10921638
                    Name ="PicTile11"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =2352
                    LayoutCachedHeight =2568
                End
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =2688
                    Width =2232
                    Height =2448
                    TabIndex =1
                    BorderColor =10921638
                    Name ="PicTile21"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =2688
                    LayoutCachedWidth =2352
                    LayoutCachedHeight =5136
                End
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =5256
                    Width =2232
                    Height =2448
                    TabIndex =2
                    BorderColor =10921638
                    Name ="PicTile31"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =5256
                    LayoutCachedWidth =2352
                    LayoutCachedHeight =7704
                End
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =7824
                    Width =2232
                    Height =2448
                    TabIndex =3
                    BorderColor =10921638
                    Name ="PicTile41"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =7824
                    LayoutCachedWidth =2352
                    LayoutCachedHeight =10272
                End
                Begin Subform
                    OverlapFlags =85
                    Left =2460
                    Top =120
                    Width =2232
                    Height =2448
                    TabIndex =4
                    BorderColor =10921638
                    Name ="PicTile12"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =2460
                    LayoutCachedTop =120
                    LayoutCachedWidth =4692
                    LayoutCachedHeight =2568
                End
                Begin Subform
                    OverlapFlags =85
                    Left =2460
                    Top =2688
                    Width =2232
                    Height =2448
                    TabIndex =5
                    BorderColor =10921638
                    Name ="PicTile22"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =2460
                    LayoutCachedTop =2688
                    LayoutCachedWidth =4692
                    LayoutCachedHeight =5136
                End
                Begin Subform
                    OverlapFlags =85
                    Left =2460
                    Top =5256
                    Width =2232
                    Height =2448
                    TabIndex =6
                    BorderColor =10921638
                    Name ="PicTile32"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =2460
                    LayoutCachedTop =5256
                    LayoutCachedWidth =4692
                    LayoutCachedHeight =7704
                End
                Begin Subform
                    OverlapFlags =85
                    Left =2460
                    Top =7824
                    Width =2232
                    Height =2448
                    TabIndex =7
                    BorderColor =10921638
                    Name ="PicTile42"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =2460
                    LayoutCachedTop =7824
                    LayoutCachedWidth =4692
                    LayoutCachedHeight =10272
                End
                Begin Subform
                    OverlapFlags =85
                    Left =4800
                    Top =120
                    Width =2232
                    Height =2448
                    TabIndex =8
                    BorderColor =10921638
                    Name ="PicTile13"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =4800
                    LayoutCachedTop =120
                    LayoutCachedWidth =7032
                    LayoutCachedHeight =2568
                End
                Begin Subform
                    OverlapFlags =85
                    Left =4800
                    Top =2688
                    Width =2232
                    Height =2448
                    TabIndex =9
                    BorderColor =10921638
                    Name ="PicTile23"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =4800
                    LayoutCachedTop =2688
                    LayoutCachedWidth =7032
                    LayoutCachedHeight =5136
                End
                Begin Subform
                    OverlapFlags =85
                    Left =4800
                    Top =5256
                    Width =2232
                    Height =2448
                    TabIndex =10
                    BorderColor =10921638
                    Name ="PicTile33"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =4800
                    LayoutCachedTop =5256
                    LayoutCachedWidth =7032
                    LayoutCachedHeight =7704
                End
                Begin Subform
                    OverlapFlags =85
                    Left =4800
                    Top =7824
                    Width =2232
                    Height =2448
                    TabIndex =11
                    BorderColor =10921638
                    Name ="PicTile43"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =4800
                    LayoutCachedTop =7824
                    LayoutCachedWidth =7032
                    LayoutCachedHeight =10272
                End
                Begin Subform
                    OverlapFlags =85
                    Left =7140
                    Top =120
                    Width =2232
                    Height =2448
                    TabIndex =12
                    BorderColor =10921638
                    Name ="PicTile14"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =7140
                    LayoutCachedTop =120
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =2568
                End
                Begin Subform
                    OverlapFlags =85
                    Left =7140
                    Top =2688
                    Width =2232
                    Height =2448
                    TabIndex =13
                    BorderColor =10921638
                    Name ="PicTile24"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =7140
                    LayoutCachedTop =2688
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =5136
                End
                Begin Subform
                    OverlapFlags =85
                    Left =7140
                    Top =5256
                    Width =2232
                    Height =2448
                    TabIndex =14
                    BorderColor =10921638
                    Name ="PicTile34"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =7140
                    LayoutCachedTop =5256
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =7704
                End
                Begin Subform
                    OverlapFlags =85
                    Left =7140
                    Top =7824
                    Width =2232
                    Height =2448
                    TabIndex =15
                    BorderColor =10921638
                    Name ="PicTile44"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =7140
                    LayoutCachedTop =7824
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =10272
                End
                Begin Subform
                    OverlapFlags =85
                    Left =9480
                    Top =120
                    Width =2232
                    Height =2448
                    TabIndex =16
                    BorderColor =10921638
                    Name ="PicTile15"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =120
                    LayoutCachedWidth =11712
                    LayoutCachedHeight =2568
                End
                Begin Subform
                    OverlapFlags =85
                    Left =9480
                    Top =2688
                    Width =2232
                    Height =2448
                    TabIndex =17
                    BorderColor =10921638
                    Name ="PicTile25"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =2688
                    LayoutCachedWidth =11712
                    LayoutCachedHeight =5136
                End
                Begin Subform
                    OverlapFlags =85
                    Left =9480
                    Top =5256
                    Width =2232
                    Height =2448
                    TabIndex =18
                    BorderColor =10921638
                    Name ="PicTile35"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =5256
                    LayoutCachedWidth =11712
                    LayoutCachedHeight =7704
                End
                Begin Subform
                    OverlapFlags =85
                    Left =9480
                    Top =7824
                    Width =2232
                    Height =2448
                    TabIndex =19
                    BorderColor =10921638
                    Name ="PicTile45"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =7824
                    LayoutCachedWidth =11712
                    LayoutCachedHeight =10272
                End
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =10380
                    Width =2232
                    Height =2448
                    TabIndex =20
                    BorderColor =10921638
                    Name ="PicTile51"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =10380
                    LayoutCachedWidth =2352
                    LayoutCachedHeight =12828
                End
                Begin Subform
                    OverlapFlags =85
                    Left =2460
                    Top =10380
                    Width =2232
                    Height =2448
                    TabIndex =21
                    BorderColor =10921638
                    Name ="PicTile52"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =2460
                    LayoutCachedTop =10380
                    LayoutCachedWidth =4692
                    LayoutCachedHeight =12828
                End
                Begin Subform
                    OverlapFlags =85
                    Left =4800
                    Top =10380
                    Width =2232
                    Height =2448
                    TabIndex =22
                    BorderColor =10921638
                    Name ="PicTile53"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =4800
                    LayoutCachedTop =10380
                    LayoutCachedWidth =7032
                    LayoutCachedHeight =12828
                End
                Begin Subform
                    OverlapFlags =85
                    Left =7140
                    Top =10380
                    Width =2232
                    Height =2448
                    TabIndex =23
                    BorderColor =10921638
                    Name ="PicTile54"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =7140
                    LayoutCachedTop =10380
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =12828
                End
                Begin Subform
                    OverlapFlags =85
                    Left =9480
                    Top =10380
                    Width =2232
                    Height =2448
                    TabIndex =24
                    BorderColor =10921638
                    Name ="PicTile55"
                    SourceObject ="Form.PicTile"
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =10380
                    LayoutCachedWidth =11712
                    LayoutCachedHeight =12828
                End
            End
        End
        Begin FormFooter
            Height =360
            BackColor =4210752
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =0
            BackTint =75.0
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
' Form:         PicCatalog
' Level:        Framework form
' Version:      1.00
'
' Description:  PicCatalog form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 12/18/2017
' References:   -
' Revisions:    BLC - 12/18/2017 - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

Private m_SaveOK As Boolean 'ok to save record (prevents bound form from immediately updating)

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)

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
' Description:  form opening actions
' Assumptions:  OpenArgs passes only the calling form name
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/18/2017 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "Main"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize calling form
    ToggleForm Me.CallingForm, -1
    
    'set context - based on TempVars
    lblContext.ForeColor = lngLime
    lblContext.Caption = GetContext()
    
    Title = "Photo Binder Photos"
    lblTitle.Caption = "" 'hide second title
    Directions = "Select the desired photos"
    lblDirections.ForeColor = lngLtBlue
    btnComment.Caption = StringFromCodepoint(uComment)
    btnComment.ForeColor = lngBlue
    
    'set hint
    
    'set hover
    btnComment.HoverColor = lngGreen
      
    'defaults
    btnComment.Enabled = False
    lblMsgIcon.Caption = ""
    lblMsg.Caption = ""
  
    '# of photos
    tbxNumPix = Me.Recordset
    
    'clear form datasource in case it was saved (to keep unbound)
    Me.RecordSource = ""
    
    'set data sources
'    Set cbxPhotoType.Recordset = GetRecords("s_temp_photos")
    
    'initialize values
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[PicCatalog form])"
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
' Source/date:  Bonnie Campbell, December 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/18/2017 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Load[PicCatalog form])"
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
' Source/date:  Bonnie Campbell, December 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/18/2017 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[PicCatalog form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_BeforeUpdate
' Description:  form current actions
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
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler
              
    If Not m_SaveOK Then
        Cancel = True
    End If
    'Cancel = True

'    Me.lblMsg.Caption = StringFromCodepoint(uRArrow) & " Updating record..."

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[PicCatalog form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_AfterUpdate
' Description:  form after update actions
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
Private Sub Form_AfterUpdate()
On Error GoTo Err_Handler
              
'    Me.lblMsg.Caption = ""

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[PicCatalog form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnClearAll_Click
' Description:  clear all button click event actions
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
Private Sub btnClearAll_Click()
On Error GoTo Err_Handler
    
    'check none
    ToggleChecks False

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClearAll_Click[PicPicTile form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSelectAll_Click
' Description:  select all click event actions
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
Private Sub btnSelectAll_Click()
On Error GoTo Err_Handler
    
    'check all
    ToggleChecks True
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSelectAll_Click[PicPicTile form])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------------------------
' Sub:          btnSave_Click
' Description:  Save button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 9/1/2016  - cleanup commented code
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    'set enable btnSave_Click save
    m_SaveOK = True
    
'    'pre-save form
'    Me![list].Form.Dirty = False
    
    UpsertRecord Me
    
    Me![list].Form.Requery
    
    'revert to disable non-btnSave_Click save
    m_SaveOK = False
    
    'clear fields
    ClearForm Me
        
'    cbxLocation.ControlSource = ""  'clear from Location_ID
'    cbxLocation.Value = ""
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[PicCatalog form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnComment_Click
' Description:  Undo button click actions
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
Private Sub btnComment_Click()
On Error GoTo Err_Handler
    
    'open comment form
'    DoCmd.OpenForm "Comment", acNormal, , , , , "event|" & tbxID & "|255"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[PicCatalog form])"
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
' Source/date:  Bonnie Campbell, December 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/18/2017 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Close[PicCatalog form])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          ToggleChecks
' Description:  Toggles checkboxes in subforms to checked or unchecked
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
Private Sub ToggleChecks(selection As Boolean)
On Error GoTo Err_Handler

    Dim ctrl As Control
    Dim sctrl As Control
    
    'iterate through all subforms
    For Each ctrl In Me.Controls
'Debug.Print ctrl.Name

        'check for subform (control type 112)
        If ctrl.ControlType = acSubform Then
            
            'iterate through subform controls
            For Each sctrl In ctrl.Form.Controls
            
                Select Case sctrl.ControlType
                    Case acCheckBox
                        If sctrl.Name = "chkSelect" Then _
                            sctrl = selection
                    'Case acTextBox
                    Case acLabel
                        If sctrl.Name = "lblName" Then _
                            sctrl.ForeColor = IIf(selection = True, lngGreen, lngLtTextGray)
                    Case acImage
                        If sctrl.Name = "imgPhoto" Then _
                            sctrl.BorderColor = IIf(selection = True, lngGreen, lngLtBgdGray)
                End Select
                
            Next
            
        End If
    
    Next
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[PicCatalog form])"
    End Select
    Resume Exit_Handler
End Sub