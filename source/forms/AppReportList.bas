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
    Width =7560
    DatasheetFontHeight =11
    ItemSuffix =37
    Right =14010
    Bottom =12105
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xb7c2305fa3cae440
    End
    RecordSource ="AppReport"
    Caption ="_List"
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
            Height =735
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
                    Width =7260
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="dirs"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4740
                    Top =420
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFormat"
                    Caption ="Format"
                    GridlineColor =10921638
                    LayoutCachedLeft =4740
                    LayoutCachedTop =420
                    LayoutCachedWidth =5985
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1440
                    Top =420
                    Width =1800
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblReport"
                    Caption ="Report"
                    GridlineColor =10921638
                    LayoutCachedLeft =1440
                    LayoutCachedTop =420
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6360
                    Top =420
                    Width =1020
                    Height =300
                    ColumnOrder =0
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =14277081
                    Name ="tbxDevMode"
                    ConditionalFormat = Begin
                        0x01000000ae000000020000000100000000000000000000001200000001000000 ,
                        0x7f7f7f00ffffff000100000000000000130000002600000001000000ffffff00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00540072007500 ,
                        0x6500000000005b007400620078004400650076004d006f00640065005d003d00 ,
                        0x460061006c007300650000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =6360
                    LayoutCachedTop =420
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =720
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =85.0
                    ConditionalFormat14 = Begin
                        0x0100020000000100000000000000010000007f7f7f00ffffff00110000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d005400720075006500 ,
                        0x0000000000000000000000000000000000000000000100000000000000010000 ,
                        0x00ffffff00ffffff00120000005b007400620078004400650076004d006f0064 ,
                        0x0065005d003d00460061006c0073006500000000000000000000000000000000 ,
                        0x000000000000
                    End
                End
            End
        End
        Begin Section
            Height =360
            Name ="Detail"
            OnMouseMove ="[Event Procedure]"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =15
                    Width =720
                    Height =300
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =15
                    LayoutCachedWidth =840
                    LayoutCachedHeight =315
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4800
                    Top =15
                    Width =1020
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxFormats"
                    GridlineColor =10921638

                    LayoutCachedLeft =4800
                    LayoutCachedTop =15
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =900
                    Width =480
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =900
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1425
                    Top =15
                    Width =3075
                    Height =300
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxReportName"
                    ControlSource ="DisplayName"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000ff00ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1425
                    LayoutCachedTop =15
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =315
                    ForeTint =50.0
                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000ff00ccffcc00000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5940
                    Width =1020
                    Height =300
                    FontSize =8
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxReport"
                    ControlSource ="ReportName"
                    GridlineColor =10921638

                    LayoutCachedLeft =5940
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7140
                    Width =300
                    Height =300
                    FontSize =8
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxFormat"
                    ControlSource ="FormatIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =7140
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3720
                    Width =1020
                    Height =300
                    FontSize =8
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxTemplate"
                    ControlSource ="ReportTemplate"
                    ConditionalFormat = Begin
                        0x01000000ae000000020000000100000000000000000000001200000001000000 ,
                        0x7f7f7f00ffffff000100000000000000130000002600000001000000ffffff00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00540072007500 ,
                        0x6500000000005b007400620078004400650076004d006f00640065005d003d00 ,
                        0x460061006c007300650000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3720
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x0100020000000100000000000000010000007f7f7f00ffffff00110000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d005400720075006500 ,
                        0x0000000000000000000000000000000000000000000100000000000000010000 ,
                        0x00ffffff00ffffff00120000005b007400620078004400650076004d006f0064 ,
                        0x0065005d003d00460061006c0073006500000000000000000000000000000000 ,
                        0x000000000000
                    End
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
' Form:         AppReportList
' Level:        Application form
' Version:      1.01
' Basis:        Dropdown form
'
' Description:  List form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, July 30, 2016
' References:   -
' Revisions:    BLC - 7/30/2016 - 1.00 - initial version
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
Public Event InvalidLabel(Value As String)
Public Event InvalidCaption(Value As String)

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
' Source/date:  Bonnie Campbell, July 30, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/30/2016 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'dev mode
    tbxDevMode = DEV_MODE
    
    'defaults
    lblTitle.Caption = ""
    lblDirections.Caption = ""
    tbxIcon.Value = StringFromCodepoint(uDocument)
    tbxIcon.ForeColor = lngDkGreen
    tbxIcon.FontSize = 14
    lblDirections.ForeColor = lngLtBlue
    
    'set hover
    
    'set data source
'    Set Me.Recordset = GetRecords("s_app_report")
        
'    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[AppReportList form])"
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
' Source/date:  Bonnie Campbell, July 30, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/30/2016 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Load[AppReportList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxReportName_Click
' Description:  Link click event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/1/2016 - initial version
'   BLC - 1/10/2018 - revised to use FetchReportRecordset() to set report
'                     recordset prior to opening report
' ---------------------------------
Private Sub tbxReportName_Click()
On Error GoTo Err_Handler
        
    DoCmd.Minimize
    
    'set report recordsource (rpt_recordset)
    SetReportRecordset (tbxTemplate.Value)
    
    'open report
    DoCmd.OpenReport tbxReport.Value, acViewReport, , , acDialog

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxReportName_Click[AppReportList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxReportName_MouseMove
' Description:  Textbox mouse over event actions
' Assumptions:  Requires similar mousemove in detail to reset textbox color
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Drew2010, Tony M, cheekybudda June 28, 2010
'   http://www.utteraccess.com/forum/change-text-color-mouse-t1947540.html
' Source/date:  Bonnie Campbell, July 30, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/30/2016 - initial version
' ---------------------------------
Private Sub tbxReportName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    With Me.tbxReportName
        'If .Tag <> "DISABLED" Then  'LinkHighlight tbxReport
        
        'avoid flicker w/ if statement
        If Not .ForeColor = LINK_HIGHLIGHT_TEXT Then .ForeColor = LINK_HIGHLIGHT_TEXT
        If Not .BackStyle = acNormalSolid Then .BackStyle = acNormalSolid
        If Not .BackColor = LINK_HIGHLIGHT_BKGD Then .BackColor = LINK_HIGHLIGHT_BKGD
        
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxReportName_MouseMove[AppReportList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Detail_MouseMove
' Description:  Detail mouse over event actions
' Assumptions:  Similar mousemove events exist for textbox links setting colors
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Drew2010, Tony M, cheekybudda June 28, 2010
'   http://www.utteraccess.com/forum/change-text-color-mouse-t1947540.html
' Source/date:  Bonnie Campbell, July 30, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/30/2016 - initial version
' ---------------------------------
Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
    
    Dim ctrl As Control
    Dim strLink As String
    Dim i As Integer
    
'    For i = 1 To 8
'
        strLink = "tbxReportName" '& i
    
        For Each ctrl In Me.Controls
            
            If ctrl.Name = strLink And ctrl.Tag <> "DISABLED" Then
            
                With ctrl
                    'avoid flicker w/ if statement
                    If Not .ForeColor = lngGray50 Then .ForeColor = lngGray50
                    If Not .BackStyle = acTransparent Then .BackStyle = acTransparent
                End With
            
            End If
            
        Next
        
'    Next
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_MouseMove[AppReportList form])"
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
' Source/date:  Bonnie Campbell, July 30, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/30/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[AppReportList form])"
    End Select
    Resume Exit_Handler
End Sub
