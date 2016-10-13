Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7440
    DatasheetFontHeight =11
    ItemSuffix =5
    Right =9480
    Bottom =8835
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0x8395ce00a4cae440
    End
    RecordSource ="AppReport"
    Caption ="_Default"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    FitToPage =1
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
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1395
            BackColor =4144959
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
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
                    Left =180
                    Top =120
                    Width =7260
                    Height =840
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Edit or Delete Records using the buttons for the record at right.\015\012Icon co"
                        "des at left identify if record may be edited/deleted."
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =960
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =4740
                    Top =1080
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFormat"
                    Caption ="Format"
                    GridlineColor =10921638
                    LayoutCachedLeft =4740
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5985
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =1440
                    Top =1080
                    Width =1800
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblReport"
                    Caption ="Report"
                    GridlineColor =10921638
                    LayoutCachedLeft =1440
                    LayoutCachedTop =1080
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            Height =420
            Name ="Detail"
            OnMouseMove ="[Event Procedure]"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Top =15
                    Width =720
                    Height =300
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedTop =15
                    LayoutCachedWidth =720
                    LayoutCachedHeight =315
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =4
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4680
                    Top =15
                    Width =1020
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxFormats"
                    ControlSource ="FormatIcon"
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =4680
                    LayoutCachedTop =15
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =315
                    ForeTint =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =4
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Width =480
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxID"
                    ControlSource ="ID"
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =780
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =315
                    ForeTint =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1305
                    Top =15
                    Width =3075
                    Height =300
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxReportName"
                    ControlSource ="DisplayName"
                    OnMouseMove ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =1305
                    LayoutCachedTop =15
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =315
                    ForeTint =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =4
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6000
                    Width =1020
                    Height =300
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxReport"
                    ControlSource ="ReportName"
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =6000
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =300
                    ForeTint =50.0
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
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
' Report:       AppReportList
' Level:        App Report
' Version:      1.00
'
' Description:  App Report List report object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, July 30, 2016
' References:
'  Allen Browne, April 2010
'  http://allenbrowne.com/ser-43.html
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

        'set the report title & caption
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

        'set the report directions
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
' Sub:          Report_Open
' Description:  report opening actions
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
Private Sub Report_Open(Cancel As Integer)
On Error GoTo Err_Handler

    lblTitle.Caption = ""
    lblDirections.Caption = ""
'    tbxIcon.Value = StringFromCodepoint(uDocument)
'    tbxIcon.ForeColor = lngDkGreen
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
            "Error encountered (#" & Err.Number & " - Report_Open[AppReportList report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Report_Load
' Description:  report loading actions
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
Private Sub Report_Load()
On Error GoTo Err_Handler

    'eliminate NULLs
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Load[AppReportList report])"
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
' ---------------------------------
Private Sub tbxReportName_Click()
On Error GoTo Err_Handler
        
    DoCmd.Minimize
    DoCmd.OpenReport tbxReport.Value, acViewNormal

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxReportName_Click[AppReportList report])"
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
        If Not .backstyle = acNormalSolid Then .backstyle = acNormalSolid
        If Not .BackColor = LINK_HIGHLIGHT_BKGD Then .BackColor = LINK_HIGHLIGHT_BKGD
        
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxReportName_MouseMove[AppReportList report])"
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
    
    Dim Ctrl As Control
    Dim strLink As String
    Dim i As Integer
    
'    For i = 1 To 8
'
        strLink = "tbxReportName" '& i
    
        For Each Ctrl In Me.Controls
            
            If Ctrl.Name = strLink Then 'And ctrl.Tag <> "DISABLED" Then
            
                With Ctrl
                    'avoid flicker w/ if statement
                    If Not .ForeColor = lngGray50 Then .ForeColor = lngGray50
                    If Not .backstyle = acTransparent Then .backstyle = acTransparent
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
            "Error encountered (#" & Err.Number & " - Detail_MouseMove[AppReportList report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Report_Close
' Description:  Report closing actions
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
Private Sub Report_Close()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Close[AppReportList report])"
    End Select
    Resume Exit_Handler
End Sub


'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------

'---------------------
' Events
'---------------------

' ---------------------------------
' Sub:          XX
' Description:  XX event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 10, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/10/2015 - initial version
' ---------------------------------
Private Sub XX()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - XX[Default report])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------
