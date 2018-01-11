Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7440
    DatasheetFontHeight =11
    ItemSuffix =11
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    Filter ="[Status] = 'Opened'"
    RecSrcDt = Begin
        0x2c44fdfacc0ce540
    End
    RecordSource ="Task"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
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
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="TaskType"
        End
        Begin BreakLevel
            ControlSource ="TaskType"
        End
        Begin BreakLevel
            ControlSource ="=[Priority]"
        End
        Begin BreakLevel
            ControlSource ="RequestDate"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1035
            BackColor =4144959
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =3480
                    Height =300
                    FontSize =12
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    Left =180
                    Top =420
                    Width =7260
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="directions"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =720
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =120
                    Top =720
                    Width =1245
                    Height =315
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblPriority"
                    Caption ="Priority"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =720
                    LayoutCachedWidth =1365
                    LayoutCachedHeight =1035
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =1500
                    Top =720
                    Width =1080
                    Height =315
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTask"
                    Caption ="Task"
                    GridlineColor =10921638
                    LayoutCachedLeft =1500
                    LayoutCachedTop =720
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =1035
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =6180
                    Top =720
                    Width =1245
                    Height =315
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblRequestDate"
                    Caption ="Requested"
                    GridlineColor =10921638
                    LayoutCachedLeft =6180
                    LayoutCachedTop =720
                    LayoutCachedWidth =7425
                    LayoutCachedHeight =1035
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =3
                    Left =3480
                    Top =60
                    Width =3900
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblContext"
                    Caption ="context"
                    GridlineColor =10921638
                    LayoutCachedLeft =3480
                    LayoutCachedTop =60
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3960
                    Top =720
                    Width =1140
                    Height =300
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxDevMode"
                    ControlSource ="=[DEV_MODE]"
                    ConditionalFormat = Begin
                        0x01000000f0000000030000000100000000000000000000001900000001000000 ,
                        0xff000000ffffff0001000000000000001a0000002f00000001000000ff990000 ,
                        0xffffff0001000000000000003000000047000000010000000000ff00ffffff00 ,
                        0x5b007400620078005000720069006f0072006900740079005d003d0022004300 ,
                        0x7200690074006900630061006c002200000000005b0074006200780050007200 ,
                        0x69006f0072006900740079005d003d0022004800690067006800220000000000 ,
                        0x5b007400620078005000720069006f0072006900740079005d003d0022004d00 ,
                        0x65006400690075006d00220000000000
                    End
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =3960
                    LayoutCachedTop =720
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =1020
                    ForeTint =50.0
                    ConditionalFormat14 = Begin
                        0x010004000000010000000000000001000000ff000000ffffff00180000005b00 ,
                        0x7400620078005000720069006f0072006900740079005d003d00220043007200 ,
                        0x690074006900630061006c002200000000000000000000000000000000000000 ,
                        0x000000010000000000000001000000ff990000ffffff00140000005b00740062 ,
                        0x0078005000720069006f0072006900740079005d003d00220048006900670068 ,
                        0x0022000000000000000000000000000000000000000000000100000000000000 ,
                        0x010000000000ff00ffffff00160000005b007400620078005000720069006f00 ,
                        0x72006900740079005d003d0022004d0065006400690075006d00220000000000 ,
                        0x0000000000000000000000000000000000010000000000000001000000009900 ,
                        0x00ffffff00130000005b007400620078005000720069006f0072006900740079 ,
                        0x005d003d0022004c006f00770022000000000000000000000000000000000000 ,
                        0x00000000
                    End
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
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            BackColor =15921906
            Name ="GroupHeader0"
            AutoHeight =255
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Width =2160
                    Height =300
                    FontSize =12
                    BorderColor =10921638
                    ForeColor =5855577
                    Name ="tbxFormats"
                    ControlSource ="TaskType"
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =60
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =300
                    ForeTint =65.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =315
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
                    Left =240
                    Top =15
                    Width =720
                    Height =300
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxPriority"
                    ControlSource ="Priority"
                    ConditionalFormat = Begin
                        0x010000009c000000030000000000000002000000000000000b00000001000000 ,
                        0xff000000ffffff0000000000020000000c0000001300000001000000ff990000 ,
                        0xffffff000000000002000000140000001d000000010000000000ff00ffffff00 ,
                        0x220043007200690074006900630061006c002200000000002200480069006700 ,
                        0x680022000000000022004d0065006400690075006d00220000000000
                    End
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =240
                    LayoutCachedTop =15
                    LayoutCachedWidth =960
                    LayoutCachedHeight =315
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                    ConditionalFormat14 = Begin
                        0x010004000000000000000200000001000000ff000000ffffff000a0000002200 ,
                        0x43007200690074006900630061006c0022000000000000000000000000000000 ,
                        0x00000000000000000000000200000001000000ff990000ffffff000600000022 ,
                        0x0048006900670068002200000000000000000000000000000000000000000000 ,
                        0x0000000002000000010000000000ff00ffffff000800000022004d0065006400 ,
                        0x690075006d002200000000000000000000000000000000000000000000000000 ,
                        0x00020000000100000000990000ffffff000500000022004c006f007700220000 ,
                        0x0000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1020
                    Width =480
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxID"
                    ControlSource ="ID"
                    ConditionalFormat = Begin
                        0x01000000f0000000030000000100000000000000000000001900000001000000 ,
                        0xff000000ffffff0001000000000000001a0000002f00000001000000ff990000 ,
                        0xffffff0001000000000000003000000047000000010000000000ff00ffffff00 ,
                        0x5b007400620078005000720069006f0072006900740079005d003d0022004300 ,
                        0x7200690074006900630061006c002200000000005b0074006200780050007200 ,
                        0x69006f0072006900740079005d003d0022004800690067006800220000000000 ,
                        0x5b007400620078005000720069006f0072006900740079005d003d0022004d00 ,
                        0x65006400690075006d00220000000000
                    End
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =1020
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =315
                    ForeTint =50.0
                    ConditionalFormat14 = Begin
                        0x010006000000010000000000000001000000ff000000ffffff00180000005b00 ,
                        0x7400620078005000720069006f0072006900740079005d003d00220043007200 ,
                        0x690074006900630061006c002200000000000000000000000000000000000000 ,
                        0x000000010000000000000001000000ff990000ffffff00140000005b00740062 ,
                        0x0078005000720069006f0072006900740079005d003d00220048006900670068 ,
                        0x0022000000000000000000000000000000000000000000000100000000000000 ,
                        0x010000000000ff00ffffff00160000005b007400620078005000720069006f00 ,
                        0x72006900740079005d003d0022004d0065006400690075006d00220000000000 ,
                        0x0000000000000000000000000000000000010000000000000001000000009900 ,
                        0x00ffffff00130000005b007400620078005000720069006f0072006900740079 ,
                        0x005d003d0022004c006f00770022000000000000000000000000000000000000 ,
                        0x00000000010000000000000001000000ffffff00ffffff00120000005b007400 ,
                        0x620078004400650076004d006f00640065005d003d00460061006c0073006500 ,
                        0x0000000000000000000000000000000000000000000100000000000000010000 ,
                        0x007f7f7f00ffffff00110000005b007400620078004400650076004d006f0064 ,
                        0x0065005d003d0054007200750065000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Width =3075
                    Height =300
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxTask"
                    ControlSource ="Task"
                    OnMouseMove ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000f0000000030000000100000000000000000000001900000001000000 ,
                        0xff000000ffffff0001000000000000001a0000002f00000001000000ff990000 ,
                        0xffffff0001000000000000003000000047000000010000000000ff00ffffff00 ,
                        0x5b007400620078005000720069006f0072006900740079005d003d0022004300 ,
                        0x7200690074006900630061006c002200000000005b0074006200780050007200 ,
                        0x69006f0072006900740079005d003d0022004800690067006800220000000000 ,
                        0x5b007400620078005000720069006f0072006900740079005d003d0022004d00 ,
                        0x65006400690075006d00220000000000
                    End
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =1680
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =300
                    ForeTint =50.0
                    ConditionalFormat14 = Begin
                        0x010004000000010000000000000001000000ff000000ffffff00180000005b00 ,
                        0x7400620078005000720069006f0072006900740079005d003d00220043007200 ,
                        0x690074006900630061006c002200000000000000000000000000000000000000 ,
                        0x000000010000000000000001000000ff990000ffffff00140000005b00740062 ,
                        0x0078005000720069006f0072006900740079005d003d00220048006900670068 ,
                        0x0022000000000000000000000000000000000000000000000100000000000000 ,
                        0x010000000000ff00ffffff00160000005b007400620078005000720069006f00 ,
                        0x72006900740079005d003d0022004d0065006400690075006d00220000000000 ,
                        0x0000000000000000000000000000000000010000000000000001000000009900 ,
                        0x00ffffff00130000005b007400620078005000720069006f0072006900740079 ,
                        0x005d003d0022004c006f00770022000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6300
                    Width =1020
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxRequestDate"
                    ControlSource ="RequestDate"
                    ConditionalFormat = Begin
                        0x01000000f0000000030000000100000000000000000000001900000001000000 ,
                        0xff000000ffffff0001000000000000001a0000002f00000001000000ff990000 ,
                        0xffffff0001000000000000003000000047000000010000000000ff00ffffff00 ,
                        0x5b007400620078005000720069006f0072006900740079005d003d0022004300 ,
                        0x7200690074006900630061006c002200000000005b0074006200780050007200 ,
                        0x69006f0072006900740079005d003d0022004800690067006800220000000000 ,
                        0x5b007400620078005000720069006f0072006900740079005d003d0022004d00 ,
                        0x65006400690075006d00220000000000
                    End
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =6300
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =300
                    ForeTint =50.0
                    ConditionalFormat14 = Begin
                        0x010004000000010000000000000001000000ff000000ffffff00180000005b00 ,
                        0x7400620078005000720069006f0072006900740079005d003d00220043007200 ,
                        0x690074006900630061006c002200000000000000000000000000000000000000 ,
                        0x000000010000000000000001000000ff990000ffffff00140000005b00740062 ,
                        0x0078005000720069006f0072006900740079005d003d00220048006900670068 ,
                        0x0022000000000000000000000000000000000000000000000100000000000000 ,
                        0x010000000000ff00ffffff00160000005b007400620078005000720069006f00 ,
                        0x72006900740079005d003d0022004d0065006400690075006d00220000000000 ,
                        0x0000000000000000000000000000000000010000000000000001000000009900 ,
                        0x00ffffff00130000005b007400620078005000720069006f0072006900740079 ,
                        0x005d003d0022004c006f00770022000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
            End
        End
        Begin PageFooter
            Height =360
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =1
                    Left =60
                    Top =60
                    Width =3060
                    Height =300
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="lblGenerated"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                End
                Begin Label
                    TextAlign =3
                    Left =4320
                    Top =60
                    Width =3060
                    Height =300
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="lblPaging"
                    GridlineColor =10921638
                    LayoutCachedLeft =4320
                    LayoutCachedTop =60
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                End
            End
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
' Report:       Tasks
' Level:        App Report
' Version:      1.00
'
' Description:  App Report List report object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, January 10, 2018
' References:
'  Allen Browne, April 2010
'  http://allenbrowne.com/ser-43.html
' Revisions:    BLC - 1/10/2018 - 1.00 - initial version
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
' Source/date:  Bonnie Campbell, January 10, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/10/2018 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'defaults
    lblTitle.Caption = "Tasks"
    lblDirections.Caption = ""
    lblContext.Caption = GetContext()
'    tbxIcon.Value = StringFromCodepoint(uDocument)
    lblGenerated.Caption = "Printed: " & Date
    lblPaging.Caption = "Page " & Report.Page & " of " & Report.Pages
    
    'colors
    Me.ReportHeader.BackColor = lngLtGray
'    tbxIcon.ForeColor = lngDkGreen
    lblTitle.ForeColor = lngBlack
    lblContext.ForeColor = lngBlack
    lblDirections.ForeColor = lngLtBlue
    lblPriority.ForeColor = lngBlack
    lblTask.ForeColor = lngBlack
    lblRequestDate.ForeColor = lngBlack
    
    'format
    lblTitle.FontBold = True
    lblContext.FontBold = False
    
    'set hover
    
    
    'set data source > use Name property of recordset
    '                >> gives table, query, SQL string recordset was opened with
    'Set Me.Recordset = GetRecords("s_tasks") 'Error #32585: feature only available in ADP?
    'Set Me.RecordSource =  GetRecords("s_tasks") 'Error: invalid use of property
     Me.RecordSource = rpt_recordset.Name
        
     'Me.tbxIcon = "[Priority]"  '<< cannot assign
     
     Me.Filter = "[Status] = 'Opened'"
     Me.FilterOn = True
     Me.FilterOnLoad = True
        
'    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[Tasks report])"
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
' Source/date:  Bonnie Campbell, January 10, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/10/2018 - initial version
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
            "Error encountered (#" & Err.Number & " - Report_Load[Tasks report])"
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
        
    'DoCmd.Minimize
    'DoCmd.OpenReport tbxReport.Value, acViewNormal

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxReportName_Click[Tasks report])"
    End Select
    Resume Exit_Handler
End Sub

'' ---------------------------------
'' Sub:          tbxReportName_MouseMove
'' Description:  Textbox mouse over event actions
'' Assumptions:  Requires similar mousemove in detail to reset textbox color
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:
''   Drew2010, Tony M, cheekybudda June 28, 2010
''   http://www.utteraccess.com/forum/change-text-color-mouse-t1947540.html
'' Source/date:  Bonnie Campbell, January 10, 2018 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 1/10/2018 - initial version
'' ---------------------------------
'Private Sub tbxReportName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error GoTo Err_Handler
'
'    With Me.tbxReportName
'        'If .Tag <> "DISABLED" Then  'LinkHighlight tbxReport
'
'        'avoid flicker w/ if statement
'        If Not .ForeColor = LINK_HIGHLIGHT_TEXT Then .ForeColor = LINK_HIGHLIGHT_TEXT
'        If Not .BackStyle = acNormalSolid Then .BackStyle = acNormalSolid
'        If Not .BackColor = LINK_HIGHLIGHT_BKGD Then .BackColor = LINK_HIGHLIGHT_BKGD
'
'    End With
'
'Exit_Handler:
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tbxReportName_MouseMove[Tasks report])"
'    End Select
'    Resume Exit_Handler
'End Sub

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
' Source/date:  Bonnie Campbell, January 10, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/10/2018 - initial version
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
            
            If ctrl.Name = strLink Then 'And ctrl.Tag <> "DISABLED" Then
            
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
            "Error encountered (#" & Err.Number & " - Detail_MouseMove[Tasks report])"
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
' Source/date:  Bonnie Campbell, January 10, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/10/2018 - initial version
' ---------------------------------
Private Sub Report_Close()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Close[Tasks report])"
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
