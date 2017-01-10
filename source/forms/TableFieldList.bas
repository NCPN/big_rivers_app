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
    Width =3720
    DatasheetFontHeight =11
    ItemSuffix =46
    Right =10260
    Bottom =7815
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x6375f9ebf1d4e440
    End
    RecordSource ="usys_temp_rs"
    Caption ="_List"
    OnCurrent ="[Event Procedure]"
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
            CanGrow = NotDefault
            Height =360
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =72
                    Top =29
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Table fields"
                    GridlineColor =10921638
                    LayoutCachedLeft =72
                    LayoutCachedTop =29
                    LayoutCachedWidth =3552
                    LayoutCachedHeight =329
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =3000
                    Width =480
                    Height =360
                    FontSize =14
                    BorderColor =10921638
                    ForeColor =65535
                    Name ="lblLinkedTable"
                    Caption ="í ½í´—"
                    ControlTipText ="Linked table"
                    GridlineColor =10921638
                    LayoutCachedLeft =3000
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2280
                    Width =480
                    Height =315
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =2366701
                    Name ="tbxCSVPseudoRecord"
                    ControlSource ="=[Forms]![ImportMap].[Controls]![tbxCSVRecord]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2280
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =315
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Width =3720
                    Height =315
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxHighlight"
                    ConditionalFormat = Begin
                        0x010000000a010000020000000100000000000000000000003500000001000000 ,
                        0x0000ff00ccff99000100000000000000360000005400000001000000ffffff00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004d0065002e005200650063006f00720064007300650074002e0041006200 ,
                        0x73006f006c0075007400650050006f0073006900740069006f006e005d003d00 ,
                        0x5b00740062007800430053005600500073006500750064006f00520065006300 ,
                        0x6f00720064005d000000000049004900660028005b0074006200780046006900 ,
                        0x65006c0064004e0061006d0065005d003c003e0022004900440022002c003100 ,
                        0x2c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedWidth =3720
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x0100020000000100000000000000010000000000ff00ccff9900340000005b00 ,
                        0x4d0065002e005200650063006f00720064007300650074002e00410062007300 ,
                        0x6f006c0075007400650050006f0073006900740069006f006e005d003d005b00 ,
                        0x740062007800430053005600500073006500750064006f005200650063006f00 ,
                        0x720064005d000000000000000000000000000000000000000000000100000000 ,
                        0x00000001000000ffffff00ffffff001d00000049004900660028005b00740062 ,
                        0x0078004600690065006c0064004e0061006d0065005d003c003e002200490044 ,
                        0x0022002c0031002c003000290000000000000000000000000000000000000000 ,
                        0x0000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Height =315
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxFieldType"
                    ControlSource ="ColType"
                    ConditionalFormat = Begin
                        0x010000009c000000010000000100000000000000000000001d00000001000100 ,
                        0xff000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b007400620078004600690065006c0064004e0061006d00 ,
                        0x65005d003d0022004900440022002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000100ff000000ffffff001c0000004900 ,
                        0x4900660028005b007400620078004600690065006c0064004e0061006d006500 ,
                        0x5d003d0022004900440022002c0031002c003000290000000000000000000000 ,
                        0x0000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Width =1740
                    Height =315
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxFieldName"
                    ControlSource ="Column"
                    ConditionalFormat = Begin
                        0x010000006c000000010000000000000002000000000000000500000001000100 ,
                        0xff000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x220049004400220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000200000001000100ff000000ffffff00040000002200 ,
                        0x490044002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1380
                    Height =315
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxAllowZLS"
                    ControlSource ="AllowZLS"
                    GridlineColor =10921638

                    LayoutCachedLeft =1380
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Height =315
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxSize"
                    ControlSource ="Length"
                    GridlineColor =10921638

                    LayoutCachedLeft =720
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1980
                    Height =315
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxColTypeSize"
                    ControlSource ="=IIf([tbxFieldType]=\"Text\",[tbxFieldType] & \"(\" & [tbxSize] & \")\",[tbxFiel"
                        "dType])"
                    ConditionalFormat = Begin
                        0x01000000d0000000020000000100000000000000000000001900000001000000 ,
                        0x0000ff00ffffff0001000000000000001a0000003700000001000100ff000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b0074006200780041006c006c006f0077005a004c005300 ,
                        0x5d003d0031002c0031002c00300029000000000049004900660028005b007400 ,
                        0x620078004600690065006c0064004e0061006d0065005d003d00220049004400 ,
                        0x22002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1980
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x0100020000000100000000000000010000000000ff00ffffff00180000004900 ,
                        0x4900660028005b0074006200780041006c006c006f0077005a004c0053005d00 ,
                        0x3d0031002c0031002c0030002900000000000000000000000000000000000000 ,
                        0x000000010000000000000001000100ff000000ffffff001c0000004900490066 ,
                        0x0028005b007400620078004600690065006c0064004e0061006d0065005d003d ,
                        0x0022004900440022002c0031002c003000290000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1800
                    Width =240
                    Height =315
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =255
                    Name ="tbxReqdIcon"
                    ControlSource ="=IIf([tbxRequired]=1,\"*\",\"\")"
                    Format ="**"
                    ConditionalFormat = Begin
                        0x0100000094000000010000000100000000000000000000001900000001000000 ,
                        0xed1c2400ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b0074006200780052006500710075006900720065006400 ,
                        0x5d003d0031002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1800
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ed1c2400ffffff00180000004900 ,
                        0x4900660028005b00740062007800520065007100750069007200650064005d00 ,
                        0x3d0031002c0031002c0030002900000000000000000000000000000000000000 ,
                        0x000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1620
                    Height =315
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxRequired"
                    ControlSource ="IsReqd"
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3360
                    Width =240
                    Height =315
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxIDField"
                    StatusBarText ="ID fields are autogenerated && cannot be imported. 'None' should be the CSV valu"
                        "e at right."
                    ControlTipText ="ID fields are autogenerated, do not import a CSV field to them!"
                    ConditionalFormat = Begin
                        0x01000000da000000020000000100000000000000000000001d00000001000000 ,
                        0xff000000ffffff0001000000000000001e0000003c00000001000000ffffff00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b007400620078004600690065006c0064004e0061006d00 ,
                        0x65005d003d0022004900440022002c0031002c00300029000000000049004900 ,
                        0x660028005b007400620078004600690065006c0064004e0061006d0065005d00 ,
                        0x3c003e0022004900440022002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3360
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010002000000010000000000000001000000ff000000ffffff001c0000004900 ,
                        0x4900660028005b007400620078004600690065006c0064004e0061006d006500 ,
                        0x5d003d0022004900440022002c0031002c003000290000000000000000000000 ,
                        0x0000000000000000000000010000000000000001000000ffffff00ffffff001d ,
                        0x00000049004900660028005b007400620078004600690065006c0064004e0061 ,
                        0x006d0065005d003c003e0022004900440022002c0031002c0030002900000000 ,
                        0x000000000000000000000000000000000000
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
' Form:         TableFieldList
' Level:        Application form
' Version:      1.00
' Basis:        Dropdown form
'
' Description:  List form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, October 6, 2016
' References:   -
' Revisions:    BLC - 10/6/2016 - 1.00 - initial version
'               BLC - 10/20/2016 - 1.01 - removed button caption, selectedID, selectedvalue properties,
'                                         button events
'               BLC - 12/13/2016 - 1.02 - added highlighting for current field
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String

Private m_Table As String
Private m_Fields As String
Private m_TableColumns As String

'listbox scrolling handles
Dim hWnd_A As Long
Dim hWnd_B As Long

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)

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

Public Property Let Table(Value As String)
        m_Table = Value

        'populate form
        PopulateForm
End Property

Public Property Get Table() As String
    Table = m_Table
End Property

Public Property Let TableColumns(Value As String)
        m_TableColumns = Value
End Property

Public Property Get TableColumns() As String
    TableColumns = m_TableColumns
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
' Source/date:  Bonnie Campbell, October 6, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/6/2016 - initial version
'   BLC - 10/20/2016 - code cleanup
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
        
    'defaults
    lblLinkedTable.Caption = StringFromCodepoint(uLinked)
    lblLinkedTable.ForeColor = lngYellow
    lblLinkedTable.Visible = False
    tbxIDField = StringFromCodepoint(uProhibited)
    tbxIDField.ControlTipText = "ID fields are autogenerated, do not import a CSV field to them!"
    tbxIDField.StatusBarText = "ID fields are autogenerated & cannot be imported. 'None' should be the CSV value at right."
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[TableFieldList form])"
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
' Source/date:  Bonnie Campbell, October 6, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/6/2016 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Load[TableFieldList form])"
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
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
       
  'PopulateForm
'  MsgBox Me.CurrentRecord
   'Me.tbxCSVPseudoRecord.Value = Forms("ImportMap").Controls("tbxCSVRecord") '[Forms]![ImportMap].[Controls]![tbxCSVRecord]
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[TableFieldList form])"
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
' Source/date:  Bonnie Campbell, October 6, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/6/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[TableFieldList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          PopulateForm
' Description:  form populating actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 6, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/6/2016 - initial version
'   BLC - 10/20/2016 - code cleanup
' ---------------------------------
Private Sub PopulateForm()
On Error GoTo Err_Handler
    
    'set displayed title
    lblTitle.Caption = Me.Table & " fields"

    'determine if table is linked
    lblLinkedTable.Visible = IsLinked(Me.Table)

    'retrieve field info
    Dim aryFieldInfo() As Variant 'string
    
    aryFieldInfo = FetchDbTableFieldInfo(Me.Table)
    
    'clear table
    ClearTable "usys_temp_rs"
    
    'populate w/ table data
    Dim rs As DAO.Recordset
    Dim aryRecord() As String
    Dim i As Integer
    Dim strTableColumns As String
    
    'default
    strTableColumns = ""
    
    Set rs = CurrentDb.OpenRecordset("usys_temp_rs", dbOpenDynaset)
    
    For i = 0 To UBound(aryFieldInfo)
        
        'create new record
        rs.AddNew
        
        aryRecord = Split(aryFieldInfo(i), "|")
        
        rs!Column = aryRecord(0)
        rs!ColType = aryRecord(5)
        rs!IsReqd = IIf(aryRecord(3) = False, 0, 1)
        rs!Length = aryRecord(2)
        rs!AllowZLS = IIf(aryRecord(4) = False, 0, 1)
    
        'add the new record
        rs.Update
        
        'prepare table columns list
        strTableColumns = strTableColumns & aryRecord(0) & ", "
        
    Next
    
    'prepare table columns list
    Me.TableColumns = Left(Trim(strTableColumns), Len(Trim(strTableColumns)) - 1)
    
Debug.Print Me.TableColumns
    
    Set Me.Recordset = rs 'Forms!TableFieldList.Recordset = rs
    
    Me.Requery
    
    tbxFieldName.ControlSource = "Column" 'rs.Fields("Column").Value
    tbxFieldType.ControlSource = "ColType" '.Fields("ColType").Value
    tbxRequired.ControlSource = "IsReqd"
    tbxSize.ControlSource = "Length"
    tbxAllowZLS.ControlSource = "AllowZLS"
    'tbxColTypeSize.ControlSource =

Exit_Handler:
    'cleanup
    Set rs = Nothing
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateForm[TableFieldList form])"
    End Select
    Resume Exit_Handler
End Sub
