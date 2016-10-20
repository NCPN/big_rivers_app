Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    SubdatasheetExpanded = NotDefault
    DefaultView =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9060
    DatasheetFontHeight =11
    ItemSuffix =44
    Left =4110
    Top =5715
    Right =13095
    Bottom =10965
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x3defb2bfced4e440
    End
    RecordSource ="usys_temp_csv"
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
    Moveable =0
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    SplitFormOrientation =1
    SplitFormDatasheet =1
    SplitFormSize =1350
    SplitFormPrinting =1
    FilterOnLoad =255
    SplitFormOrientation =1
    SplitFormDatasheet =1
    SplitFormSize =1350
    SplitFormPrinting =1
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =255
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
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin ToggleButton
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackColor =-1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-1
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverColor =0
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedColor =0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeColor =0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeColor =0
            PressedForeThemeColorIndex =1
        End
        Begin FormHeader
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
                    OverlapFlags =223
                    Left =2820
                    Top =120
                    Width =5280
                    Height =1320
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="directions"
                    GridlineColor =10921638
                    LayoutCachedLeft =2820
                    LayoutCachedTop =120
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =1440
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =8220
                    Top =600
                    Width =270
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblHdrID"
                    Caption ="ID"
                    GridlineColor =10921638
                    LayoutCachedLeft =8220
                    LayoutCachedTop =600
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =915
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontFamily =2
                    Left =7980
                    Top =780
                    Width =720
                    FontSize =14
                    ForeColor =255
                    Name ="btnImportCSVData"
                    Caption ="import"
                    StatusBarText ="Import CSV data to usys_temp_csv"
                    OnClick ="[Event Procedure]"
                    FontName ="Academy Engraved LET"
                    ControlTipText ="Import CSV data to usys_temp_csv"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000707880ff608090ff607880ff507080ff506070ff405860ff ,
                        0x405060ff404850ff404850ff404040ff303030ff202020ff1018208000000000 ,
                        0x0000000000000000708090ff80a0a0ff70b0d0ff0090d0ff0090d0ff1090d0ff ,
                        0x30a0d0ff50a8d0ff80b8d0ff80b8d0ff70a8c0ff5090b0ff305860ff30384050 ,
                        0x0000000000000000708090ff80c0d0ff80a0b0ff80e0ffff60d0ffff60d0ffff ,
                        0x70d0ffff509860ff308040ff206030ff90b8a0ff80c0e0ff5088a0ff303840c0 ,
                        0xfff8f00000000000708890ff80d0f0ff80a0b0ff80c0d0ff70d8ffff70d8ffff ,
                        0x80d8ffffb0e0ffff308040ff60a870ff206830ff80a890ff70b0e0ff406070ff ,
                        0x2038403000000000708890ff80d8f0ff80c8e0ff80a0b0ff80e0ffff70d8ffff ,
                        0x80d8ffffa0e0ffffd0f0ffff308040ff60a870ff206030ffa0d8f0ff5088a0ff ,
                        0x30587090fff8f000808890ff90e0f0ff90e0ffff90a0b0ff90b8c0ff80d8ffff ,
                        0x80d8ffffb0e8ffffe0f0ffff308040ff80d8a0ff206030ffd0e8f0ff80c8e0ff ,
                        0x707880f0705040608090a0ff90e0f0ffa0e8ffff80c0e0ff90a0b0ff90e0ffff ,
                        0xb0e8ffff308050ff308040ff60a870ff80d8a0ff308040ff206830ff307040ff ,
                        0x90c0e0ff706860d08090a0ffa0e8f0ffa0e8ffffa0e8ffff80a8b0ff90a8b0ff ,
                        0xa0b8c0ffb0c0b0ff308050ff70c080ff80d8a0ff50a060ff408050ffb0c0b0ff ,
                        0xa0a8b0ff8090a0ff8098a0ffa0e8f0ffa0f0ffffa0e8ffffa0e8ffff80d8ffff ,
                        0xc0b0a0fffff8f0ffd0e0d0ff408050ff60a870ff408050ffc0d0c0fffff8f0ff ,
                        0xffe8e0ff705040ff8098a0ffa0f0f0ffb0f0f0ffa0f0ffffa0e8ffffa0e8ffff ,
                        0xc0a8a0ffd0c0b0ffe0d0c0ffc0c8c0ff408050ffc0c8c0ffe0c8c0ffd0b8b0ff ,
                        0xc0b0a0ff604830ff8098a0ffa0d0e0ffb0f0f0ffb0f0f0ffa0f0ffffa0e8ffff ,
                        0xb0a8a0fffffffffffff8ffffd0c0c0fffff8f0fffff0e0ffd0b8b0fffff8f0ff ,
                        0xffe8e0ff604830ff8098a0508098a0ff8098a0ff8098a0ff8098a0ff8098a0ff ,
                        0xb0a8a0ffc0b0a0ffc0b8a0ffc0b0a0ffc0b0a0ffc0b0a0ffc0b0a0ffc0b0a0ff ,
                        0xc0b0a0ff604830ff000000000000000000000000000000000000000000000000 ,
                        0xb0a8a0ffffffffffffffffffc0b0a0fffff8fffffff0f0ffc0b0a0fffff8f0ff ,
                        0xfff0f0ff604830ff000000000000000000000000000000000000000000000000 ,
                        0xb09080ffb08060ffb08060ffb08060ffb07860ffb07860ffb07860ffb07860ff ,
                        0xb08060ffb08060ff000000000000000000000000000000000000000000000000 ,
                        0xb08870ffe0c8b0ffe0c0b0ffb08060ffe0c0b0ffe0c0b0ffb07860ffe0b8b0ff ,
                        0xe0b8b0ffb08060ff000000000000000000000000000000000000000000000000 ,
                        0xb08870ffc09080ffc09070ffb08870ffb08070ffb08060ffb08060ffb08060ff ,
                        0xb08060ffb08060ff
                    End

                    LayoutCachedLeft =7980
                    LayoutCachedTop =780
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =1140
                    ForeThemeColorIndex =-1
                    BackColor =14136213
                    BorderColor =14136213
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7320
                    Top =600
                    Width =720
                    Height =300
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =1
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =7320
                    LayoutCachedTop =600
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =900
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8580
                    Top =600
                    Width =480
                    Height =315
                    ColumnOrder =2
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =8580
                    LayoutCachedTop =600
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =915
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =525
                    Top =240
                    Width =1980
                    Height =504
                    ColumnOrder =0
                    TabIndex =3
                    BorderColor =10921638
                    Name ="optgView"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =525
                    LayoutCachedTop =240
                    LayoutCachedWidth =2505
                    LayoutCachedHeight =744
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =180
                            Top =330
                            Width =600
                            Height =315
                            BackColor =4144959
                            BorderColor =8355711
                            ForeColor =16777215
                            Name ="lblView"
                            Caption ="View"
                            GridlineColor =10921638
                            LayoutCachedLeft =180
                            LayoutCachedTop =330
                            LayoutCachedWidth =780
                            LayoutCachedHeight =645
                            BackThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =885
                            Top =330
                            Height =317
                            OptionValue =1
                            Name ="tglTable"
                            StatusBarText ="View table data"
                            Caption ="Table"
                            FontName ="Calibri"
                            OnKeyPress ="[Event Procedure]"
                            ControlTipText ="View table data"
                            LeftPadding =90
                            TopPadding =90
                            RightPadding =90
                            BottomPadding =120
                            GridlineColor =10921638

                            LayoutCachedLeft =885
                            LayoutCachedTop =330
                            LayoutCachedWidth =1605
                            LayoutCachedHeight =647
                            ForeTint =100.0
                            Bevel =0
                            Gradient =12
                            BackColor =12419407
                            BackTint =100.0
                            OldBorderStyle =1
                            BorderColor =12419407
                            BorderTint =100.0
                            HoverColor =65280
                            HoverThemeColorIndex =-1
                            HoverTint =100.0
                            PressedColor =65280
                            PressedThemeColorIndex =-1
                            PressedShade =100.0
                            HoverForeColor =2366701
                            HoverForeThemeColorIndex =-1
                            HoverForeTint =100.0
                            PressedForeColor =16711680
                            PressedForeThemeColorIndex =-1
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =6
                            WebImagePaddingTop =6
                            WebImagePaddingRight =5
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1665
                            Top =330
                            Height =317
                            TabIndex =1
                            OptionValue =2
                            Name ="tglCSV"
                            Caption ="CSV"
                            FontName ="Calibri"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120
                            GridlineColor =10921638

                            LayoutCachedLeft =1665
                            LayoutCachedTop =330
                            LayoutCachedWidth =2385
                            LayoutCachedHeight =647
                            ForeTint =100.0
                            Bevel =0
                            Gradient =12
                            BackColor =12419407
                            BackTint =100.0
                            OldBorderStyle =1
                            BorderColor =12419407
                            BorderTint =100.0
                            HoverColor =65280
                            HoverThemeColorIndex =-1
                            HoverTint =100.0
                            PressedColor =65280
                            PressedThemeColorIndex =-1
                            PressedShade =100.0
                            HoverForeColor =2366701
                            HoverForeThemeColorIndex =-1
                            HoverForeTint =100.0
                            PressedForeColor =16711680
                            PressedForeThemeColorIndex =-1
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =1140
                    Width =1290
                    Height =285
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblCurrentData"
                    Caption ="Current Data:"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =1140
                    LayoutCachedWidth =1410
                    LayoutCachedHeight =1425
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =4740
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    CanShrink = NotDefault
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =8940
                    Height =4620
                    BorderColor =10921638
                    Name ="DataList"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =4680
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
' Form:         CSVDataList
' Level:        Application form
' Version:      1.00
' Basis:        Dropdown form
'
' Description:  List form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, October 19, 2016
' References:   -
' Revisions:    BLC - 10/19/2016 - 1.00 - initial version
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
' Source/date:  Bonnie Campbell, October 19, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/19/2016 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Dim strTableOrCSV As String
    
    strTableOrCSV = ""
    strTableOrCSV = IIf(Me.optgView.Value = 1, "table", "CSV")

    lblTitle.Caption = ""
'    lblDirections.Caption = "Current " & strTableOrCSV & " data is shown below. " _
'                            & "To change CSV data or re-import it to the CSV table" _
'                            & "click the button at right." _
'                            & vbCrLf & "Choose Table or CSV at right to switch the data being viewed."
    lblDirections.Caption = "To view your selected table data below, choose Table or " _
                            & "choose CSV to view CSV data." _
                            & vbCrLf & "To change CSV data or re-import it to the " _
                            & "temporary CSV table, click the button at right."
                            
    tbxIcon.Value = StringFromCodepoint(uLocked)
    tbxIcon.ForeColor = lngDkGreen
    lblDirections.ForeColor = lngLtBlue
    
    'set hover
    btnImportCSVData.HoverColor = lngGreen

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[CSVDataList form])"
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
' Source/date:  Bonnie Campbell, October 19, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/19/2016 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Load[CSVDataList form])"
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
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[CSVDataList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnImportCSVData_Click
' Description:  Enter button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 19, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/19/2016 - initial version
' ---------------------------------
Private Sub btnImportCSVData_Click()
On Error GoTo Err_Handler
    
    Dim StartFolder As String, strPath As String
    
    'handle upload
    StartFolder = GetSpecialFolderPath("FOLDERID_Recent")
    
    strPath = BrowseFolder("Select CSV file to upload", "Confirm File", _
                        StartFolder, , msoFileDialogFilePicker, "CSV files")
    
    If Len(strPath) > 0 Then
        'upload CSV file
        UploadCSVFile strPath
    
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnImportCSVData_Click[CSVDataList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          optgView_Click
' Description:  option group click actions
' Assumptions:  -
' Note:         Ensure Form AllowEdits = Yes, otherwise option group cannot
'               be changed & toggles will not set the option group value!
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   ParanoidAndroid, March 4, 2009
'   http://www.access-programmers.co.uk/forums/showthread.php?t=167066
' Source/date:  Bonnie Campbell, October 19, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/19/2016 - initial version
' ---------------------------------
Private Sub optgView_Click()
On Error GoTo Err_Handler

    'MsgBox "Hello! CLICK " & Me.optgView.Value

   ' If Me.tglCSV.OptionValue = 2 Then Me.tglCSV.BackColor = lngGreen
    
    'Me.Controls("optgView").Value << error 438 object doesn't support property - EVEN tho immediate does!
    
    Select Case Me.optgView.Value
        Case 0  'Default
        Case 1  'Table
            'Me.tglTable.BackColor = lngLime
            'MsgBox "TABLE!"
            Dim strTable As String
            strTable = Nz(Me.Parent.SelectedTable, "") 'Me.Parent.Form.SelectedTable, "")
            If Len(strTable) > 0 Then _
                Me.DataList.SourceObject = "Table." & strTable
        Case 2  'CSV
            'Me.tglCSV.BackColor = lngLime
            'MsgBox "CSV!"
            Me.DataList.SourceObject = "Table.usys_temp_CSV"
    End Select
    
    Me.Refresh
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - optgView_Click[CSVDataList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          optgView_AfterUpdate
' Description:  option group AfterUpdate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 19, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/19/2016 - initial version
' ---------------------------------
Private Sub optgView_AfterUpdate()
On Error GoTo Err_Handler

    'MsgBox "AU:" & Me.optgView.Value

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - optgView_AfterUpdate[CSVDataList form])"
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
' Source/date:  Bonnie Campbell, October 19, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/19/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[CSVDataList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub tglTable_KeyPress(KeyAscii As Integer)
    MsgBox "Table KeyPress = " & Me.optgView.Value
End Sub
