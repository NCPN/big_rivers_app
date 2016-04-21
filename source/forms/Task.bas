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
    Width =6300
    DatasheetFontHeight =11
    ItemSuffix =25
    Left =3150
    Top =2415
    Right =18945
    Bottom =14175
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
            Height =447
            BackColor =65280
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
                    Name ="lblTitle"
                    Caption ="Comment"
                    OnClick ="[Event Procedure]"
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
                    BorderColor =65280
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
                Begin Label
                    OverlapFlags =85
                    Left =4200
                    Top =60
                    Width =1980
                    Height =300
                    ForeColor =8355711
                    Name ="lblContext"
                    Caption ="Context"
                    GridlineColor =10921638
                    LayoutCachedLeft =4200
                    LayoutCachedTop =60
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =360
                    BorderTint =100.0
                End
            End
        End
        Begin Section
            Height =3060
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =60
                    Width =5700
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblInstructions"
                    Caption ="Instructions"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =300
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4980
                    Top =2700
                    Width =1200
                    Height =240
                    ForeColor =4210752
                    Name ="btnCancel"
                    Caption ="Cancel"
                    GridlineColor =10921638

                    LayoutCachedLeft =4980
                    LayoutCachedTop =2700
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =2940
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =3660
                    Top =2700
                    Width =1200
                    Height =240
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnAdd"
                    Caption ="Add"
                    GridlineColor =10921638

                    LayoutCachedLeft =3660
                    LayoutCachedTop =2700
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =2940
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
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =720
                    Width =5700
                    Height =720
                    TabIndex =2
                    BackColor =15921906
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTask"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =180
                    LayoutCachedTop =720
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1440
                    BackShade =95.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =2700
                    Top =420
                    Width =1380
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCharacterCount"
                    Caption ="Character Count"
                    GridlineColor =10921638
                    LayoutCachedLeft =2700
                    LayoutCachedTop =420
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =660
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =4800
                    Top =420
                    Width =1380
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblMaxCount"
                    Caption ="/ XX characters"
                    GridlineColor =10921638
                    LayoutCachedLeft =4800
                    LayoutCachedTop =420
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =660
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =4140
                    Top =420
                    Width =660
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCount"
                    Caption ="25"
                    GridlineColor =10921638
                    LayoutCachedLeft =4140
                    LayoutCachedTop =420
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =660
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1080
                    Left =660
                    Top =1740
                    Width =1080
                    Height =360
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"30\""
                    Name ="cbxStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ID, Status, Icon, Sequence FROM Status ORDER BY Sequence; "
                    ColumnWidths ="0;1080"
                    GridlineColor =10921638

                    LayoutCachedLeft =660
                    LayoutCachedTop =1740
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =2100
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =1740
                            Width =984
                            Height =314
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblStatus"
                            Caption ="Status"
                            GridlineColor =10921638
                            LayoutCachedTop =1740
                            LayoutCachedWidth =984
                            LayoutCachedHeight =2054
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1080
                    Left =2940
                    Top =1740
                    Width =1080
                    Height =360
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"30\""
                    Name ="cbxPriority"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ID, Priority, Sequence FROM Priority; "
                    ColumnWidths ="0;1080"
                    GridlineColor =10921638

                    LayoutCachedLeft =2940
                    LayoutCachedTop =1740
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =2100
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2160
                            Top =1740
                            Width =984
                            Height =314
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPriority"
                            Caption ="Priority"
                            GridlineColor =10921638
                            LayoutCachedLeft =2160
                            LayoutCachedTop =1740
                            LayoutCachedWidth =3144
                            LayoutCachedHeight =2054
                        End
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
' Form:         Task
' Level:        Framework form
' Version:      1.00
'
' Description:  Task form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private p_Task As Form_Comment
Private p_oTask As Task

'---------------------
' Events
'---------------------
Public Event Initialize()
Public Event Terminate()

'---------------------
' Properties
'---------------------

'---------------------
' Events
'---------------------

' ---------------------------------
' Sub:          lblTitle_Click
' Description:  Title click event actions
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
Private Sub lblTitle_Click()
On Error GoTo Err_Handler

    MsgBox "Click event...", vbOKOnly

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblTitle_Click[Comment form])"
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
'   BLC - 11/3/2015 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

    MsgBox "Initializing...", vbOKOnly
    
    Set p_Task = New Form_Comment

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[Comment form])"
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
'   BLC - 11/3/2015 - initial version
' ---------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler
    
    MsgBox "Terminating...", vbOKOnly
    
    Set p_Task = Nothing
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[Comment form])"
    End Select
    Resume Exit_Handler
End Sub

Public Sub testme()
    With p_oTask
        .TaskType = "TaskType.Photo"
        .Task = "Testing description"
        .Status = Status.Opened
        .Priority = Priority.High
        .RequestedByID = 3
        .CompletedByID = 1
        .AddTask
    End With
End Sub

' ---------------------------------
' Sub:          AddTask
' Description:  Add new task item
' Assumptions:  -
' Parameters:   context - what the task is about/task type (string)
'               task
'               recordID - ID for the record the task references (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 28, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/3/2015 - initial version
' ---------------------------------
Private Sub AddTask(context As String, recordID As Integer, Description As String, _
                    Status As Integer, Priority As Integer, Requestor As Integer, _
                    Optional completor As Integer)
On Error GoTo Err_Handler

    
    Dim oTask As New Task
    
    With oTask
        .TaskType = context
        .Task = Description
        .Status = Status
        .Priority = Priority
        .RequestedByID = Requestor
        If IsNumeric(completor) Then
            .CompletedByID = completor
        End If
        .AddTask
    
        If IsNumeric(.ID) Then
            MsgBox "New task ID = " & .ID
        End If
    
    End With
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddTask[Task class])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnCancel_Click
' Description:  Cancel task form entry
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 4, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/4/2015 - initial version
' ---------------------------------
'Private Sub btnCancel_Click()
'On Error GoTo Err_Handler
'
'    DoCmd.Close
'
'Exit_Handler:
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - btnCancel_Click[Task form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' ---------------------------------
' Sub:          Form_Load
' Description:  Form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 4, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/4/2015 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

'    Me.Instructions = "Enter your task."
'    Me.CountLabelVisible = False
'    Me.CurrentCount = "Characters Remaining:"
'    Me.lblCharacterCount.Visible = False
'    Me.MaxCount = 50

'    Me.cbxPriority.AddItem "Set Priority", 0
'    Me.cbxStatus.AddItem "Set Status", 0
    
    PopulateCombobox cbxPriority, "priority"
    PopulateCombobox cbxStatus, "status"
    
    'Me.context = Me.OpenArgs

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Task form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxComment_Change
' Description:  tbxComment actions on change event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 4, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/4/2015 - initial version
' ---------------------------------
'Private Sub tbxComment_Change()
'On Error GoTo Err_Handler
'
'    Me.lblMaxCount.Caption = Me.MaxCount - Len(tbxComment.Text) & " remaining"
'
'    If Me.MaxCount - Len(tbxComment.Text) < 10 Then
'        Me.MaxCountFontColor = vbRed
'    Else
'        Me.MaxCountFontColor = vbBlack
'    End If
'
'    If Len(tbxComment.Text) > Me.MaxCount Then
'        Me.lblMaxCount.Caption = -(Me.MaxCount - Len(tbxComment.Text)) & " over"
'    End If
'
'Exit_Handler:
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tbxComment_Change[Task form])"
'    End Select
'    Resume Exit_Handler
'End Sub
