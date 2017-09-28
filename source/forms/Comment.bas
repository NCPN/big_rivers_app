Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6300
    DatasheetFontHeight =11
    ItemSuffix =18
    Left =14085
    Top =3510
    Right =26490
    Bottom =14895
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
                    Width =6300
                    BorderColor =65280
                    Name ="lineIndicator"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedTop =432
                    LayoutCachedWidth =6300
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
                    Caption ="comment"
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
                    OnClick ="[Event Procedure]"
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
                    OnClick ="[Event Procedure]"
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
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =180
                    Top =720
                    Width =5700
                    Height =1860
                    TabIndex =2
                    BackColor =15921906
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxComment"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =180
                    LayoutCachedTop =720
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =2580
                    BackShade =95.0
                End
                Begin Label
                    Visible = NotDefault
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
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =4140
                    Top =420
                    Width =660
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =255
                    Name ="lblCount"
                    GridlineColor =10921638
                    LayoutCachedLeft =4140
                    LayoutCachedTop =420
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =-1
                End
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =4740
                    Top =360
                    Width =1500
                    Height =360
                    BorderColor =10921638
                    Name ="rctAlert"
                    GridlineColor =10921638
                    LayoutCachedLeft =4740
                    LayoutCachedTop =360
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =720
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4800
                    Top =420
                    Width =1380
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =255
                    Name ="lblMaxCount"
                    Caption ="-1 remaining"
                    GridlineColor =10921638
                    LayoutCachedLeft =4800
                    LayoutCachedTop =420
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =-1
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
' Form:         AppComment
' Level:        Framework form
' Version:      1.04
'
' Description:  Comment form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
'               BLC - 8/9/2016  - 1.01 - revised Comment to AppComment (comment reserved word)
'               BLC - 12/5/2016 - 1.02 - added instruction & max count
'               BLC - 9/25/2017 - 1.03 - revise for NCPN_framework.XX classes
'               BLC - 9/27/2017 - 1.04 - update to use Factory.NewClassXX() vs GetClass()
' =================================

'---------------------
' Declarations
'---------------------
Private m_oComment As AppComment

Private m_Title As String
Private m_Context As String
Private m_Instructions As String
Private m_CountLabel As String
Private m_CurrentCount As String
Private m_MaxCount As String
Private m_AlertCount As Integer
Private m_RemainingCount As String
Private m_Comment As String

Private m_CommentHeaderColor As Long
Private m_TitleFontColor As Long
Private m_InstructionFontColor As Long
Private m_CountLabelFontColor As Long
Private m_CurrentCountFontColor As Long
Private m_MaxCountFontColor As Long
Private m_RemainingCountFontColor As Long
Private m_AlertBoxBackgroundColor As Long

Private m_CommentVisible As Byte
Private m_ContextVisible As Byte
Private m_InstructionVisible As Byte
Private m_CountLabelVisible As Byte
Private m_CurrentCountVisible As Byte
Private m_MaxCountVisible As Byte
Private m_RemainingCountVisible As Byte
Private m_AlertCountVisible As Byte
Private m_AlertBoxVisible As Byte

Private m_AddButtonText As String
Private m_AddButtonForeColor As Long
Private m_AddButtonColor As Long

Private m_CancelButtonText As String
Private m_CancelButtonForeColor As Long
Private m_CancelButtonColor As Long

Private m_AddButtonVisible As Byte
Private m_CancelButtonVisible As Byte

Private m_AddAction As String
Private m_CancelAction As String
Private m_EditAction As String

'---------------------
' Event Declarations
'---------------------
Public Event Initialize()
Public Event Terminate()

'---------------------
' Properties
'---------------------

' ==== Values ====
Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Form Title"
    If ValidateString(Value, "alphanumdash") Then
        m_Title = Value
    End If
    lblTitle.Caption = m_Title
End Property

Public Property Get Context() As String
    Context = m_Context
End Property

Public Property Let Context(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Context"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_Context = Value
    End If
    lblContext.Caption = m_Context
End Property

Public Property Get Instructions() As String
    Instructions = m_Instructions
End Property

Public Property Let Instructions(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Instructions"
    If ValidateString(Value, "paragraph") Then
        m_Instructions = Value
    End If
    lblInstructions.Caption = m_Instructions
End Property

Public Property Get CountLabel() As String
    CountLabel = m_CountLabel
End Property

Public Property Let CountLabel(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Character Count"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_CountLabel = Value
    End If
    lblCharacterCount.Caption = m_CountLabel
End Property

Public Property Get CurrentCount() As String
    CurrentCount = m_CurrentCount
End Property

Public Property Let CurrentCount(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "1"
    If ValidateString(Value, "numeric") Then
        m_CurrentCount = Value
    End If
    lblCount.Caption = m_CurrentCount
End Property

Public Property Get MaxCount() As String
    MaxCount = m_MaxCount
End Property

Public Property Let MaxCount(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "/ XX characters"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_MaxCount = Value
    End If
    lblMaxCount.Caption = m_MaxCount
End Property

'set the value at which the count display changes color
Public Property Get AlertCount() As Integer
    AlertCount = m_AlertCount
End Property

Public Property Let AlertCount(Value As Integer)
    If Len(Trim(Value)) = 0 Then Value = 10
    m_AlertCount = Value
End Property

Public Property Get RemainingCount() As String
    RemainingCount = m_RemainingCount
End Property

Public Property Let RemainingCount(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "XX characters remain"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_RemainingCount = Value
    End If
    lblMaxCount.Caption = m_RemainingCount
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let Comment(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Comment"
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_Comment = Value
    End If
    tbxComment.Value = m_Comment
End Property

' ==== Color ====
Public Property Get CommentHeaderColor() As Long
    CommentHeaderColor = m_CommentHeaderColor
End Property

Public Property Let CommentHeaderColor(Value As Long)
    m_CommentHeaderColor = Value
    FormHeader.BackColor = m_CommentHeaderColor
End Property

Public Property Get TitleFontColor() As Long
    TitleFontColor = m_TitleFontColor
End Property

Public Property Let TitleFontColor(Value As Long)
    m_TitleFontColor = Value
    lblTitle.ForeColor = m_TitleFontColor
End Property

Public Property Get InstructionFontColor() As Long
    InstructionFontColor = m_InstructionFontColor
End Property

Public Property Let InstructionFontColor(Value As Long)
    m_InstructionFontColor = Value
    lblInstructions.ForeColor = m_InstructionFontColor
End Property

Public Property Get CountLabelFontColor() As Long
    CountLabelFontColor = m_CountLabelFontColor
End Property

Public Property Let CountLabelFontColor(Value As Long)
    m_CountLabelFontColor = Value
    lblCount.ForeColor = m_CountLabelFontColor
End Property

Public Property Get CurrentCountFontColor() As Long
    CurrentCountFontColor = m_CurrentCountFontColor
End Property

Public Property Let CurrentCountFontColor(Value As Long)
    m_CurrentCountFontColor = Value
    lblCount.ForeColor = m_CurrentCountFontColor
End Property

Public Property Get MaxCountFontColor() As Long
    MaxCountFontColor = m_MaxCountFontColor
End Property

Public Property Let MaxCountFontColor(Value As Long)
    m_MaxCountFontColor = Value
    lblMaxCount.ForeColor = m_MaxCountFontColor
End Property

Public Property Get RemainingCountFontColor() As Long
    RemainingCountFontColor = m_RemainingCountFontColor
End Property

Public Property Let RemainingCountFontColor(Value As Long)
    m_RemainingCountFontColor = Value
    lblMaxCount.ForeColor = m_RemainingCountFontColor
End Property

Public Property Get AlertBoxBackgroundColor() As Long
    AlertBoxBackgroundColor = m_AlertBoxBackgroundColor
End Property

Public Property Let AlertBoxBackgroundColor(Value As Long)
    rctAlert.backstyle = 1 '1 = Normal, 0 = Transparent
    m_AlertBoxBackgroundColor = Value
    rctAlert.BackColor = m_AlertBoxBackgroundColor
End Property

' ==== Visibility ====
Public Property Get CommentVisible() As Byte
    CommentVisible = m_CommentVisible
End Property

Public Property Let CommentVisible(Value As Byte)
    m_CommentVisible = Value
    tbxComment.Visible = m_CommentVisible
End Property

Public Property Get InstructionVisible() As Byte
    InstructionVisible = m_InstructionVisible
End Property

Public Property Let InstructionVisible(Value As Byte)
    m_InstructionVisible = Value
    lblInstructions.Visible = m_InstructionVisible
End Property

Public Property Get CountLabelVisible() As Byte
    CountLabelVisible = m_CountLabelVisible
End Property

Public Property Let CountLabelVisible(Value As Byte)
    m_CountLabelVisible = Value
    lblCount.Visible = m_CountLabelVisible
End Property

Public Property Get CurrentCountVisible() As Byte
    CurrentCountVisible = m_CurrentCountVisible
End Property

Public Property Let CurrentCountVisible(Value As Byte)
    m_CurrentCountVisible = Value
    lblCount.Visible = m_CurrentCountVisible
End Property

Public Property Get MaxCountVisible() As Byte
    MaxCountVisible = m_MaxCountVisible
End Property

Public Property Let MaxCountVisible(Value As Byte)
    m_MaxCountVisible = Value
    lblMaxCount.Visible = m_MaxCountVisible
End Property

Public Property Get RemainingCountVisible() As Byte
    RemainingCountVisible = m_RemainingCountVisible
End Property

Public Property Let RemainingCountVisible(Value As Byte)
    m_RemainingCountVisible = Value
End Property

Public Property Get AlertBoxVisible() As Byte
    AlertBoxVisible = m_AlertBoxVisible
End Property

Public Property Let AlertBoxVisible(Value As Byte)
    m_AlertBoxVisible = Value
    Me.rctAlert.Visible = m_AlertBoxVisible
End Property

' ==== Buttons ====
Public Property Get AddButtonText() As String
    AddButtonText = m_AddButtonText
End Property

Public Property Let AddButtonText(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Add"
    If ValidateString(Value, "alphaspace") Then
        m_AddButtonText = Value
    End If
    btnAdd.Caption = m_AddButtonText
End Property

Public Property Get CancelButtonText() As String
    CancelButtonText = m_CancelButtonText
End Property

Public Property Let CancelButtonText(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "Cancel"
    If ValidateString(Value, "alphaspace") Then
        m_CancelButtonText = Value
    End If
    btnCancel.Caption = m_CancelButtonText
End Property

Public Property Get AddButtonForeColor() As Long
    AddButtonForeColor = m_AddButtonForeColor
End Property

Public Property Let AddButtonForeColor(Value As Long)
    m_AddButtonForeColor = Value
    btnAdd.ForeColor = m_AddButtonForeColor
End Property

Public Property Get AddButtonColor() As Long
    AddButtonColor = m_AddButtonColor
End Property

Public Property Let AddButtonColor(Value As Long)
    m_AddButtonColor = Value
    btnAdd.BackColor = m_AddButtonColor
End Property

Public Property Get CancelButtonForeColor() As Long
    CancelButtonForeColor = m_CancelButtonForeColor
End Property

Public Property Let CancelButtonForeColor(Value As Long)
    m_CancelButtonForeColor = Value
    btnCancel.ForeColor = m_CancelButtonForeColor
End Property

Public Property Get CancelButtonColor() As Long
    CancelButtonColor = m_CancelButtonColor
End Property

Public Property Let CancelButtonColor(Value As Long)
    m_CancelButtonColor = Value
    btnCancel.BackColor = m_CancelButtonColor
End Property

Public Property Get AddButtonVisible() As Byte
    AddButtonVisible = m_AddButtonVisible
End Property

Public Property Let AddButtonVisible(Value As Byte)
    m_AddButtonVisible = Value
End Property

Public Property Get CancelButtonVisible() As Byte
    CancelButtonVisible = m_CancelButtonVisible
End Property

Public Property Let CancelButtonVisible(Value As Byte)
    m_CancelButtonVisible = Value
End Property

Public Property Get AddAction() As String
    AddAction = m_AddAction
End Property

Public Property Let AddAction(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "add"
    If ValidateString(Value, "alphanumdashunder") Then
        m_AddAction = Value
    End If
End Property

Public Property Get CancelAction() As String
    CancelAction = m_CancelAction
End Property

Public Property Let CancelAction(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "cancel"
    If ValidateString(Value, "alpha") Then
        m_CancelAction = Value
    End If
End Property
Public Property Get EditAction() As String
    EditAction = m_EditAction
End Property

Public Property Let EditAction(Value As String)
    If Len(Trim(Value)) = 0 Then Value = "edit"
    If ValidateString(Value, "alpha") Then
        m_EditAction = Value
    End If
End Property

'---------------------
' Events
'---------------------

' ---------------------------------
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 4, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/4/2015 - initial version
'   BLC - 12/5/2016 - added instruction and max count inputs
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler
    
    Dim ary() As String
    
    Me.FormHeader.BackColor = lngBrown
    Me.TitleFontColor = lngWhite
    
    Me.lineIndicator.Width = Me.Form.Width
    Me.lineIndicator.borderColor = lngLime
    
    'defaults
    Dim instruction As String
    Dim MaxCount As Integer
    
    instruction = "Enter your establishment comment."
    MaxCount = 50
    
    'set comment context
    ary = Split(Nz(Me.OpenArgs, ""), "|")
    If IsArray(ary) Then
        Me.Context = ary(0) & " - " & ary(1) '"Plot - 24"
        MaxCount = ary(2)
        
        'set instructions based on calling form
        Select Case LCase(ary(0))
            Case "importeddata"
                instruction = "Enter your import comment."
        End Select
    Else
        GoTo Exit_Handler
    End If
    
    Me.Instructions = instruction
    Me.CountLabelVisible = False
    Me.CurrentCount = "Characters Remaining:"
    Me.lblCharacterCount.Visible = False
    Me.MaxCount = MaxCount
    Me.AlertCount = 10
   
    Me.AddAction = "add_"
    
    Me.Context = Me.OpenArgs

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Comment form])"
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
Private Sub tbxComment_Change()
On Error GoTo Err_Handler
    
    Dim CurrentCount As Integer
    
    CurrentCount = CInt(Me.MaxCount) - Len(tbxComment.Text)

    Me.lblMaxCount.Caption = CurrentCount & " remaining"
    
    Me.CurrentCountFontColor = vbBlack
    Me.AlertBoxVisible = False
    Me.MaxCountFontColor = vbBlack
    
    Select Case CurrentCount
        Case Is < Me.AlertCount
            Me.AlertBoxVisible = True
            Me.AlertBoxBackgroundColor = lngYellow
        Case Is = 0
            Me.CurrentCountFontColor = vbRed
        Case Else
    End Select
    
    If CurrentCount < 1 Then 'CInt(Me.MaxCount) Then
        Me.MaxCountFontColor = vbRed
    End If
    
    If Len(tbxComment.Text) > CInt(Me.MaxCount) Then
        Me.lblMaxCount.Caption = -CurrentCount & " over"
        'disable add comment button until count is < or = MaxCount
        Me.btnAdd.Enabled = False
    ElseIf Len(tbxComment.Text) = 0 Then
        'disable add comment button if count = 0
        Me.btnAdd.Enabled = False
    Else
        're-enable add comment button
        Me.btnAdd.Enabled = True
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxComment_Change[Comment form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAdd_Click
' Description:  Add comment form entry
' Assumptions:  Person using the application is the "commentor"
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 12, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/12/2015 - initial version
'   BLC - 8/9/2016   - revised Comment > AppComment (comment reserved word)
'   BLC - 12/6/2016 - revise so comment type = context before "- ID#"
'   BLC - 9/25/2017 - revise for NCPN_framework.XX classes
'   BLC - 9/27/2017 - update to use Factory.NewClassXX() vs GetClass()
' ---------------------------------
Private Sub btnAdd_Click()
On Error GoTo Err_Handler
    
    'Dim oComment As New AppComment
    Dim oComment As NCPN_framework.AppComment
    Set oComment = Factory.NewAppComment
    
    With oComment
        .CommentType = Left(lblContext.Caption, InStr(lblContext.Caption, " - "))
        .TypeID = RemoveChars(lblContext.Caption, True) 'return only numbers
        .Comment = tbxComment.Value
        .CommentorID = TempVars("AppUserID") '3 'Requestor
        '.RequestedByID = 3 'Requestor
        .AddComment
    
        If IsNumeric(.ID) Then
'            MsgBox "New Comment ID = " & .ID

            'show added record message & clear
            DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
                        "Comment added (# " & .ID & " )" & _
                        "|Type" & PARAM_SEPARATOR & "info"
            
            'close comment form
            DoCmd.Close acForm, "Comment"
            
        End If
    
    End With

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAdd_Click[Comment form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnCancel_Click
' Description:  Cancel comment form entry
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
Private Sub btnCancel_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCancel_Click[Comment form])"
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
'   BLC - 11/3/2015 - initial version
' ---------------------------------
Private Sub SetHeaderColor(color As Long)
On Error GoTo Err_Handler
    
    MsgBox "SetHeaderColor...", vbOKOnly
    Me.CommentHeaderColor = color

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetHeaderColor[Comment form])"
    End Select
    Resume Exit_Handler
End Sub
