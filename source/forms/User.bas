Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7860
    DatasheetFontHeight =11
    ItemSuffix =36
    Left =4455
    Top =3150
    Right =21885
    Bottom =14160
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x741eede0d9c6e440
    End
    RecordSource ="Contact"
    Caption ="User"
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =1380
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =60
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="User"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =420
                    Width =6840
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Please confirm the user entering/viewing data."
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =1065
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblUser"
                    Caption ="User"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =1065
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =1380
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =720
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =6660
                    Top =60
                    Width =720
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnNext"
                    Caption ="Next"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Continue"
                    GridlineColor =10921638

                    LayoutCachedLeft =6660
                    LayoutCachedTop =60
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =420
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
                    Overlaps =1
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =75
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =75
                    LayoutCachedWidth =960
                    LayoutCachedHeight =375
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7560
                    Top =105
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxID"
                    GridlineColor =10921638

                    LayoutCachedLeft =7560
                    LayoutCachedTop =105
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =405
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1020
                    Top =60
                    Width =3420
                    Height =315
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxUser"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT c.ID, LastName +',  '+ FirstName + ' ('+ UserName +')' AS AppUser, Access"
                        "Level FROM (Contact AS c INNER JOIN Contact_Access AS ca ON ca.Contact_ID = c.ID"
                        ") INNER JOIN Access AS a ON a.ID = ca.Access_ID WHERE UserName <> \"\" AND IsAct"
                        "ive = 1 ORDER BY LastName, FirstName, Username; "
                    ColumnWidths ="0;2160;0"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
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
' Form:         User
' Level:        Application form
' Version:      1.00
' Basis:        Dropdown form
'
' Description:  List form object related properties, User, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, June 155, 2016
' References:   -
' Revisions:    BLC - 6/15/2016 - 1.00 - initial version
'               BLC - 6/30/2016 - 1.01 - added cbxUser GotFocus() & KeyDown() actions
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_ButtonCaption
Private m_SelectedID As Integer
Private m_SelectedValue As String

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

Public Property Let ButtonCaption(Value As String)
    If Len(Value) > 0 Then
        m_ButtonCaption = Value

        'set the form button caption
        Me.btnNext.Caption = m_ButtonCaption
    Else
        RaiseEvent InvalidCaption(Value)
    End If
End Property

Public Property Get ButtonCaption() As String
    ButtonCaption = m_ButtonCaption
End Property

Public Property Let SelectedID(Value As Integer)
        m_SelectedID = Value
End Property

Public Property Get SelectedID() As Integer
    SelectedID = m_SelectedID
End Property

Public Property Let SelectedValue(Value As String)
        m_SelectedValue = Value
End Property

Public Property Get SelectedValue() As String
    SelectedValue = m_SelectedValue
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
' Source/date:  Bonnie Campbell, June 155, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/15/2016 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Title = "User"
    Directions = "Please confirm the user entering/viewing data."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
    
    'set hover
    btnNext.HoverColor = lngGreen
      
    'defaults
    tbxIcon.ForeColor = lngRed
    btnNext.Enabled = False
    cbxUser.BackColor = lngYellow
  
    'set list of users
    Me.cbxUser.RowSource = GetTemplate("s_app_user") '"SELECT ID, LastName +','+ FirstName + '('+UserName +')' AS User FROM Contact;"
  
    'ID default -> value used only for edits of existing table values
    tbxID.Value = 0
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[User form])"
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
' Source/date:  Bonnie Campbell, June 155, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/15/2016 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'eliminate NULLs
    'If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[User form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxUser_GotFocus
' Description:  Drops down the combox on focus
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:
'   missinglinq & ADezii, July 29, 2010
'   https://bytes.com/topic/access/answers/892371-how-do-you-allow-arrow-keys-combo-box
' Adapted:      Bonnie Campbell, June 30, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/30/2016 - initial version
' ---------------------------------
Private Sub cbxUser_GotFocus()
On Error GoTo Err_Handler
    
    cbxUser.Dropdown

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxUser_GotFocus[User form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxUser_KeyDown
' Description:  Combobox keystroke actions, handles up & down arrow keys
' Assumptions:  -
' Note:         .ItemData and .Value are used as setting .ListIndex directly triggers
'               the AfterUpdate() event causing cbxUser to lose focus & results in
'               error #7777 You've used the ListIndex property incorrectly.
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:
'   missinglinq & ADezii, July 29, 2010
'   https://bytes.com/topic/access/answers/892371-how-do-you-allow-arrow-keys-combo-box
'   Dirk Goldgar, June 6, 2014
'   https://social.msdn.microsoft.com/Forums/office/en-US/3932cc02-9d3f-4430-97bf-c8c95999d870/problem-using-listindex-in-access-2010?forum=accessdev
' Adapted:      Bonnie Campbell, June 30, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/30/2016 - initial version
' ---------------------------------
Private Sub cbxUser_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler
    
    With cbxUser
'        Select Case KeyCode
'            Case vbKeyDown
'              If .ListIndex <> .ListCount - 1 Then
'                .ListIndex = .ListIndex + 1
'              Else
'                .ListIndex = 0
'              End If
'           Case vbKeyUp
'              If .ListIndex <> 0 Then
'                .ListIndex = .ListIndex - 1
'              Else
'                .ListIndex = .ListCount - 1
'              End If
'        End Select
    
        Select Case KeyCode
            Case vbKeyDown
              If .ListIndex <> .ListCount - 1 Then
                .Value = .ItemData(.ListIndex + 1)
              Else
                .Value = .ItemData(0)
              End If
           Case vbKeyUp
              If .ListIndex <> 0 Then
                .Value = .ItemData(.ListIndex - 1)
              Else
                .Value = .ItemData(.ListCount - 1)
              End If
        End Select
    
    End With

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxUser_KeyDown[User form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxUser_AfterUpdate
' Description:  Combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 29, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/29/2016 - initial version
' ---------------------------------
Private Sub cbxUser_AfterUpdate()
On Error GoTo Err_Handler
    
    'set global values
    TempVars.Add "AppUserID", CInt(cbxUser.Column(0))
    TempVars.Add "UserAccessLevel", cbxUser.Column(2)

    ReadyToContinue
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxUser_AfterUpdate[User form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnNext_Click
' Description:  Undo button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 15, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub btnNext_Click()
On Error GoTo Err_Handler
    
    DoCmd.Close
    DoCmd.OpenForm "Main", acNormal, , , , , "User|" & TempVars("AppUserID")

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNext_Click[User form])"
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
' Source/date:  Bonnie Campbell, June 155, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/15/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[User form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ReadyToContinue
' Description:  Check if form values are ready to save
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub ReadyToContinue()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    'requires: site code & name (directions & description optional)
    If Len(Nz(cbxUser.Value, "")) > 0 Then
        isOK = True
    End If
    
    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    btnNext.Enabled = isOK
    
    'refresh form
    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyToContinue[Site form])"
    End Select
    Resume Exit_Handler
End Sub
