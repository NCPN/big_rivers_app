Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =7560
    DatasheetFontHeight =11
    ItemSuffix =41
    Left =8205
    Top =2505
    Right =16020
    Bottom =7845
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x0680db994fd0e440
    End
    RecordSource ="tsys_Db_Templates"
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
            CanGrow = NotDefault
            Height =1680
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
                    Caption ="Template SQL"
                    GridlineColor =10921638
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =45
                    Width =6600
                    Height =1260
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="directions"
                    GridlineColor =10921638
                    LayoutCachedLeft =45
                    LayoutCachedWidth =6645
                    LayoutCachedHeight =1260
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =6420
                    Top =1320
                    Width =900
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFormat"
                    Caption ="Format"
                    GridlineColor =10921638
                    LayoutCachedLeft =6420
                    LayoutCachedTop =1320
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =1635
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =1320
                    Width =540
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblHdrID"
                    Caption ="1"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =1320
                    LayoutCachedWidth =660
                    LayoutCachedHeight =1635
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1500
                    Top =1320
                    Width =3300
                    Height =315
                    FontSize =9
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTemplateName"
                    Caption ="s_get_parks"
                    GridlineColor =10921638
                    LayoutCachedLeft =1500
                    LayoutCachedTop =1320
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =1635
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =720
                    Top =1320
                    Width =720
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblVersion"
                    Caption ="1"
                    GridlineColor =10921638
                    LayoutCachedLeft =720
                    LayoutCachedTop =1320
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =1635
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =4860
                    Top =1320
                    Width =1440
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblEffectiveDate"
                    Caption ="Effective Date"
                    GridlineColor =10921638
                    LayoutCachedLeft =4860
                    LayoutCachedTop =1320
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =1635
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6720
                    Top =120
                    Width =720
                    ForeColor =4210752
                    Name ="btnRunSQL"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Run template SQL (select statements only)"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2048b080102890ff1030a0700000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x7088e0ff1048ffff102890ff0000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x7088d0807088e0ff2040b0500000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000002040a070000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000004050b0ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x203890700038f0ff001860700000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2040c0ff0038f0ff002890f00000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005068c070 ,
                        0x5070e0ff0040ffff0030d0ff0018503000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005068c0c0 ,
                        0x5078e0ff1048ffff0040f0ff0018608000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005068d0ff ,
                        0x7090ffff1050ffff1040f0ff0028a0f000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006078d0ff ,
                        0x8098ffff3060ffff1050ffff1038c0f000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000007088e0ff ,
                        0x90a8f0ff80a0ffff6080f0ff2040a0e000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000007088c030 ,
                        0x7088e0ff6078d0ff5068d0ff4068d02000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6720
                    LayoutCachedTop =120
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =480
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =3675
            Name ="Detail"
            AutoHeight =255
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =7440
                    Height =3555
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSQL"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =3615
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
' Form:         TemplateSQL
' Level:        Application form
' Version:      1.01
' Basis:        Dropdown form
'
' Description:  List form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
'               BLC - 1/10/2017 - 1.01 - added tbxSQL property documentation
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
'        Me.btnEdit.Caption = m_ButtonCaption
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
' Assumptions:  tbxSQL is an unbound textbox which is enabled, locked,
'               can grow, can shrink, & has vertical scrollbars enabled
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 1/10/2017 - added documentation @ error 2448, tbxSQL properties, ColorizeText()
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'minimize TemplateList
    ToggleForm "TemplateList", -1

    Me.Caption = "Template SQL"
    lblTitle.Caption = ""
    lblDirections.Caption = "Templates are read only to prevent inadvertent edits." _
                            & vbCrLf _
                            & " Run SQL SELECT statements by clicking the button at right." _
                            & vbCrLf _
                            & " Please contact NCPN data managers to make changes."
    lblDirections.ForeColor = lngLtBlue
    
    'retrieve data from OpenArgs
    If Len(OpenArgs) > 0 Then
        Dim aryOA() As String
        
        aryOA = Split(OpenArgs, "|")
        
        Me.lblHdrID.Caption = "#" & aryOA(0)
        Me.lblVersion.Caption = "vers. " & aryOA(1)
        Me.lblTemplateName.Caption = aryOA(2)
        '---------------------------------------------------
        ' NOTE: tbxSQL must be unbound or Error 2448 occurs
        '       "can't assign a value to this object"
        '---------------------------------------------------
        Me.tbxSQL.Value = aryOA(3) 'ColorizeText(ColorizeText(aryOA(3), "SQL", "blue"), "NEGATIVE")
        Me.lblEffectiveDate.Caption = aryOA(4)
        Me.lblFormat.Caption = aryOA(5)
    Else
        GoTo Exit_Handler
    End If
    
    'set hover
    btnRunSQL.HoverColor = lngGreen
    btnRunSQL.Enabled = False

    'only enable if it's a SELECT query
    If Left(aryOA(2), 1) = "s" Then btnRunSQL.Enabled = True
    
    'don't select SQL
    

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[TemplateSQL form])"
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
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Load[TemplateSQL form])"
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
    
    'open args
    If Len(Me.lblTemplateName.Caption) > 0 Then
        Dim aryOA() As String
        
        aryOA = Split(Me.OpenArgs, "|")
        
    Else
        GoTo Exit_Handler
    End If
       
    'only enable if it's a SELECT query
    If Left(aryOA(2), 1) = "s" Then btnRunSQL.Enabled = True
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[TemplateSQL form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnEdit_Click
' Description:  Enter button click actions
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
Private Sub btnEdit_Click()
On Error GoTo Err_Handler
    
    'populate the parent form
    PopulateForm Me.Parent, ID

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEdit_Click[TemplateSQL form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnRunSQL_Click
' Description:  Run SQL button click actions
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
Private Sub btnRunSQL_Click()
On Error GoTo Err_Handler
    
    'present message if SQL starts w/ params
    If Left(Me.Template, 6) = "PARAMS" Then
        
        'show deleted record message & clear
        DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
                    "msg" & PARAM_SEPARATOR & "SQL Templates may require entry " _
                        & "of parameters before results are presented. " & PARAM_SEPARATOR & _
                    "|Type" & PARAM_SEPARATOR & "info"
    End If
    
    'edit, update, delete
    'selects
    'only run if it's a SELECT query
    If Left(Me.lblTemplateName.Caption, 1) = "s" Then
        Dim db As DAO.Database
        Dim qdf As DAO.QueryDef
        Dim rs As DAO.Recordset
        
        Set db = CurrentDb
        
        With db
            Set qdf = .QueryDefs("usys_temp_qdf")
            
            With qdf
                .sql = Me.tbxSQL
                
                'don't .OpenRecordset here --> causes missing param errors
            End With
            
            'open & run query to provide parameter prompts
            DoCmd.OpenQuery "usys_temp_qdf", acViewNormal
            
            'minimize TemplateSQL
            'ToggleForm "TemplateSQL", -1
            
            'close form
            DoCmd.Close acForm, "TemplateSQL"
            
        End With
    
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Private Sub btnRunSQL_Click[TemplateSQL form])"
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
' References:
'   Fionnuala, October 21, 2010
'   http://stackoverflow.com/questions/3992232/access-how-to-detect-with-vba-whether-a-query-is-opened
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore list if query isn't open, otherwise just close
    If Not SysCmd(acSysCmdGetObjectState, acQuery, "usys_temp_qdf") = acObjStateOpen Then
        'restore TemplateList
        ToggleForm "TemplateList", 0
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[TemplateSQL form])"
    End Select
    Resume Exit_Handler
End Sub
