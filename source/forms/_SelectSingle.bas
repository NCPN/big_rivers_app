Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4500
    DatasheetFontHeight =11
    ItemSuffix =5
    Left =4875
    Top =3390
    Right =17295
    Bottom =14385
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x98f234fbd5a9e440
    End
    Caption ="Data Entry"
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
            Height =360
            BackColor =4144959
            Name ="FormHeader"
            AutoHeight =1
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
                    Caption ="Title"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =1560
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =540
                    Width =1740
                    Height =300
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"4\";\"4\""
                    Name ="cbxDropdown"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ID, FirstName + ' ' + LastName AS Name FROM Contact; "
                    ColumnWidths ="1440"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1560
                    LayoutCachedTop =540
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =540
                            Width =1260
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblDropdown"
                            Caption ="Data Enterer"
                            GridlineColor =10921638
                            LayoutCachedLeft =180
                            LayoutCachedTop =540
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3060
                    Top =1080
                    Width =1260
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnEnter"
                    Caption ="Next >"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3060
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1440
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
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Width =4140
                    Height =300
                    BorderColor =8355711
                    ForeColor =5855577
                    Name ="lblDirections"
                    Caption ="directions"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =300
                    ForeTint =65.0
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
' Form:         _SelectSingle
' Level:        Application form
' Version:      1.01
' Basis:        Dropdown form
'
' Description:  Select single form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, April 20, 2016
' References:   -
' Revisions:    BLC - 4/20/2016 - 1.00 - initial version
'               BLC - 6/28/2016 - 1.01 - revised s_site_list to s_site_list_active
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_DropdownLabel As String
Private m_DropdownDataSource As String
Private m_ButtonCaption
Private m_SelectedID As Integer
Private m_SelectedValue As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidLabel(Value As String)
Public Event InvalidDataSource(Value As String)
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

Public Property Let DropdownLabel(Value As String)
    If Len(Value) > 0 Then
        m_DropdownLabel = Value

        'set the form dropdown
        Me.lblDropdown.Caption = m_DropdownLabel
    Else
        RaiseEvent InvalidLabel(Value)
    End If
End Property

Public Property Get DropdownLabel() As String
    DropdownLabel = m_DropdownLabel
End Property

Public Property Let DropdownDataSource(Value As String)
    If Len(Value) > 0 Then
        m_DropdownDataSource = Value

        'set the form dropdown
        Me.cbxDropdown.RowSource = m_DropdownDataSource
    Else
        RaiseEvent InvalidDataSource(Value)
    End If
End Property

Public Property Get DropdownDataSource() As String
    DropdownDataSource = m_DropdownDataSource
End Property

Public Property Let ButtonCaption(Value As String)
    If Len(Value) > 0 Then
        m_ButtonCaption = Value

        'set the form button caption
        Me.btnEnter.Caption = m_ButtonCaption
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
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
'   BLC - 6/28/2016 - revised s_site_list to s_site_list_active
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'eliminate NULLs
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

    Dim strSQL As String
    Dim DropDownWidth As Double
    
    'default
    DropDownWidth = 1.2083 * TWIPS_PER_INCH
    
    'determine which form info
    Select Case Me.OpenArgs
        Case "park"
            Me.Title = "Park"
            Me.DropdownLabel = "Park"
            Me.Directions = "Select the desired park."
            Me.ButtonCaption = "Next >"
            
            strSQL = GetTemplate("s_park_list")
        
        Case "river"
            Me.Title = "River"
            Me.DropdownLabel = "River"
            Me.Directions = "Select the desired river segment."
            Me.ButtonCaption = "Next >"
                    
            strSQL = Replace(GetTemplate("s_river_list"), "[ParkCode]", TempVars("ParkCode"))
            
        Case "site"
            Me.Title = "Site"
            Me.DropdownLabel = "Site"
            Me.Directions = "Select the desired site."
            Me.ButtonCaption = "Next >"
                        
            strSQL = Replace(GetTemplate("s_site_list_active"), "[ParkCode]", TempVars("ParkCode"))
                        
            DropDownWidth = 2 * TWIPS_PER_INCH
        
        Case "feature"
            Me.Title = "Feature"
            Me.DropdownLabel = "Feature"
            Me.Directions = "Select the desired feature."
            Me.ButtonCaption = "Next >"
                    
            strSQL = Replace(GetTemplate("s_feature_list"), "[ParkCode]", TempVars("ParkCode"))
        
        Case "data_entry"
            Me.Title = "Data Entry"
            Me.DropdownLabel = "User"
            Me.Directions = "Select the current user."
            Me.ButtonCaption = "Next >"
            
            strSQL = GetTemplate("s_contact_list")

    End Select
        
        Debug.Print strSQL
        
        'fetch data
        Me.DropdownDataSource = strSQL
        
        With Me.cbxDropdown
            .Requery
            .BoundColumn = 1
            .ColumnCount = 2
            .ColumnWidths = "0;1.6"
            .Width = DropDownWidth
        End With

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[_SelectSingle form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxDropdown_Change
' Description:  Dropdown change actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub cbxDropdown_Change()
On Error GoTo Err_Handler

    Me.SelectedID = CInt(cbxDropdown.Column(0))
    Me.SelectedValue = CStr(cbxDropdown.Column(1))
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxDropdown_Change[_SelectSingle form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnEnter_Click
' Description:  Enter button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub btnEnter_Click()
On Error GoTo Err_Handler
    
    Dim iClearBelow As Integer
    
    'default (clear no values)
    iClearBelow = 4
    
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler
    
    Select Case Me.OpenArgs
        Case "park"
            'store selected ID
            TempVars.Add "park", Me.SelectedID
            TempVars.Add "ParkCode", Me.SelectedValue
            iClearBelow = 0
        Case "river"
            'store selected ID
            TempVars.Add "segment", Me.SelectedID
            TempVars.Add "river", Me.SelectedValue
            iClearBelow = 1
        Case "site"
            'store selected ID
            TempVars.Add "site", Me.SelectedID
            TempVars.Add "SiteCode", Mid(Me.SelectedValue, Len(Me.SelectedValue) - 2, 2)
            iClearBelow = 2
        Case "feature"
            'store selected ID
            TempVars.Add "feature", Me.SelectedID
            TempVars.Add "Feature", Me.SelectedValue
            iClearBelow = 3
        Case "data_entry"
    End Select
    
    'update calling form
    Call Forms("Main").UpdateBreadcrumb(iClearBelow)
    
    DoCmd.Close acForm, "_SelectSingle"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEnter_Click[_SelectSingle form])"
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
' Source/date:  Bonnie Campbell, April 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/20/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[_SelectSingle form])"
    End Select
    Resume Exit_Handler
End Sub
