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
    Width =3480
    DatasheetFontHeight =11
    ItemSuffix =42
    Right =7530
    Bottom =11790
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xeecc3f14b0d0e440
    End
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
            Height =315
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
                    Caption ="Site"
                    GridlineColor =10921638
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3180
                    Width =240
                    Height =315
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =16777215
                    Name ="tbxLinkedIcon"
                    ControlSource ="=IIf([tbxLinked]=1,StringFromCodepoint([uLinked]),\"\")"
                    ConditionalFormat = Begin
                        0x0100000094000000010000000100000000000000000000001900000001000000 ,
                        0xed1c2400ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b0074006200780041006c006c006f0077005a004c005300 ,
                        0x5d003d0031002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3180
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ed1c2400ffffff00180000004900 ,
                        0x4900660028005b0074006200780041006c006c006f0077005a004c0053005d00 ,
                        0x3d0031002c0031002c0030002900000000000000000000000000000000000000 ,
                        0x000000
                    End
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
                    Visible = NotDefault
                    Enabled = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Height =315
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxFieldType"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
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
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
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
                    CanGrow = NotDefault
                    CanShrink = NotDefault
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
                        0x0100000094000000010000000100000000000000000000001900000001000000 ,
                        0x0000ff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b0074006200780041006c006c006f0077005a004c005300 ,
                        0x5d003d0031002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1980
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010000000000ff00ffffff00180000004900 ,
                        0x4900660028005b0074006200780041006c006c006f0077005a004c0053005d00 ,
                        0x3d0031002c0031002c0030002900000000000000000000000000000000000000 ,
                        0x000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
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
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3060
                    Width =240
                    Height =315
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxLinkedTable"
                    ConditionalFormat = Begin
                        0x0100000094000000010000000100000000000000000000001900000001000000 ,
                        0xed1c2400ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b0074006200780041006c006c006f0077005a004c005300 ,
                        0x5d003d0031002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3060
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ed1c2400ffffff00180000004900 ,
                        0x4900660028005b0074006200780041006c006c006f0077005a004c0053005d00 ,
                        0x3d0031002c0031002c0030002900000000000000000000000000000000000000 ,
                        0x000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1620
                    Height =315
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxRequired"
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
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
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
'Private m_Directions As String
'Private m_ButtonCaption
'Private m_SelectedID As Integer
'Private m_SelectedValue As String

Private m_Table As String
Private m_Fields As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
'Public Event InvalidLabel(Value As String)
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

'Public Property Let Directions(Value As String)
'    If Len(Value) > 0 Then
'        m_Directions = Value
'
'        'set the form directions
'        Me.lblDirections.Caption = m_Directions
'    Else
'        RaiseEvent InvalidDirections(Value)
'    End If
'End Property
'
'Public Property Get Directions() As String
'    Directions = m_Directions
'End Property

Public Property Let Table(Value As String)
        m_Table = Value

        'populate form
        PopulateForm
End Property

Public Property Get Table() As String
    Table = m_Table
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
'   mbizup, 5/29/2008
'   https://www.experts-exchange.com/questions/23441990/moving-data-from-array-to-a-table-in-Vba.html
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'close if no table identified
'    If Len(Nz(Me.OpenArgs, "")) = 0 Then GoTo Exit_Handler
    
    lblTitle.Caption = Me.Title 'Me.OpenArgs
'    lblDirections.Caption = "" '"Edit or Delete Records using the buttons for the record at right." _
                            '& vbCrLf & "Icon codes at left identify if record may be edited/deleted."
'    tbxIcon.Value = StringFromCodepoint(uLocked)
'    tbxIcon.ForeColor = lngDkGreen
'    lblDirections.ForeColor = lngLtBlue
    
    'set hover
'    btnEdit.HoverColor = lngGreen
'    btnDelete.HoverColor = lngGreen

'    btnDelete.Caption = StringFromCodepoint(uDelete)
'    btnDelete.ForeColor = lngRed
    
'    Me.Table = Me.OpenArgs
'
'    'retrieve field info
'    Dim aryFieldInfo() As Variant 'string
'
'    aryFieldInfo = FetchDbTableFieldInfo(Me.Table)
'
'    'save to temp table
'
'
'    'populate fields
''    Dim rs As ADODB.Recordset
''    Dim cols As Integer
''    Dim aryFieldData() As Variant
''
''    Set rs = New ADODB.Recordset 'CreateObject("ADODB.Recordset")
''    rs.Open
''
''    For i = 0 To UBound(aryFieldInfo)
''        cols = CountInString(aryFieldInfo(i), "|") + 1
''        aryFieldData(i) = Split(aryFieldInfo(i))
''
''    Next
'
'    'clear table
'    ClearTable "usys_temp_rs"
'
'    'populate w/ table data
'    Dim rs As DAO.Recordset
'    Dim aryRecord() As String
'    Dim i As Integer
'
'    Set rs = CurrentDb.OpenRecordset("usys_temp_rs", dbOpenDynaset)
'
'    For i = 0 To UBound(aryFieldInfo)
'
'        'create new record
'        rs.AddNew
'
'        aryRecord = Split(aryFieldInfo(i), "|")
'
'        rs!Column = aryRecord(0)
'        rs!ColType = aryRecord(5)
'        rs!IsReqd = IIf(aryRecord(3) = False, 0, 1)
'        rs!Length = aryRecord(2)
'        rs!AllowZLS = IIf(aryRecord(4) = False, 0, 1)
'
'        'add the new record
'        rs.Update
'
'    Next
'
'    Set Forms!TableFieldList.Recordset = rs
'
'    'Me.Requery
'
'    tbxFieldName.ControlSource = "Column" 'rs.Fields("Column").Value
'    tbxFieldType.ControlSource = "ColType" '.Fields("ColType").Value
'    tbxRequired.ControlSource = "IsReqd"
'    tbxSize.ControlSource = "Length"
'    tbxAllowZLS.ControlSource = "AllowZLS"
'    'tbxColTypeSize.ControlSource =
'
''    Me.RecordSource = "" '
'
''    Me.tbxFieldName

Exit_Handler:
    'cleanup
'    Set rs = Nothing
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
'    PopulateForm Me.Parent, ID

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEdit_Click[TableFieldList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnDelete_Click
' Description:  Delete button click actions
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
Private Sub btnDelete_Click()
On Error GoTo Err_Handler
    
    Dim result As Integer
    
    'identify the record ID
'     result = MsgBox("Delete Record this record: #" & tbxID & " ?" _
'                        & vbCrLf & "This action cannot be undone.", vbYesNo, "Delete Record?")

'    If result = vbYes Then DeleteRecord "Event", ID
    
    'clear the deleted record
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[TableFieldList form])"
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
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
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
'   mbizup, 5/29/2008
'   https://www.experts-exchange.com/questions/23441990/moving-data-from-array-to-a-table-in-Vba.html
' Source/date:  Bonnie Campbell, October 6, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/6/2016 - initial version
' ---------------------------------
Private Sub PopulateForm()
On Error GoTo Err_Handler
    
 '   Me.Table = Me.OpenArgs
    'set displayed title
    lblTitle.Caption = Me.Table & " fields"

    'retrieve field info
    Dim aryFieldInfo() As Variant 'string
    
    aryFieldInfo = FetchDbTableFieldInfo(Me.Table)
    
    'clear table
    ClearTable "usys_temp_rs"
    
    'populate w/ table data
    Dim rs As DAO.Recordset
    Dim aryRecord() As String
    Dim i As Integer
    
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
        
    Next
    
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
