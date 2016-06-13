Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =8640
    DatasheetFontHeight =9
    ItemSuffix =20
    Left =5085
    Top =2280
    Right =14010
    Bottom =9375
    DatasheetGridlinesColor =12632256
    Filter ="[User_name] = \"Holmgren_Mandy\""
    OrderBy ="Edit_type"
    RecSrcDt = Begin
        0x4d48cd53ef57e340
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT CStr(Year([Edit_date])) AS Calendar_year, tbl_Edit_Log.* FROM tbl_Edit_Lo"
        "g ORDER BY tbl_Edit_Log.Edit_date; "
    Caption =" Edit Log - to document edits to certified project data"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =480
            BackColor =13025979
            Name ="FormHeader"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3390
                    Top =97
                    Width =1170
                    ColumnOrder =2
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cmbTypeFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Edit_Type.Edit_type AS Type FROM tlu_Edit_Type ORDER BY tlu_Edit_Type"
                        ".Sort_order  UNION SELECT tbl_Edit_Log.Edit_type FROM tbl_Edit_Log LEFT JOIN tlu"
                        "_Edit_Type ON tbl_Edit_Log.Edit_type = tlu_Edit_Type.Edit_type WHERE (((tlu_Edit"
                        "_Type.Edit_type) Is Null)) GROUP BY tbl_Edit_Log.Edit_type;"
                    StatusBarText ="Filter by edit type"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by edit type"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2340
                            Top =90
                            Width =930
                            Height =240
                            Name ="labTypeFilter"
                            Caption ="Edit type:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =4680
                    Top =60
                    Width =480
                    Height =300
                    ColumnOrder =3
                    TabIndex =3
                    Name ="togFilterByType"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the type filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListRows =20
                    Left =6030
                    Top =97
                    Width =1860
                    ColumnOrder =4
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cmbUserFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Edit_Log.User_name FROM tbl_Edit_Log GROUP BY tbl_Edit_Log.User_name "
                        "ORDER BY tbl_Edit_Log.User_name; "
                    StatusBarText ="Filter by user who logged the edits"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by user who logged the edits"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5310
                            Top =97
                            Width =600
                            Height =228
                            Name ="labUserFilter"
                            Caption ="User:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =8010
                    Top =60
                    Width =480
                    Height =300
                    ColumnOrder =5
                    TabIndex =5
                    Name ="togFilterByUser"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the user filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =840
                    Top =97
                    Width =900
                    ColumnOrder =0
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="cmbYearFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CStr(Year([Edit_date])) AS Edit_year FROM tbl_Edit_Log GROUP BY CStr(Year"
                        "([Edit_date])) ORDER BY CStr(Year([Edit_date])) DESC; "
                    StatusBarText ="Filter by edit year"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by edit year"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =97
                            Width =600
                            Height =228
                            Name ="labYearFilter"
                            Caption ="Year:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    TabStop = NotDefault
                    OverlapFlags =87
                    Left =1860
                    Top =60
                    Width =480
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="togFilterByYear"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the year filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =6630
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =120
                    Width =2040
                    Height =255
                    ColumnWidth =1875
                    Name ="txtEdit_date"
                    ControlSource ="Edit_date"
                    Format ="mm/dd/yyyy hh:nn:ss"
                    StatusBarText ="Date on which the edits took place"
                    DefaultValue ="Now()"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =120
                            Width =1080
                            Height =255
                            FontWeight =700
                            Name ="labEdit_date"
                            Caption ="Edit date"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5040
                    Left =1200
                    Top =480
                    Width =1560
                    Height =255
                    ColumnWidth =1485
                    TabIndex =2
                    Name ="cmbEdit_type"
                    ControlSource ="Edit_type"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Edit_Type.Edit_type, tlu_Edit_Type.Edit_type_desc FROM tlu_Edit_Type;"
                        " "
                    ColumnWidths ="1008;4032"
                    StatusBarText ="Type of edits made: deletion, update, append, reformat, tbl design"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =480
                            Width =1080
                            Height =255
                            FontWeight =700
                            Name ="labEdit_type"
                            Caption ="Edit type"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =840
                    Width =7320
                    Height =480
                    ColumnWidth =3000
                    TabIndex =3
                    Name ="txtEdit_reason"
                    ControlSource ="Edit_reason"
                    StatusBarText ="Brief description of the reason for edits"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =840
                            Width =1080
                            Height =480
                            FontWeight =700
                            Name ="labEdit_reason"
                            Caption ="Reason for edits"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListRows =20
                    Left =1200
                    Top =1440
                    Width =3780
                    Height =255
                    ColumnWidth =2310
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cmbTable_affected"
                    ControlSource ="Table_affected"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MSysObjects.Name, MSysObjects.Type FROM MSysObjects WHERE (((MSysObjects."
                        "Name) Like \"t*\" And (MSysObjects.Name) Not Like \"tsys*\") AND ((MSysObjects.T"
                        "ype)=1 Or (MSysObjects.Type)=4 Or (MSysObjects.Type)=6)) ORDER BY MSysObjects.Na"
                        "me; "
                    StatusBarText ="Table affected by edits"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =1440
                            Width =1080
                            Height =255
                            Name ="labTable_affected"
                            Caption ="Table affected"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =1800
                    Width =7320
                    Height =720
                    ColumnWidth =2235
                    TabIndex =5
                    Name ="txtFields_affected"
                    ControlSource ="Fields_affected"
                    StatusBarText ="Description of the fields affected"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =1800
                            Width =1080
                            Height =480
                            Name ="labFields_affected"
                            Caption ="Field(s) affected"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =2640
                    Width =7320
                    Height =720
                    ColumnWidth =3000
                    TabIndex =6
                    Name ="txtRecords_affected"
                    ControlSource ="Records_affected"
                    StatusBarText ="Description of the records affected"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =2640
                            Width =1080
                            Height =480
                            Name ="labRecords_affected"
                            Caption ="Record(s) affected"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =3480
                    Width =7320
                    Height =2700
                    ColumnWidth =3000
                    TabIndex =7
                    Name ="txtData_edit_notes"
                    ControlSource ="Data_edit_notes"
                    StatusBarText ="Comments about the data edits"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =3480
                            Width =1080
                            Height =255
                            Name ="labData_edit_notes"
                            Caption ="Notes"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1980
                    Top =6300
                    Width =3180
                    Height =270
                    ColumnWidth =3330
                    TabIndex =8
                    Name ="txtData_edit_ID"
                    ControlSource ="Data_edit_ID"
                    StatusBarText ="Unique identifier for each data edit record"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1320
                            Top =6300
                            Width =600
                            Height =255
                            Name ="labData_edit_ID"
                            Caption ="Edit ID"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6360
                    Top =6300
                    Width =900
                    Height =255
                    ColumnWidth =1200
                    TabIndex =9
                    Name ="txtProject_code"
                    ControlSource ="Project_code"
                    StatusBarText ="Project code, for linking information with other data sets and applications"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5280
                            Top =6300
                            Width =1020
                            Height =255
                            Name ="labProject_code"
                            Caption ="Project code"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3600
                    Left =4620
                    Top =120
                    Width =2760
                    Height =255
                    ColumnWidth =1605
                    TabIndex =1
                    Name ="cmbUser_name"
                    ControlSource ="User_name"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active],'Active','') AS Is_a"
                        "ctive FROM tlu_Project_Crew ORDER BY IIf([Contact_is_active],'Active','') DESC ,"
                        " tlu_Project_Crew.Contact_ID; "
                    ColumnWidths ="2592;1008"
                    StatusBarText ="Name of the person making data edits"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3480
                            Top =120
                            Width =1080
                            Height =255
                            FontWeight =700
                            Name ="labUser_name"
                            Caption ="User name"
                            FontName ="Arial"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
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
' FORM NAME:    frm_Edit_Log
' Description:  Standard form for logging edits to certified event data
' Data source:  In-line SQL statement based on tbl_Edit_Log
' Data access:  add, edit, delete (add only for AppMode = data entry)
' Pages:        none
' Functions:    fxnFilterRecords
' References:   fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, 6/6/2008 - added to Form_Open for more robust handling of open args
'                   and table names
'               JRB, 6/17/2008 - updated fxnFilterRecords to trap errors related to filtering
'                   before a record is validated
'               JRB, 9/4/2008 - updated Form_Open to determine settings based on AppMode;
'                   added Form_Dirty event to avoid edits in read only mode
'               JRB, 10/17/2008 - revised Form_Open to remove a set focus cmd; updated tab
'                   stops; updated form validation code
'               JRB, 5/22/2009 - updated fxnFilterRecords
' =================================

Dim strCurrentUser As String
Dim strProjectCode As String

' ---------------------------------
' SUB:     Form_Open
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - updated to use TempVars.Item("User") vs. cUser,
'                                similarly TempVars.Item("Project") vs. cProject
'               BLC, 8/5/2014 - changed to use setUserAccess for initializing control settings based on app mode
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim strOpenArgs As String
    Dim strAction As String
    Dim strTable As String
    Dim intInStr As Integer

    ' Set the current user, project code, and form settings depending on application mode
    If fxnSwitchboardIsOpen Then
        If Not IsNull(TempVars.item("User")) Then
            strCurrentUser = TempVars.item("User")
        Else
            strCurrentUser = Environ("Username")
        End If
        If Not IsNull(TempVars.item("Project")) Then
            strProjectCode = TempVars.item("Project")
        Else
            strProjectCode = "NONE"
        End If
        
        'initialize controls based on app mode
        setUserAccess Me
        
    Else    ' Switchboard not open
        strCurrentUser = Environ("Username")
        strProjectCode = "NONE"
        ' Disable delete, edit and filter capabilities
        Me.DataEntry = True
        Me.AllowDeletions = False
        Me.cmbYearFilter.Enabled = False
        Me.togFilterByYear.Enabled = False
        Me.cmbTypeFilter.Enabled = False
        Me.togFilterByType.Enabled = False
        Me.cmbUserFilter.Enabled = False
        Me.togFilterByUser.Enabled = False
    End If
    Me.cmbUser_name.DefaultValue = """" & strCurrentUser & """"
    If Me.AllowAdditions Then DoCmd.GoToRecord , , acNewRec
    If Me.OpenArgs <> "" Then
        ' Note: open args sent by referring form should be the action followed by a space
        '   followed by the table name
        strOpenArgs = Me.OpenArgs
        intInStr = InStr(strOpenArgs, " ")
        If intInStr = 0 Then
        ' No space in the open args string ... action only
            strAction = strOpenArgs
        Else
            strAction = Left(strOpenArgs, intInStr - 1)
            strTable = Right(strOpenArgs, Len(strOpenArgs) - intInStr)
            Me.cmbTable_affected = strTable
        End If
        Me.cmbEdit_type = strAction
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_Dirty
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
' ---------------------------------
Private Sub Form_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    If fxnSwitchboardIsOpen Then
        If TempVars.item("UserAccessLevel") = "read only" Then DoCmd.CancelEvent
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     Form_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    If IsNull(Me.txtProject_code) Then Me.Project_code = strProjectCode
    If IsNull(Me.cmbUser_name) Then Me.User_name = strCurrentUser

    ' Validate the record and cancel updates if not valid
    If IsNull(Me.txtEdit_date) Then
        MsgBox "Please indicate the edit date", vbOKOnly, "Validation error"
        Me.txtEdit_date.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.cmbEdit_type) Then
        MsgBox "Please indicate the edit type", vbOKOnly, "Validation error"
        Me.cmbEdit_type.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    ElseIf IsNull(Me.txtEdit_reason) Then
        MsgBox "Please describe the reason for edits", vbOKOnly, "Validation error"
        Me.txtEdit_reason.SetFocus
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     cmbYearFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
' ---------------------------------
Private Sub cmbYearFilter_AfterUpdate()
    On Error GoTo Err_Handler

    If Me.cmbYearFilter = TempVars.item("Timeframe") Then
        Me.FormHeader.backcolor = 11056034
    Else
        Me.FormHeader.backcolor = 13025979
    End If
    Me.togFilterByYear = Not IsNull(Me.cmbYearFilter)
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:     togFilterByYear_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByYear_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbYearFilter) = False Then fxnFilterRecords Else Me.togFilterByYear = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbTypeFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbTypeFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByType = Not IsNull(Me.cmbTypeFilter)
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          togFilterByType_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByType_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbTypeFilter) = False Then fxnFilterRecords Else Me.togFilterByType = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbUserFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub cmbUserFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByUser = Not IsNull(Me.cmbUserFilter)
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          togFilterByUser_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub togFilterByUser_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbUserFilter) = False Then fxnFilterRecords Else Me.togFilterByUser = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' FUNCTION:     fxnFilterRecords
' Description:  Filter the records by the indicated field
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 5, 2006
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, May 2008 - made code more robust and error-proof
'               JRB, 5/22/2009 - updated filter AND clauses
' ---------------------------------
Private Function fxnFilterRecords()
    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim bFilterOn As Boolean

    bFilterOn = False
    strFilter = ""

    If Me.togFilterByYear Then
        bFilterOn = True
        strFilter = "[Calendar_year] = """ & Me.cmbYearFilter & """"
    End If
    If Me.togFilterByType Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Edit_type] = """ & Me.cmbTypeFilter & """"
    End If
    If Me.togFilterByUser Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[User_name] = """ & Me.cmbUserFilter & """"
    End If

    ' Save the record (to trigger validation)
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    ' Apply the filter or move to a new record
    If bFilterOn Then
        Me.Filter = strFilter
    Else
        Me.AllowAdditions = True
        Me.DataEntry = True
    End If
    Me.FilterOn = bFilterOn

    ' Make the labels bold or not depending on filter settings
    Me.labYearFilter.fontBold = Me.togFilterByYear
    Me.labTypeFilter.fontBold = Me.togFilterByType
    Me.labUserFilter.fontBold = Me.togFilterByUser

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 2001   ' Run time canceled event (validation error) - do nothing
        Me.togFilterByYear = False
        Me.togFilterByType = False
        Me.togFilterByUser = False
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnFilterRecords)"
    End Select
    Resume Exit_Procedure

End Function


' ---------------------------------
' SUB:          Form_Close
' Description:  handle form closing tasks
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, June, 2014 for NCPN tools
' Adapted:      -
' Revisions:    BLC - 8/21/2014 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'unhighlight edit log button (toggle off)
    buttonUnHighlight Forms!frm_Switchboard.Controls("btnEditLog"), 1

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[Form_frm_Edit_Log])"
    End Select
    Resume Exit_Procedure
End Sub
