Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowUpdating =2
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12301
    DatasheetFontHeight =10
    ItemSuffix =29
    Left =2295
    Top =360
    Right =14595
    Bottom =9330
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa0a341e20e76e340
    End
    RecordSource ="qfrm_Task_List"
    Caption =" Task List Browser - Tasks associated with sample locations"
    OnOpen ="[Event Procedure]"
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
    ShowPageMargins =0
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
        Begin Line
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
            Height =1428
            BackColor =11651021
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =1200
                    Width =672
                    Height =228
                    Name ="labPark_code"
                    Caption ="Park*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =900
                    Top =1200
                    Width =1620
                    Height =228
                    FontWeight =700
                    Name ="labLoc_code"
                    Caption ="Sample location*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =9660
                    Top =1200
                    Width =1200
                    Height =227
                    Name ="labRequest_date"
                    Caption ="Request date*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4800
                    Top =1200
                    Width =1140
                    Height =228
                    Name ="labTask_desc"
                    Caption ="Description*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =10864
                    Top =1200
                    Width =1437
                    Height =228
                    Name ="labDate_completed"
                    Caption ="Date completed*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =2520
                    Top =1200
                    Width =1332
                    Height =228
                    Name ="labTask_status"
                    Caption ="Task status*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11340
                    Top =120
                    Width =720
                    Height =294
                    FontWeight =700
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =637
                    Width =1020
                    Height =270
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    Name ="cmbParkFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Parks"
                    StatusBarText ="Filter by park code"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by park code"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =637
                            Width =540
                            Height =228
                            FontWeight =700
                            Name ="labParkFilter"
                            Caption ="Park:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =1920
                    Top =600
                    Width =480
                    Height =300
                    TabIndex =2
                    Name ="togFilterByPark"
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
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the park filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =6540
                    Top =637
                    Width =1380
                    Height =270
                    TabIndex =5
                    Name ="cmbStatusFilter"
                    RowSourceType ="Value List"
                    RowSource ="Active;Complete;Inactive"
                    StatusBarText ="Filter by task status"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by task status"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5280
                            Top =637
                            Width =1140
                            Height =228
                            FontWeight =700
                            Name ="labStatusFilter"
                            Caption ="Task status:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =8040
                    Top =600
                    Width =480
                    Height =300
                    TabIndex =6
                    Name ="togFilterByStatus"
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
                    ControlTipText ="Turn the task status filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9000
                    Top =120
                    Width =1980
                    Height =300
                    FontWeight =700
                    TabIndex =9
                    Name ="cmdNewTask"
                    Caption ="New task item"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Add a new task item"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Line
                    OverlapFlags =85
                    Top =1080
                    Width =12300
                    Name ="Line26"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8940
                    Top =1200
                    Width =660
                    Height =227
                    Name ="labRequest_year"
                    Caption ="Year*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =20
                    Left =10200
                    Top =637
                    Width =900
                    Height =270
                    TabIndex =7
                    Name ="cmbYearFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qfrm_Task_List.Request_year FROM qfrm_Task_List WHERE (((qfrm_Task_List.R"
                        "equest_year) Is Not Null)) GROUP BY qfrm_Task_List.Request_year ORDER BY qfrm_Ta"
                        "sk_List.Request_year DESC; "
                    StatusBarText ="Filter by request year"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by request year"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8640
                            Top =630
                            Width =1440
                            Height =240
                            Name ="labYearFilter"
                            Caption ="Year requested:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =11220
                    Top =600
                    Width =480
                    Height =300
                    TabIndex =8
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
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =20
                    ListWidth =2880
                    Left =3540
                    Top =637
                    Width =1020
                    Height =270
                    TabIndex =3
                    Name ="cmbSiteFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Sites.Site_ID, tbl_Sites.Site_code, tbl_Sites.Site_status FROM tbl_Si"
                        "tes WHERE (((tbl_Sites.Park_code)=[Forms]![frm_Task_List]![cmbParkFilter])) ORDE"
                        "R BY tbl_Sites.Site_status, tbl_Sites.Site_code; "
                    ColumnWidths ="0;1440;1440"
                    StatusBarText ="Filter by site"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Filter by site"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2580
                            Top =637
                            Width =840
                            Height =228
                            Name ="labSiteFilter"
                            Caption ="Site:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =4680
                    Top =600
                    Width =480
                    Height =330
                    TabIndex =4
                    Name ="togFilterBySite"
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
                    ControlTipText ="Turn the site filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6360
                    Top =120
                    Width =2100
                    Height =300
                    FontWeight =700
                    TabIndex =10
                    ForeColor =-2147483630
                    Name ="cmdTaskListRpt"
                    Caption ="View report"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Generate the task list report"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =360
            BackColor =11651021
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9664
                    Top =60
                    Width =1200
                    ColumnWidth =1710
                    TabIndex =5
                    ForeColor =0
                    Name ="txtRequest_date"
                    ControlSource ="Request_date"
                    Format ="yyyy mmm dd"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10924
                    Top =60
                    Width =1320
                    ColumnWidth =1710
                    TabIndex =6
                    Name ="txtDate_completed"
                    ControlSource ="Date_completed"
                    Format ="yyyy mmm dd"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4800
                    Top =60
                    Width =4080
                    ColumnWidth =2310
                    TabIndex =3
                    Name ="txtTask_desc"
                    ControlSource ="Task_desc"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =720
                    ColumnWidth =2310
                    Name ="txtPark_code"
                    ControlSource ="Park_code"
                    StatusBarText ="Park code"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =960
                    Top =60
                    ColumnWidth =2310
                    TabIndex =1
                    ForeColor =0
                    Name ="txtLoc_code"
                    ControlSource ="Loc_code"
                    StatusBarText ="Sample location"
                    FontName ="Arial"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =36
                    Width =780
                    Height =300
                    TabIndex =7
                    Name ="cmdCloseup"
                    Caption ="Closeup"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Open this task record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2520
                    Top =60
                    Width =1320
                    TabIndex =2
                    Name ="txtTask_status"
                    ControlSource ="Task_status"
                    StatusBarText ="Task status"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8940
                    Top =60
                    Width =660
                    TabIndex =4
                    Name ="txtRequest_year"
                    ControlSource ="Request_year"
                    StatusBarText ="Year the task request was made"
                    FontName ="Arial"
                    ControlTipText ="Year the task request was made"

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
' FORM NAME:    frm_Task_List
' Description:  Standard form for viewing tasks associated with sample locations
' Data source:  qfrm_Task_List
' Data access:  view only
' Pages:        none
' Functions:    fxnSortRecords, fxnFilterRecords
' References:   fxnSaveFile
' Source/date:  John R. Boetsch, July 26, 2006
' Revisions:    JRB, 5/21/2007 - Updated to replace cmbType with cmbStatusFilter, and
'                   togFilterByType with togFilterByStatus
'               JRB, 5/21/2008 - updated description and title bar
'               JRB, 6/5/2008 - updated header field and filters for parallel code with
'                   Data Gateway
'               JRB, 9/26/2008 - updated Form_Open to include read only mode
'               JRB, 2/12/2008 - renamed cmbSiteFilter to cmbLocFilter, and togFilterBySite to
'                   togFilterByLoc to be consistent with data browser code; updated Form_Open
'                   to permit filtering on open; updated fxnFilterRecords
'               JRB, 6/9/2009 - added cmdTaskListRpt
'               JRB, 12/17/2009 - updated form to capitalize data elements
' =================================

Dim strSortField As String    ' Keeps track of current sort settings
Dim strSortOrder As String
Dim strSortFieldLabel As String

' ---------------------------------
' SUB:          Form_Open
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
'                                TempVars.Item("Park") vs. cPark
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim varReturn As Variant

    ' Close the form if the switchboard is not open
    If SwitchboardIsOpen = False Then
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this form to function properly.", , "Cannot open the form ..."
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

    ' On opening the form, set the initial sort order
    strSortFieldLabel = "labLoc_code"
    varReturn = fxnSortRecords("Loc_code", "Request_date")

    If Me.FilterOn Then
    ' If the form is filtered, set the filter according to the filtered record
        Me.cmbParkFilter = Me.Park_code
        Me.togFilterByPark = True
        Me.cmbSiteFilter = Me.Site_ID
        Me.togFilterBySite = True
        fxnFilterRecords (True)
    Else
    ' Set the default form filter
        Me.cmbParkFilter = TempVars.Item("Park")
        Me.togFilterByPark = True
        Me.cmbStatusFilter = "Active"
        Me.togFilterByStatus = True
        fxnFilterRecords
    End If

    ' Disable new task button for read only application mode
    If TempVars.Item("UserAccessLevel") = "read only" Then _
        Me.cmdNewTask.Enabled = False Else Me.cmdNewTask.Enabled = True

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbParkFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmbParkFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByPark = Not IsNull(Me.cmbParkFilter)
    Me.cmbSiteFilter = Null
    Me.togFilterBySite = False
    fxnFilterRecords
    Me.cmbSiteFilter.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          togFilterByPark_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub togFilterByPark_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbParkFilter) = False Then fxnFilterRecords _
        Else Me.togFilterByPark = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbSiteFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmbSiteFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterBySite = Not IsNull(Me.cmbSiteFilter)
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          togFilterBySite_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub togFilterBySite_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbSiteFilter) = False Then fxnFilterRecords _
        Else Me.togFilterBySite = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbStatusFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmbStatusFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByStatus = Not IsNull(Me.cmbStatusFilter)
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          togFilterByStatus_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub togFilterByStatus_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbStatusFilter) = False Then fxnFilterRecords _
        Else Me.togFilterByStatus = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbYearFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmbYearFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByYear = Not IsNull(Me.cmbYearFilter)
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          togFilterByYear_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub togFilterByYear_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbYearFilter) = False Then fxnFilterRecords _
        Else Me.togFilterByYear = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdTaskListRpt_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
' ---------------------------------
Private Sub cmdTaskListRpt_Click()
    On Error GoTo Err_Handler

    ' Generate the task list report
    Dim strRptName As String
    Dim strMsg As String
    Dim strFilter As String
    Dim strTimeframe As String
    Dim bFilterOn As Boolean
    Dim strCaption As String
    Dim strPark As String
    Dim strSite As String
    Dim strInitFile As String
    Dim strSaveFile As String
    Dim varResponse As VbMsgBoxResult

    strRptName = "rpt_Task_List"

    strFilter = ""
    bFilterOn = False
    strTimeframe = TempVars.Item("Timeframe")

    strMsg = "This will generate the task list report ..." & vbCrLf & vbCrLf & _
        "Would you like to limit task list output to " & vbCrLf & _
        "scheduled sampling locations for " & strTimeframe & "?" & vbCrLf & vbCrLf & _
        "Select NO to output all active task items ..."
    varResponse = MsgBox(strMsg, vbYesNoCancel, "Task list report")

    Select Case varResponse
      Case vbCancel
        GoTo Exit_Procedure
      Case vbYes
        bFilterOn = True
        strFilter = "[Calendar_year]=""" & strTimeframe & """"
        strCaption = strTimeframe
      Case Else
        ' Do not filter by calendar year
        strCaption = ""
    End Select
    
    ' Get user input for the park and/or location to filter on
    strPark = Trim(InputBox("Enter the park code to filter by" & vbCrLf & _
        "(or leave blank to show all):", "Filter by park", Me.cmbParkFilter))
    strSite = Trim(InputBox("Enter the site code" & vbCrLf & _
        "(or leave blank to show all):", "Filter by site code"))
    ' Create the filter string
    If strPark <> "" Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Park_code]=""" & strPark & """"
    End If
    If strSite <> "" Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Site_code]=""" & strSite & """"
    End If

    DoCmd.OpenReport strRptName, acViewPreview, , strFilter, , strCaption
    If MsgBox("Would you like to save this report?", vbYesNo + vbDefaultButton2, _
        "Save report to a file?") = vbYes Then
        If varResponse = vbYes And strTimeframe <> "" Then
            ' Add timeframe to file name
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                strTimeframe & "_" & CStr(Format(Now(), "yyyymmdd")) & ".snp"
        Else
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                CStr(Format(Now(), "yyyymmdd")) & ".snp"
        End If
        ' Open the save file dialog and update to the actual name given by the user
        strSaveFile = SaveFile(strInitFile, "Snapshot Viewer (*.snp)", "*.snp")
        ' Close the report because opening the new file in its own application
        DoCmd.Close acReport, strRptName, acSaveNo
        DoCmd.OutputTo acOutputReport, strRptName, acFormatSNP, strSaveFile, True
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdNewTask_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmdNewTask_Click()
    On Error GoTo Err_Handler

    Set gvarRefForm = Me.Form
    DoCmd.OpenForm "frm_Task_List_Popup", , , , acFormAdd, , Me.cmbParkFilter

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdCloseup_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmdCloseup_Click()
    On Error GoTo Err_Handler

    If IsNull(Me.Location_ID) = False Then
    ' If there is a location ID in the record ...

    ' Set the global reference variables
    Set gvarRefForm = Me.Form
    DoCmd.OpenForm "frm_Task_List_Popup", , , "[Location_ID]=""" & Me.Location_ID & _
        """ AND [Request_date] = #" & Me.Request_date & "# AND [Task_desc] = """ & _
        Me.Task_desc & """", acFormEdit, , Me.Park_code

    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdClose_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' The next several procedures re-sort the records if the user
'   double-clicks on a field label
' ---------------------------------
' SUB:          labPark_code_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub labPark_code_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Park_code")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          labLoc_code_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub labLoc_code_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Loc_code")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          labTask_status_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub labTask_status_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Task_status")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          labTask_desc_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub labTask_desc_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Task_desc")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          labRequest_year_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub labRequest_year_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Request_year")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          labRequest_date_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub labRequest_date_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Request_date")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          labDate_completed_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub labDate_completed_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Date_completed")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' FUNCTION:     fxnSortRecords
' Description:  Sorts the records by the indicated field
' Parameters:   strFieldName
' Returns:      none
' Throws:       none
' References:   strFieldName, strSortOrder, strSortFieldLabel
'               (form-level variables)
' Source/date:  John R. Boetsch, May 5, 2006
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Function fxnSortRecords(ByVal strFieldName As String, _
    Optional ByVal strField2Name As String)
    On Error GoTo Err_Handler

    Dim strOrderBy As String

    ' If already sorting in ascending order by this field, sort descending
    If strFieldName = strSortField And strSortOrder = "" Then
        strSortOrder = " DESC"
    Else: strSortOrder = ""
    End If
    ' Create the order by string and activate the filter
    strOrderBy = strFieldName & strSortOrder
    If strField2Name <> "" Then
        strOrderBy = strField2Name & " DESC, " & strOrderBy
    End If
    strSortField = strFieldName
    Me.Form.OrderBy = strOrderBy
    Me.Form.OrderByOn = True

    ' Change the label format to indicate the sorted field
    Me.Controls.Item(strSortFieldLabel).FontItalic = False
    Me.Controls.Item(strSortFieldLabel).fontBold = False
    strSortFieldLabel = "lab" & strFieldName
    Me.Controls.Item(strSortFieldLabel).FontItalic = True
    Me.Controls.Item(strSortFieldLabel).fontBold = True

Exit_Procedure:
    Exit Function

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
        "Error encountered (#" & Err.Number & " - fxnSortRecords)"
    Resume Exit_Procedure

End Function

' ---------------------------------
' FUNCTION:     fxnFilterRecords
' Description:  Filter the records by the indicated field
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 5, 2006
' Revisions:    JRB, 5/21/2007 - Updated to replace cmbType with cmbStatusFilter, and
'                   togFilterByType with togFilterByStatus
'               JRB, 6/5/2008 - made code more robust and error-proof
'               JRB, 2/12/2009 - updated to add bOpenFilterOn and code for appending 'AND'
'                   statements
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Function fxnFilterRecords(Optional ByVal bOpenFilterOn As Boolean)
    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim bFilterOn As Boolean

    bFilterOn = bOpenFilterOn   ' default is false
    If bOpenFilterOn Then GoTo Reformat_controls

    strFilter = ""

    ' Build the filter string depending on which fields are being filtered on
    If Me.togFilterByPark Then
        bFilterOn = True
        strFilter = "[Park_code] = """ & Me.cmbParkFilter & """"
    End If
    If Me.togFilterBySite Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Site_ID] = """ & Me.cmbSiteFilter & """"
    End If
    If Me.togFilterByStatus Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Task_status] = """ & Me.cmbStatusFilter & """"
    End If
    If Me.togFilterByYear Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Request_year] = """ & Me.cmbYearFilter & """"
    End If

    Me.Filter = strFilter
    Me.FilterOn = bFilterOn

Reformat_controls:
    ' Make the labels bold or not depending on filter settings
    Me.labParkFilter.fontBold = Me.togFilterByPark
    Me.labSiteFilter.fontBold = Me.togFilterBySite
    Me.labStatusFilter.fontBold = Me.togFilterByStatus
    Me.labYearFilter.fontBold = Me.togFilterByYear

Exit_Procedure:
    Exit Function

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
        "Error encountered (#" & Err.Number & " - fxnFilterRecords)"
    Resume Exit_Procedure

End Function
