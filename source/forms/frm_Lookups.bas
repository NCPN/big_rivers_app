Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =2
    GridX =24
    GridY =24
    Width =13215
    DatasheetFontHeight =10
    ItemSuffix =601
    Left =4665
    Top =3315
    Right =12315
    Bottom =14310
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3dabdfc735cde240
    End
    Caption =" Manage Lookup Tables"
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
            FontItalic = NotDefault
            OldBorderStyle =1
            TextAlign =1
            FontWeight =700
            BackColor =8388608
            BorderColor =8388608
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =2
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderWidth =2
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            BorderColor =16776960
        End
        Begin CommandButton
            FontItalic = NotDefault
            FontSize =8
            ForeColor =-2147483630
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =4
            BorderWidth =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =8388608
        End
        Begin CheckBox
            SpecialEffect =4
            BorderWidth =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =8388608
        End
        Begin OptionGroup
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
        End
        Begin BoundObjectFrame
            BorderLineStyle =0
            BackStyle =0
            BorderColor =16776960
        End
        Begin TextBox
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin ListBox
            BorderLineStyle =0
            BackColor =8421376
            ForeColor =16777215
            BorderColor =16776960
            FontName ="Arial"
        End
        Begin ComboBox
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =16776960
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =16776960
        End
        Begin ToggleButton
            FontItalic = NotDefault
            FontSize =8
            ForeColor =-2147483630
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin Tab
            FontItalic = NotDefault
            BackStyle =0
            FontWeight =700
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =10095
            BackColor =13025979
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Top =525
                    Width =13215
                    Height =9570
                    FontSize =9
                    Name ="PageTabs"
                    OnChange ="[Event Procedure]"

                    LayoutCachedTop =525
                    LayoutCachedWidth =13215
                    LayoutCachedHeight =10095
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =120
                            Top =960
                            Width =12960
                            Height =9000
                            Name ="pgSpeciesList"
                            Caption =" Species list"
                            LayoutCachedLeft =120
                            LayoutCachedTop =960
                            LayoutCachedWidth =13080
                            LayoutCachedHeight =9960
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =215
                                    Left =9360
                                    Top =1020
                                    Width =1380
                                    Height =311
                                    FontSize =9
                                    ForeColor =0
                                    Name ="cmdViewDetails"
                                    Caption ="View details"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="View closeup of selected record"

                                    LayoutCachedLeft =9360
                                    LayoutCachedTop =1020
                                    LayoutCachedWidth =10740
                                    LayoutCachedHeight =1331
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =215
                                    Left =10920
                                    Top =1020
                                    Width =1380
                                    Height =311
                                    FontSize =9
                                    TabIndex =1
                                    ForeColor =0
                                    Name ="cmdNewTaxon"
                                    Caption ="New record"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Add a new record"

                                    LayoutCachedLeft =10920
                                    LayoutCachedTop =1020
                                    LayoutCachedWidth =12300
                                    LayoutCachedHeight =1331
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin Subform
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    SpecialEffect =2
                                    Left =120
                                    Top =1440
                                    Width =12960
                                    Height =8520
                                    TabIndex =2
                                    BorderColor =0
                                    Name ="subTaxa"

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =1440
                                    LayoutCachedWidth =13080
                                    LayoutCachedHeight =9960
                                End
                                Begin OptionGroup
                                    SpecialEffect =3
                                    OverlapFlags =215
                                    Left =6120
                                    Top =960
                                    Width =1914
                                    Height =355
                                    TabIndex =3
                                    BackColor =16777215
                                    BorderColor =0
                                    Name ="optgFormMode"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    ControlTipText ="Change the data mode"

                                    LayoutCachedLeft =6120
                                    LayoutCachedTop =960
                                    LayoutCachedWidth =8034
                                    LayoutCachedHeight =1315
                                    Begin
                                        Begin OptionButton
                                            SpecialEffect =2
                                            OverlapFlags =215
                                            BorderWidth =0
                                            Left =7200
                                            Top =1044
                                            OptionValue =1
                                            BorderColor =0
                                            Name ="optEditMode"

                                            LayoutCachedLeft =7200
                                            LayoutCachedTop =1044
                                            LayoutCachedWidth =7460
                                            LayoutCachedHeight =1284
                                            Begin
                                                Begin Label
                                                    FontItalic = NotDefault
                                                    BackStyle =0
                                                    OldBorderStyle =0
                                                    OverlapFlags =247
                                                    TextAlign =0
                                                    Left =7434
                                                    Top =1020
                                                    Width =390
                                                    Height =270
                                                    FontSize =9
                                                    BackColor =16777215
                                                    BorderColor =0
                                                    ForeColor =0
                                                    Name ="labEditMode"
                                                    Caption ="Edit"
                                                    ControlTipText ="View all other records that matched valid fish codes"
                                                    LayoutCachedLeft =7434
                                                    LayoutCachedTop =1020
                                                    LayoutCachedWidth =7824
                                                    LayoutCachedHeight =1290
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            SpecialEffect =2
                                            OverlapFlags =215
                                            BorderWidth =0
                                            Left =6240
                                            Top =1050
                                            OptionValue =0
                                            BorderColor =0
                                            Name ="optViewMode"

                                            LayoutCachedLeft =6240
                                            LayoutCachedTop =1050
                                            LayoutCachedWidth =6500
                                            LayoutCachedHeight =1290
                                            Begin
                                                Begin Label
                                                    FontItalic = NotDefault
                                                    BackStyle =0
                                                    OldBorderStyle =0
                                                    OverlapFlags =247
                                                    TextAlign =2
                                                    Left =6474
                                                    Top =1020
                                                    Width =540
                                                    Height =270
                                                    FontSize =9
                                                    BackColor =16777215
                                                    BorderColor =0
                                                    ForeColor =0
                                                    Name ="labViewMode"
                                                    Caption ="View"
                                                    ControlTipText ="View mode"
                                                    LayoutCachedLeft =6474
                                                    LayoutCachedTop =1020
                                                    LayoutCachedWidth =7014
                                                    LayoutCachedHeight =1290
                                                End
                                            End
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =960
                            Width =12960
                            Height =9000
                            Name ="pgProjectCrew"
                            Caption =" Project crew list"
                            LayoutCachedLeft =120
                            LayoutCachedTop =960
                            LayoutCachedWidth =13080
                            LayoutCachedHeight =9960
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =9420
                                    Top =1020
                                    Width =2040
                                    Height =314
                                    FontSize =9
                                    ForeColor =0
                                    Name ="cmdEditContacts"
                                    Caption ="View / edit contacts"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Add or edit contact records"

                                    LayoutCachedLeft =9420
                                    LayoutCachedTop =1020
                                    LayoutCachedWidth =11460
                                    LayoutCachedHeight =1334
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =8
                                    Left =120
                                    Top =1440
                                    Width =12960
                                    Height =8520
                                    FontSize =9
                                    TabIndex =1
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="lstContacts"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active]=-1,'Yes','No') AS Ac"
                                        "tive, tlu_Project_Crew.Contact_ID AS Name, tlu_Project_Crew.Organization, tlu_Pr"
                                        "oject_Crew.Position_title AS Title, tlu_Project_Crew.Email, Format([Work_voice],"
                                        "\"!@@@-@@@-@@@@\") & IIf(IsNull([Work_ext]),\"\",\" ext. \" & [Work_ext]) AS [Wo"
                                        "rk], tlu_Project_Crew.Contact_updated AS [Last updated] FROM tlu_Project_Crew OR"
                                        "DER BY IIf([Contact_is_active]=-1,'Yes','No') DESC , tlu_Project_Crew.Contact_ID"
                                        "; "
                                    ColumnWidths ="0;576;2592;2880;2592;2304;2016;1440"
                                    OnDblClick ="[Event Procedure]"

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =1440
                                    LayoutCachedWidth =13080
                                    LayoutCachedHeight =9960
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =960
                            Width =12960
                            Height =9000
                            Name ="pgOtherLookups"
                            Caption =" Other lookup tables"
                            LayoutCachedLeft =120
                            LayoutCachedTop =960
                            LayoutCachedWidth =13080
                            LayoutCachedHeight =9960
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =2
                                    Left =120
                                    Top =1500
                                    Width =12960
                                    Height =8457
                                    BorderColor =0
                                    Name ="subLookupTables"

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =1500
                                    LayoutCachedWidth =13080
                                    LayoutCachedHeight =9957
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListRows =20
                                    ListWidth =11520
                                    Left =840
                                    Top =1020
                                    Width =4320
                                    Height =252
                                    FontSize =9
                                    TabIndex =1
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"0\";\"0\""
                                    Name ="cmbTableFilter"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tsys_Link_Tables.Link_table, tsys_Link_Tables.Description_text FROM tsys_"
                                        "Link_Tables WHERE (((tsys_Link_Tables.Link_table) Like \"tlu_*\" And (tsys_Link_"
                                        "Tables.Link_table)<>\"tlu_Project_Crew\" And (tsys_Link_Tables.Link_table)<>\"tl"
                                        "u_Project_Taxa\" And (tsys_Link_Tables.Link_table)<>\"tlu_Park_Taxa\")) ORDER BY"
                                        " tsys_Link_Tables.Link_table; "
                                    ColumnWidths ="4320;7200"
                                    StatusBarText ="Select the lookup table to view"
                                    AfterUpdate ="[Event Procedure]"
                                    OnEnter ="[Event Procedure]"
                                    ControlTipText ="Select the lookup table to view"

                                    LayoutCachedLeft =840
                                    LayoutCachedTop =1020
                                    LayoutCachedWidth =5160
                                    LayoutCachedHeight =1272
                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =180
                                            Top =1020
                                            Width =588
                                            Height =252
                                            FontSize =9
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labTableFilter"
                                            Caption ="Table:"
                                            LayoutCachedLeft =180
                                            LayoutCachedTop =1020
                                            LayoutCachedWidth =768
                                            LayoutCachedHeight =1272
                                        End
                                    End
                                End
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =0
                                    Left =5280
                                    Top =960
                                    Width =7800
                                    Height =480
                                    FontSize =9
                                    FontWeight =400
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labMsg"
                                    Caption ="Note:  Only users with admin or power user access privileges may change lookup d"
                                        "omain values. Please contact the Project Lead if you need to make changes but ar"
                                        "e unable to."
                                    ControlTipText ="View mode"
                                    LayoutCachedLeft =5280
                                    LayoutCachedTop =960
                                    LayoutCachedWidth =13080
                                    LayoutCachedHeight =1440
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =12240
                    Top =120
                    Width =720
                    Height =354
                    FontSize =9
                    TabIndex =1
                    ForeColor =0
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close the form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
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
' FORM NAME:    frm_Lookups
' Description:  Standard form for viewing and editing lookup domains
' Data source:  unbound
' Data access:  edit only, no additions or deletions
' Pages:        pgVegWalk, pgProjectCrew, pgOtherLookups
' Functions:    fxnSwitchboardIsOpen, fxnTableExists
' References:   none
' Source/date:  John R. Boetsch, Jan 2006
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, May 5, 2006 - added page for other lookups, reset mode on change page
'               JRB, 6/21/2006 - added orderby for taxa subform on open, changed widths on
'                   contacts listbox, reversed order of new and view buttons on taxa page
'               JRB, June 2008 - updated form open to determine form settings by AppMode
'               JRB, 9/24/2008 - revised Form_Open to determine settings based on AppMode;
'                   updated row source for cmdTableFilter
'               JRB, 10/10/2008 - updated Form_Open and call to frm_Contacts
'               JRB, 7/9/2009 - updated cmbTableFilter to rely on tsys_Link_Tables, if present
'               BLC, 7/29/2014 - added documentation & updated to use TempVars.Item("UserAccessLevel") vs. cAppMode
' =================================

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
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Set up form depending on application mode
    If SwitchboardIsOpen Then
        'initialize controls based on app mode
        setUserAccess Me
    
    Else
        Me.cmdNewTaxon.Enabled = False
        Me.optgFormMode.Enabled = False
        Me.subLookupTables.Locked = True
    End If
    Me.subTaxa.Form.OrderBy = "tlu_Project_Taxa.Species_code"
    Me.subTaxa.Form.OrderByOn = True

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          PageTabs_Change
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub PageTabs_Change()
    On Error GoTo Err_Handler

    ' Revert to view mode upon changing pages
    Me.optgFormMode = 0
    optgFormMode_AfterUpdate

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
' PAGE NAME:    Species List (pgVegWalk)
' Description:  browse and edit species list
' Unbound ctls: cmdNewTaxon, cmdViewDetails
' Subforms:     subTaxa (fsub_Project_Taxa)
' =================================
' ---------------------------------
' SUB:          optgFormMode_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub optgFormMode_AfterUpdate()
    On Error GoTo Err_Handler

    ' Change the subform data mode depending on the user choice
    If Me.optgFormMode = 0 Then
    ' View mode
        Me.subTaxa.Locked = True
        Me.Detail.backcolor = 13025979 ' steel blue (default)
    Else
    ' Edit mode
        Me.subTaxa.Locked = False
        Me.Detail.backcolor = 12574431 ' haystack
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdNewTaxon_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmdNewTaxon_Click()
    On Error GoTo Err_Handler

    Dim strFormName As String

    ' Set the form name and global control variable for requerying after updates
    strFormName = "fsub_Project_Taxa"
    Set gvarRefTaxonCtl = Me.subTaxa

    ' Open the subform in a popup window to enter a new record
    DoCmd.OpenForm strFormName, , , , acFormAdd, acDialog

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdViewDetails_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmdViewDetails_Click()
    On Error GoTo Err_Handler

    Dim ctl As Control
    Dim strRecID As String

    Set ctl = Forms!frm_Lookups.subTaxa.Form!txtTaxonID
    strRecID = ctl.Value

    ' Set the global reference control variable for requerying after updates
    Set gvarRefTaxonCtl = Me.subTaxa
    ' Open the form to edit selected records
    DoCmd.OpenForm "fsub_Project_Taxa", , , "(tlu_Project_Taxa.Taxon_ID) = " & _
        strRecID, acFormEdit, acDialog

Exit_Procedure:
    On Error Resume Next
    Set ctl = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2427   ' No records in the subform
        ' Do nothing ...
      Case Else
        MsgBox Err.Number & ": " & Err.Description
    End Select
    Resume Exit_Procedure

End Sub

' =================================
' PAGE NAME:    Contacts List (pgProjectCrew)
' Description:  browse and edit project crew list
' Unbound ctls: cmdEditContacts, lstContacts
' Subforms:     none
' =================================
' ---------------------------------
' SUB:          cmdEditContacts_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmdEditContacts_Click()
    On Error GoTo Err_Handler

    ' View and edit contact records in a popup form

    Dim ctl As Control
    Dim varItem As Variant
    Dim strFormName As String
    Dim strFilter As String

    Set ctl = Me.lstContacts
    strFormName = "frm_Contacts"
    strFilter = ""

    ' Indicate the record to filter by if one is selected
    If IsNull(ctl.Value) = False Then
        strFilter = "[Contact_ID]=""" & ctl.Value & """"
        ' De-select the record
        ctl.Value = Null
    End If
    ' Set the global reference control variable for requerying after updates
    Set gvarRefContactCtl = ctl
    DoCmd.OpenForm strFormName, , , strFilter

Exit_Procedure:
    On Error Resume Next
    Set ctl = Nothing
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          lstContacts_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub lstContacts_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler
    
    cmdEditContacts_Click

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' PAGE NAME:    Other Lookups (pgOtherLookups)
' Description:  browse other lookup tables
' Unbound ctls: cmbTableFilter
' Subforms:     subLookupTables (unbound until a table is selected)
' =================================
' ---------------------------------
' SUB:          cmbTableFilter_Enter
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmbTableFilter_Enter()
     On Error GoTo Err_Handler

    Dim strSysTable As String

    strSysTable = "tsys_Link_Tables"     ' System table listing linked tables

    ' If the system table does not exist, replace the row source with one that doesn't use it
    If TableExists(strSysTable) = False Then
        Me.cmbTableFilter.RowSource = "SELECT MSysObjects.Name " & _
            "FROM MSysObjects " & _
            "WHERE (((MSysObjects.Name) Like 'tlu_*' " & _
            "And (MSysObjects.Name)<>'tlu_Project_Crew')) " & _
            "And (((MSysObjects.Name)<>'tlu_Project_Taxa')) " & _
            "And (((MSysObjects.Name)<>'tlu_Park_Taxa'));"
        Me.cmbTableFilter.ColumnCount = 1
        Me.cmbTableFilter.ListWidth = Me.cmbTableFilter.Width
        Me.cmbTableFilter.Requery
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbTableFilter_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - XX
' ---------------------------------
Private Sub cmbTableFilter_AfterUpdate()
    On Error GoTo Err_Handler

    ' Once a table is selected, bind the subform to this table
    If IsNull(Me.cmbTableFilter) Then
    ' If none selected ...
        Me.subLookupTables.SourceObject = ""
    Else:
    ' If a table is selected ...
        If TableExists(Me.cmbTableFilter) Then
            Me.subLookupTables.SourceObject = "Table." & Me.cmbTableFilter.Value
        Else
            MsgBox "Unable to find the selected table in the database ...", , _
                "Table not found"
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
