Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =26520
    DatasheetFontHeight =10
    ItemSuffix =23
    Left =465
    Top =2820
    Right =14355
    Bottom =12600
    DatasheetGridlinesColor =12632256
    Filter ="[Is_done] = False"
    OrderBy ="Query_name"
    RecSrcDt = Begin
        0x5dcbdb2ab6c4e440
    End
    RecordSource ="SELECT tbl_QA_Results.* FROM tbl_QA_Results ORDER BY IIf([Is_done],2,1), tbl_QA_"
        "Results.Query_name; "
    Caption ="fsub_QA_Results"
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
    FilterOnLoad =255
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
            CanGrow = NotDefault
            Height =300
            BackColor =13025979
            Name ="FormHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =60
                    Width =1200
                    Height =240
                    FontWeight =700
                    Name ="lblQuery_name"
                    Caption ="Query name*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4680
                    Top =60
                    Width =624
                    Height =240
                    Name ="lblQuery_type"
                    Caption ="Type*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6240
                    Top =60
                    Width =654
                    Height =240
                    Name ="lblQuery_result"
                    Caption ="N recs*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7042
                    Top =60
                    Width =1005
                    Height =240
                    Name ="lblQuery_run_time"
                    Caption ="Last run time"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8280
                    Top =60
                    Width =870
                    Height =240
                    Name ="lblQuery_description"
                    Caption ="Description"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =14880
                    Top =60
                    Width =1059
                    Height =240
                    Name ="lblRemedy_desc"
                    Caption ="Action taken"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =19020
                    Top =60
                    Width =1044
                    Height =240
                    Name ="lblQA_user"
                    Caption ="Remedy by*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =21492
                    Top =60
                    Width =1350
                    Height =240
                    Name ="lblQuery_expression"
                    Caption ="Query expression"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =20112
                    Top =60
                    Width =1164
                    Height =240
                    Name ="lblRemedy_date"
                    Caption ="Remedy date*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5640
                    Top =60
                    Width =540
                    Height =240
                    Name ="lblIs_done"
                    Caption ="Done*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =300
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =4500
                    ColumnWidth =3000
                    ForeColor =16711680
                    Name ="tbxQuery_name"
                    ControlSource ="Query_name"
                    StatusBarText ="Name of the quality assurance query"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x0100000086000000010000000100000000000000000000001200000001010000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00510075006500720079005f0072006500730075006c00740073005d003d00 ,
                        0x300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ffffff00110000005b00 ,
                        0x510075006500720079005f0072006500730075006c00740073005d003d003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6202
                    Top =60
                    Width =600
                    ColumnWidth =2568
                    TabIndex =2
                    Name ="tbxQuery_result"
                    ControlSource ="Query_result"
                    StatusBarText ="Query result as the number of records returned the last time the query was run"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000002000000000000000200000001000000 ,
                        0x00000000ffffff0000000000040000000300000005000000010100000000ff00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000200000000000000020000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010100 ,
                        0x000000ff00ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6862
                    Top =60
                    Width =1320
                    ColumnWidth =1896
                    TabIndex =3
                    Name ="tbxQuery_run_time"
                    ControlSource ="Query_run_time"
                    Format ="mm/dd/yyyy hh:nn"
                    StatusBarText ="Run time of the query results"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8220
                    Top =60
                    Width =6600
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="tbxQuery_description"
                    ControlSource ="Query_description"
                    StatusBarText ="Description of the query"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14880
                    Top =60
                    Width =4080
                    ColumnWidth =3000
                    TabIndex =5
                    Name ="tbxRemedy_desc"
                    ControlSource ="Remedy_desc"
                    StatusBarText ="Details about actions taken and/or not taken to resolve errors"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19020
                    Top =60
                    Width =960
                    ColumnWidth =2568
                    TabIndex =6
                    Name ="tbxQA_user"
                    ControlSource ="QA_user"
                    StatusBarText ="Name of the person doing quality assurance"
                    FontName ="Arial"

                End
                Begin ComboBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4620
                    Top =60
                    Width =1140
                    ColumnWidth =2568
                    TabIndex =1
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000002000000000000000400000001010000 ,
                        0xff000000ffffff00000000000200000005000000090000000101010080008000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x2200310022000000000022003200220000000000
                    End
                    Name ="cbxQuery_type"
                    ControlSource ="Query_type"
                    RowSourceType ="Value List"
                    RowSource ="1;Critical;2;Warning;3;Information"
                    ColumnWidths ="0;2160"
                    StatusBarText ="Severity of data errors being trapped: 1=critical, 2=warning, 3=information"
                    FontName ="Arial"

                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001010000ff000000ffffff00030000002200 ,
                        0x3100220000000000000000000000000000000000000000000000000000020000 ,
                        0x000101010080008000ffffff0003000000220032002200000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21420
                    Top =60
                    Width =5100
                    TabIndex =8
                    Name ="tbxQuery_expression"
                    ControlSource ="Query_expression"
                    StatusBarText ="Description of the query"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20040
                    Top =60
                    Width =1320
                    ColumnWidth =1680
                    TabIndex =7
                    Name ="tbxRemedy_date"
                    ControlSource ="Remedy_date"
                    Format ="mm/dd/yyyy hh:nn"
                    StatusBarText ="When the remedy description was last edited"
                    FontName ="Arial"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5880
                    Top =60
                    Width =240
                    TabIndex =9
                    Name ="chkIs_done"
                    ControlSource ="Is_done"
                    ControlTipText ="Temporary flag to indicate that the user is done reviewing this query even if so"
                        "me records remain"

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
' FORM NAME:    fsub_QA_Results
' Description:  Standard subform for viewing data validation results
' Data source:  tbl_QA_Results
' Data access:  edit only, no deletions
' Pages:        none
' Functions:    fxnOpenClickedQuery, fxnSortRecords
' References:   none
' Source/date:  John R. Boetsch, May 5, 2006
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, 12/17/2007 - updated fxnOpenClickedQuery; also added Is_done field
'               JRB, 7/1/2008 - updated fxnSortRecords, added strSortFieldLabel2
'               JRB, 10/7/2008 - removed Form_DblClick; changed single click events on query
'                   name and results fields to double click events to avoid a run-time error
'               JRB, 11/12/2008 - added an error trap to fxnOpenClickedQuery
'               JRB, 2/23/2009 - added a condition to fxnOpenClickedQuery so query results will
'                   not be displayed if the selected timeframe does not equal the db timeframe
' =================================

Dim strSortField As String    ' Keeps track of current sort settings
Dim strSortOrder As String
Dim strSortFieldLabel As String
Dim strSortFieldLabel2 As String

' ---------------------------------
' SUB:     Form_Open
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim varReturn As Variant

    ' On opening the form, set the initial sort order
    strSortFieldLabel = "lblQuery_name"
    varReturn = fxnSortRecords("Query_name")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' The following several procedures re-sort the records if the user
'   double-clicks on a field label

' ---------------------------------
' SUB:          lblIs_done_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub lblIs_done_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Is_done")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          lblQuery_name_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub lblQuery_name_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Query_name")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          lblQuery_type_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub lblQuery_type_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Query_type")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub lblQuery_result_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Query_result")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          lblQA_user_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub lblQA_user_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("QA_user")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          lblRemedy_date_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub lblRemedy_date_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Remedy_date")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' The following several procedures open the selected query in the parent
'   form after the user clicks

' ---------------------------------
' SUB:          tbxQuery_name_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub tbxQuery_name_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnOpenClickedQuery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          tbxQuery_result_DblClick
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - XX
' ---------------------------------
Private Sub tbxQuery_result_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnOpenClickedQuery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' FUNCTION:     fxnOpenClickedQuery
' Description:  opens the selected record for viewing and editing results
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 5, 2006
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, 8/2/2006 - added an error trap for error 2113
'               JRB, 12/17/2007 - updated the parent form filter to also filter on
'                   time frame
'               JRB, 11/12/2008 - added an error trap in case of missing query object
'               JRB, 2/23/2009 - added a condition to not display query results if the timeframe
'                   doesn't match the current timeframe string
'               BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
' ---------------------------------
Private Function fxnOpenClickedQuery()
    On Error GoTo Err_Handler

    ' Make sure a query is selected
    If IsNull(Me.tbxQuery_name) Then
        MsgBox "No query selected", vbOKOnly
    Else
        ' Set the object selector to the selected query
        Me.Parent.Form!selObject = Me.tbxQuery_name
        ' Bind the subform to the selected query - only if matching the current timeframe
'FIX If Me.Time_frame = TempVars.item("Timeframe") Then
        If Me.Parent.Form!cbxTimeframe = TempVars.item("Timeframe") Then
            Me.Parent.Form!subQueryResults.SourceObject = "Query." & Me.tbxQuery_name
        End If
        ' Set the form to the selected record
        Me.Parent.Form.Filter = "[Query_name] = """ & Me.tbxQuery_name & _
            """ AND [Time_frame] = """ & Me.Parent.Form!cbxTimeframe & """"
'            """ AND [Time_frame] = """ & Me.Time_frame & """"
        Me.Parent.Form.FilterOn = True
        Me.Parent.Form.AllowAdditions = False
        Me.Parent.Form.tabQueries.SetFocus 'Me.Parent.pgQueryViews.SetFocus
        Me.Form.OrderByOn = True
    End If

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 2113   ' No results set records available
        ' Do nothing
      Case 3011, 7874   ' The query object isn't found
        MsgBox "This query is no longer available in the application." & _
            vbCrLf & """" & Me.tbxQuery_name & """", , "Query not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnOpenClickedQuery)"
    End Select
    Resume Exit_Procedure

End Function

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
' Revisions:    JRB, 7/1/2008 - documentation and sorting by second field
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
        strOrderBy = strOrderBy & ", " & strField2Name
    End If
    strSortField = strFieldName
    Me.Form.OrderBy = strOrderBy
    Me.Form.OrderByOn = True

    ' Change the label format to indicate the sorted field
    Me.Controls.item(strSortFieldLabel).FontItalic = False
    Me.Controls.item(strSortFieldLabel).fontBold = False
    strSortFieldLabel = "lbl" & strFieldName
    Me.Controls.item(strSortFieldLabel).FontItalic = True
    Me.Controls.item(strSortFieldLabel).fontBold = True
    ' Do the same for the second sort field, if applicable
    If strField2Name <> "" Then
        Me.Controls.item(strSortFieldLabel2).FontItalic = False
        Me.Controls.item(strSortFieldLabel2).fontBold = False
        strSortFieldLabel = "lbl" & strField2Name
        Me.Controls.item(strSortFieldLabel2).FontItalic = True
        Me.Controls.item(strSortFieldLabel2).fontBold = True
    End If

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnSortRecords)"
    End Select
    Resume Exit_Procedure

End Function
