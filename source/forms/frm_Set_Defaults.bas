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
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =10
    ItemSuffix =9
    Left =4440
    Top =3105
    Right =23400
    Bottom =14895
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x13562948b201e340
    End
    Caption ="  Set application default values"
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
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =3120
            BackColor =12433075
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =972
                    Top =1320
                    Width =1245
                    Height =252
                    FontSize =9
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    Name ="cbxPark"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Parks"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=[Forms]![frm_Switchboard]![fsub_DbAdmin].[Form]![tbxPark]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1320
                            Width =444
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblPark"
                            Caption ="Park"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3600
                    Left =972
                    Top =600
                    Width =3225
                    Height =252
                    FontSize =9
                    TabIndex =1
                    Name ="cbxUser"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active],'Active','') AS Is_a"
                        "ctive FROM tlu_Project_Crew ORDER BY IIf([Contact_is_active],'Active','') DESC ,"
                        " tlu_Project_Crew.Contact_ID; "
                    ColumnWidths ="2592;1008"
                    DefaultValue ="=[Forms]![frm_Switchboard]![fsub_DbAdmin].[Form]![tbxUser]"
                    FontName ="Arial"
                    OnNotInList ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =600
                            Width =468
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblUser"
                            Caption ="User"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =837
                    Top =2760
                    Width =1023
                    Height =252
                    FontSize =9
                    TabIndex =8
                    Name ="tbxProject"
                    DefaultValue ="=[Forms]![frm_Switchboard]![fsub_DbAdmin].[Form]![tbxProject]"
                    FontName ="Arial"

                    LayoutCachedLeft =837
                    LayoutCachedTop =2760
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =3012
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2760
                            Width =672
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="labProject"
                            Caption ="Project"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3480
                    Top =120
                    Width =720
                    Height =354
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="cmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =2040
                    Width =1512
                    Height =252
                    FontSize =9
                    TabIndex =6
                    Name ="tbxDeclination"
                    DefaultValue ="=[Forms]![frm_Switchboard]![fsub_DbAdmin].[Form]![tbxDeclination]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2040
                            Width =1020
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblDeclination"
                            Caption ="Declination"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2040
                    Top =120
                    Width =1035
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="cmdNewUser"
                    Caption ="New user"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Add a new user"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =972
                    Top =1680
                    Width =1245
                    Height =285
                    FontSize =9
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"10\";\"10\""
                    Name ="cbxDatum"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Datum"
                    DefaultValue ="=[Forms]![frm_Switchboard]![fsub_DbAdmin].[Form]![tbxDatum]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1680
                            Width =672
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblDatum"
                            Caption ="Datum"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1620
                    Top =2400
                    Width =2460
                    Height =252
                    FontSize =9
                    TabIndex =7
                    Name ="tbxTimeframe"
                    StatusBarText ="Year corresponding to the current field season, or range of dates for the curren"
                        "t data set"
                    DefaultValue ="=[Forms]![frm_Switchboard]![fsub_DbAdmin].[Form]![tbxTimeframe]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =2400
                            Width =1497
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblTimeframe"
                            Caption ="Data timeframe"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListRows =20
                    Left =1197
                    Top =972
                    Width =3000
                    Height =252
                    FontSize =9
                    TabIndex =3
                    Name ="cbxGPS_model"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_GPS_Model.GPS_model FROM tlu_GPS_Model ORDER BY tlu_GPS_Model.Sort_or"
                        "der; "
                    DefaultValue ="=[Forms]![frm_Switchboard]![fsub_DbAdmin].[Form]![tbxGPS_model]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =960
                            Width =1005
                            Height =255
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblGPS_model"
                            Caption ="GPS model"
                            FontName ="Arial"
                        End
                    End
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
' FORM NAME:    frm_Set_Defaults
' Description:  Standard form for setting application defaults
' Data source:  unbound
' Data access:  edit only, no deletions
' Pages:        none
' Functions:    none
' References:   fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, May 16, 2006
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, 6/20/2008 - various updates and standardization
'               JRB, 9/4/2008 - updated Form_Open to include settings based on AppMode;
'                   changed cmbProject to txtProject; changed txtProject to be an unbound
'                   ctl and set its default value to the switchboard control; updated the
'                   validation and update code to include txtProject
'               JRB, 12/31/2009 - added Nz() to comparisons with switchboard values in cmdOK
'               --------------------------------------------------------------------------------------
'               BLC, 6/3/2014 - Adapted for NCPN WQ Utilities tool
'               BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel"),
'                    modified label & textbox control names to lblXX, cbxXX & tbxXX
'                    vs. labXX, cmbXX & txtXX
' =================================

Dim varOpenArgs As Variant

' ---------------------------------
' SUB:     Form_Open
' Description: Opens sub form & sets controls based on UserAccessLevel
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
'               BLC - 8/27/2014 - Realigned TempVars indices
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Set up form depending on application mode
    varOpenArgs = Me.OpenArgs
    
    If SwitchboardIsOpen Then
        setUserAccess Me
    Else
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this form to function properly.", , "Cannot open the form ..."
        DoCmd.CancelEvent
    End If
    
    'Set control default values
    Dim i As Integer
    Dim ctrlName As String
    Dim idxShift As Integer
 
    'iterate through control names in TempVars
    ' --------------------------------------------------------------------------
    ' 1-4 Skip UserAccessLevel, Connected, HasAccessBE, WritePermission
    '     these are not applicable, not controls (also action, analysis & other indices)
    ' 5 - User
    ' 6 - GPS_model
    ' 7 - Park
    ' 8 - Datum
    ' 9 - Declination
    ' 10 - Timeframe
    ' 11 - Project
    ' * TempVars indices shift depending on if certain vars are added
    '   idxShift accommodates based on the location of TempVars("User")
    ' --------------------------------------------------------------------------
    idxShift = GetTempVarIndex("User")
    
    'TempVars not yet populated -> use fsub_DbAdmin control defaults
    If idxShift = -1 Then
'        initializeControls Me
        GoTo Exit_Procedure
    End If
    
    For i = idxShift To idxShift + 6 ' (w/o shift) -> control values, beyond this TempVars has other values
                                     ' so can 't use TempVars.Count - 1, 0-based so number - 1
        With TempVars.item(i)
        
            If .Name = "Declination" Or _
               .Name = "Timeframe" Or _
               .Name = "Project" Then
                ctrlName = "tbx"
            Else
                ctrlName = "cbx"
            End If
            
            ctrlName = ctrlName & TempVars.item(i).Name
            
            Me.Controls(ctrlName) = TempVars.item(i).value
        End With
    Next

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cbxUser_NotInList
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 6/xx/2014 - XX
' ---------------------------------
Private Sub cbxUser_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    Me.ActiveControl.Undo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdNewUser_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 6/xx/2014 - XX
' ---------------------------------
Private Sub cmdNewUser_Click()
    On Error GoTo Err_Handler
    
    ' Set the global reference control variable for requerying after updates
    Set gvarRefContactCtl = Me.cbxUser
    ' Open the contacts form
    DoCmd.OpenForm "frm_Contacts", , , , , , "new"

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cbxPark_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC - 6/xx/2014 - XX
' ---------------------------------
Private Sub cbxPark_AfterUpdate()
    On Error GoTo Err_Handler

    Dim strDec As String
    Dim strDatum As String

    If IsNull(Me.cbxDatum) = False Or IsNull(Me.tbxDeclination) = False Then
    ' On changing the park, prompt for resetting the datum and declination
        If IsNull(Me.tbxDeclination) Then strDec = "---" Else: strDec = Me.tbxDeclination
        If IsNull(Me.cbxDatum) Then strDatum = "---" Else: strDatum = Me.cbxDatum
        If MsgBox("Please confirm the datum: " & strDatum, vbYesNo, "Confirm park info") _
            = vbNo Then Me.cbxDatum = Null
        If MsgBox("Please confirm the declination: " & strDec, vbYesNo, "Confirm park info") _
            = vbNo Then Me.tbxDeclination = Null
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdOK_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel")
'               BLC, 8/22/2014 - Removed setting tbxAppMode since this is set on
'                                initializing app & form open via getDbUserAccess()
'               BLC, 8/25/2014 - Added tsys_app_defaults update for selected values
'                                Access SQL doesn't have LIMIT 1 clause so use
'                                SELECT for username in WHERE clause instead to update only
'                                current record
'               BLC, 9/1/2014  - Added fsub_DbAdmin repaint to refresh controls with new values
' ---------------------------------
Private Sub cmdOK_Click()
    On Error GoTo Err_Handler
    Dim frm As Form
    Dim strSQL As String
    
    ' Make sure the information is valid before updating the record
    If varOpenArgs <> 0 Then
        '  Confirm that the critical data elements have been completed before saving
        If IsNull(Me.cbxUser) And TempVars.item("UserAccessLevel") <> "admin" Then
            MsgBox "Please indicate the user name", vbOKOnly, "Validation error"
            Me.cbxUser.SetFocus
            GoTo Exit_Procedure
        ElseIf IsNull(Me.cbxPark) Then
            MsgBox "Please indicate the park", vbOKOnly, "Validation error"
            Me.cbxPark.SetFocus
            GoTo Exit_Procedure
        ElseIf IsNull(Me.cbxDatum) Then
            MsgBox "Please indicate the datum", vbOKOnly, "Validation error"
            Me.cbxDatum.SetFocus
            GoTo Exit_Procedure
        ElseIf IsNull(Me.tbxTimeframe) Then
            MsgBox "Please indicate the data timeframe", vbOKOnly, "Validation error"
            Me.tbxTimeframe.SetFocus
            GoTo Exit_Procedure
        ElseIf IsNull(Me.tbxProject) Then
            MsgBox "Please indicate the project code", vbOKOnly, "Validation error"
            Me.tbxProject.SetFocus
            GoTo Exit_Procedure
        End If
    End If

    ' Save changes to the switchboard record
    If Nz(Me.cbxUser) <> Nz(TempVars.item("User")) Then TempVars.item("User") = Me.cbxUser.value
    If Nz(Me.cbxGPS_model) <> Nz(TempVars.item("GPS_model")) Then TempVars.item("GPS_model") = Me.cbxGPS_model.value
    If Nz(Me.cbxPark) <> Nz(TempVars.item("Park")) Then TempVars.item("Park") = Me.cbxPark.value
    If Nz(Me.cbxDatum) <> Nz(TempVars.item("Datum")) Then TempVars.item("Datum") = Me.cbxDatum.value
    If Nz(Me.tbxDeclination) <> Nz(TempVars.item("Declination")) Then TempVars.item("Declination") = Me.tbxDeclination.value
    If Nz(Me.tbxTimeframe) <> Nz(TempVars.item("Timeframe")) Then TempVars.item("Timeframe") = Me.tbxTimeframe.value
    If Nz(Me.tbxProject) <> Nz(TempVars.item("Project")) Then TempVars.item("Project") = Me.tbxProject.value

    strSQL = "UPDATE tsys_App_Defaults " _
        & "SET GPS_model = '" & TempVars.item("GPS_model") & "', " _
        & "Park = '" & TempVars.item("Park") & "', " _
        & "Datum = '" & TempVars.item("Datum") & "', " _
        & "Declination = '" & TempVars.item("Declination") & "', " _
        & "Data_timeframe = " & TempVars.item("Timeframe") & ", " _
        & "Project = '" & TempVars.item("Project") & "' " _
        & "WHERE User_name IN (" _
        & " SELECT TOP 1 User_name FROM tsys_App_Defaults " _
        & " WHERE User_name = '" & TempVars.item("User") & "' " _
        & " ORDER BY User_name);"
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
    
    'repaint fsub_DbAdmin to update values if switchboard is open
    If SwitchboardIsOpen Then Forms!frm_Switchboard.fsub_DBAdmin.Form.Repaint
    
    'close form & return to calling form
    DoCmd.Close , , acSaveNo

    ' Open the form specified by the open arguments
    Select Case varOpenArgs
      Case 1
        DoCmd.OpenForm "frm_Data_Gateway"
      Case 2
        DoCmd.OpenForm "frm_Data_Browser"
      Case 3
        DoCmd.OpenForm "frm_QA_Tool"
      Case 4
        ' opened by switchboard only ... do nothing
      Case Else
        MsgBox "Unexpected open arguments - please notify the database administrator"
    End Select

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
