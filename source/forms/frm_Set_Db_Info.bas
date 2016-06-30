Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
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
    Width =5520
    DatasheetFontHeight =10
    ItemSuffix =14
    Left =4665
    Top =3315
    Right =12315
    Bottom =14310
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x99f562689172e440
    End
    RecordSource ="tsys_App_Releases"
    Caption ="Set Application Version Info"
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
            Height =3840
            BackColor =14341081
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =8640
                    Left =1320
                    Top =720
                    Width =3900
                    Height =252
                    FontSize =9
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cbxVersion"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tsys_App_Releases.Release_ID, 'Version ' & [Version_number] & ' (' & [Rel"
                        "ease_date] & ')' AS Version, IIf([Is_supported]=0,'Not supported',IIf([Is_suppor"
                        "ted]=1,'Supported','Current')) AS Supported, tsys_App_Releases.Database_title FR"
                        "OM tsys_App_Releases ORDER BY tsys_App_Releases.Release_date DESC; "
                    ColumnWidths ="0;2880;1440;4320"
                    StatusBarText ="Set the current application version"
                    DefaultValue ="=[Forms]![frm_Switchboard]![cmbVersion]"
                    FontName ="Arial"
                    OnGotFocus ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    ControlTipText ="Set the current application version"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =720
                            Width =1065
                            Height =255
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblVersion"
                            Caption ="Db version: "
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4380
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
                    Left =1860
                    Top =1200
                    Width =2694
                    Height =252
                    FontSize =9
                    TabIndex =2
                    Name ="tbxContactName"
                    StatusBarText ="Enter the developer contact name for the application"
                    FontName ="Arial"
                    ControlTipText ="Enter the developer contact name for the application"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =480
                            Top =1200
                            Width =1305
                            Height =255
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblContactName"
                            Caption ="Contact name:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =1560
                    Width =2700
                    Height =252
                    FontSize =9
                    TabIndex =3
                    Name ="tbxContactOrg"
                    StatusBarText ="Enter the contact organization for the application"
                    FontName ="Arial"
                    ControlTipText ="Enter the contact organization for the application"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =480
                            Top =1560
                            Width =1257
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblContactOrg"
                            Caption ="Contact org:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =1932
                    Width =2700
                    Height =252
                    FontSize =9
                    TabIndex =4
                    Name ="tbxContactPhone"
                    StatusBarText ="Enter the contact phone number"
                    FontName ="Arial"
                    ControlTipText ="Enter the contact phone number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =480
                            Top =1932
                            Width =1317
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="labContactPhone"
                            Caption ="Contact phone:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =2304
                    Width =2700
                    Height =252
                    FontSize =9
                    TabIndex =5
                    Name ="tbxContactEmail"
                    StatusBarText ="Enter the contact organization for the application"
                    FontName ="Arial"
                    ControlTipText ="Enter the contact organization for the application"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =480
                            Top =2304
                            Width =1317
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblContactEmail"
                            Caption ="Contact email:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1980
                    Top =120
                    Width =1980
                    Height =354
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    Name ="cmdReleaseHistory"
                    Caption ="View release history"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =2940
                    Width =5280
                    Height =720
                    FontSize =9
                    TabIndex =7
                    Name ="tbxWeb_address"
                    StatusBarText ="Web address for application downloads"
                    FontName ="Arial"

                    LayoutCachedLeft =120
                    LayoutCachedTop =2940
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =3660
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2640
                            Width =4245
                            Height =270
                            FontSize =9
                            Name ="lblWeb_address"
                            Caption ="Web address to which version updates are posted:"
                            FontName ="Arial"
                            LayoutCachedLeft =120
                            LayoutCachedTop =2640
                            LayoutCachedWidth =4365
                            LayoutCachedHeight =2910
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
' FORM NAME:    frm_Set_Db_Info
' Description:  Standard form for setting application version and contact information
' Data source:  unbound
' Data access:  edit only, no deletions
' Pages:        none
' Functions:    none
' References:   fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, September 2008
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    JRB, 9/26/2008 - revised Form_Open to turn edits off except for AppMode='admin'
'               JRB, 10/2/2008 - added web address and updated code for cmdReleaseHistory
'               --------------------------------------------------------------------------------------
'               BLC, 6/3/2014 - Adapted for NCPN WQ Utilities tool
'               BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel"),
'                    modified label & textbox control names to lblXX, cbxXX & tbxXX
'                    vs. labXX, cmbXX & txtXX
' =================================


' ---------------------------------
' SUB:     Form_Open
' Description: Opens sub form & sets controls based on UserAccessLevel
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel"),
'                    modified label & textbox control names to lblXX, cbxXX & tbxXX
'                    vs. labXX, cmbXX & txtXX
'               BLC, 8/25/2014 - adjusted source for form fields to fsub_DbAdmin vs frm_Switchboard
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Close if switchboard is not open
    If SwitchboardIsOpen = False Then
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this form to function properly.", , "Cannot open the form ..."
        DoCmd.CancelEvent
    Else
        With Me
            If TempVars.Item("UserAccessLevel") = "admin" Then
                .AllowEdits = True
                .tbxWeb_address.visible = True
            Else
                .AllowEdits = False
                .tbxWeb_address.visible = False
            End If
            
            With Forms!frm_Switchboard!fsub_DbAdmin.Form
                Me.tbxContactName = !tbxContact_Name
                Me.tbxContactOrg = !tbxContact_Org
                Me.tbxContactPhone = !tbxContact_Phone
                Me.tbxContactEmail = !tbxContact_Email
                Me.tbxWeb_address = !tbxWeb_address
            End With
        End With
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cbxVersion_GotFocus
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel"),
'                    modified label & textbox control names to lblXX, cbxXX & tbxXX
'                    vs. labXX, cmbXX & txtXX
' ---------------------------------
Private Sub cbxVersion_GotFocus()
    On Error GoTo Err_Handler

    Me.ActiveControl.Requery

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cbxVersion_NotInList
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - XX
' ---------------------------------
Private Sub cbxVersion_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    Me.ActiveControl.Undo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmdReleaseHistory_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John Boetsch - NCCN Landbirds db
' Adapted:      Bonnie Campbell, May 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel"),
'                    modified control names to lblXX, cbxXX & tbxXX vs. labXX, cmbXX & txtXX
' ---------------------------------
Private Sub cmdReleaseHistory_Click()
    On Error GoTo Err_Handler

    ' View the release history form
    If TempVars.Item("UserAccessLevel") = "admin" Then
        DoCmd.OpenForm "frm_App_Releases"
    Else    ' read-only for all but admin users
        DoCmd.OpenForm "frm_App_Releases", , , , acFormReadOnly
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
' Revisions:    BLC, 6/19/2014 - Replaced cAppMode with TempVars.Item("UserAccessLevel"),
'                    modified label & textbox control names to lblXX, cbxXX & tbxXX
'                    vs. labXX, cmbXX & txtXX
'               BLC, 8/25/2014 - Fixed frm to refer to fsub_DbAdmin vs frm_Switchboard
' ---------------------------------
Private Sub cmdOK_Click()
    On Error GoTo Err_Handler

    Dim frm As Form

    '  Confirm that the critical data elements have been completed before saving
    If IsNull(Me.cbxVersion) Then
        MsgBox "Please indicate the version number", vbOKOnly, "Validation error"
        Me.cbxVersion.SetFocus
        GoTo Exit_Procedure
    ElseIf IsNull(Me.tbxContactName) Then
        MsgBox "Please enter the developer contact name", vbOKOnly, "Validation error"
        Me.tbxContactName.SetFocus
        GoTo Exit_Procedure
    ElseIf IsNull(Me.tbxContactPhone) Then
        MsgBox "Please enter a contact phone number", vbOKOnly, "Validation error"
        Me.tbxContactPhone.SetFocus
        GoTo Exit_Procedure
    ElseIf IsNull(Me.tbxContactEmail) Then
        MsgBox "Please enter a contact email", vbOKOnly, "Validation error"
        Me.tbxContactEmail.SetFocus
        GoTo Exit_Procedure
    End If

    Set frm = [Forms]![frm_Switchboard]![fsub_DbAdmin].[Form]

    ' Save changes to the switchboard
    If Me.cbxVersion <> frm.cbxVersion Then frm.cbxVersion = Me.cbxVersion
    If Me.tbxContactName <> frm.tbxContact_Name Then frm.tbxContact_Name = Me.tbxContactName
    If Me.tbxContactOrg <> frm.tbxContact_Org Then frm.tbxContact_Org = Me.tbxContactOrg
    If Me.tbxContactPhone <> frm.tbxContact_Phone Then frm.tbxContact_Phone = Me.tbxContactPhone
    If Me.tbxContactEmail <> frm.tbxContact_Email Then frm.tbxContact_Email = Me.tbxContactEmail
    If Me.tbxContactEmail <> frm.tbxContact_Email Then _
        frm.tbxContact_Email.Hyperlink = "mailto:" & Me.tbxContactEmail = Me.tbxContactEmail
    If Me.tbxWeb_address <> frm.tbxWeb_address Then frm.tbxWeb_address = Me.tbxWeb_address

    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
