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
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =110
    Left =5280
    Top =2715
    Right =16080
    Bottom =8760
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xabea2039169de340
    End
    RecordSource ="SELECT tsys_Link_Dbs.* FROM tsys_Link_Dbs ORDER BY tsys_Link_Dbs.Sort_order, tsy"
        "s_Link_Dbs.Link_db;"
    Caption =" Update Database Connections"
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
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin FormHeader
            Height =1020
            BackColor =11258796
            Name ="FormHeader"
            Begin
                Begin Label
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =2
                    Left =3240
                    Top =60
                    Width =4314
                    Height =276
                    FontSize =9
                    FontWeight =700
                    Name ="lblTitle"
                    Caption ="Update links to back-end databases"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =480
                    Width =9540
                    Height =420
                    FontSize =9
                    Name ="lblFormDesc"
                    Caption ="Data tables are stored in one or more separate database files.  Use the browse b"
                        "utton to update the database connections for Access back-ends, or indicate the n"
                        "ew server and db name for SQL Server / ODBC connections."
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9660
                    Top =60
                    Width =1023
                    Height =324
                    FontSize =9
                    FontWeight =700
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =8220
                    Top =60
                    Width =1263
                    Height =324
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="btnUpdateLinks"
                    Caption ="Update links"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Update links to the file(s) indicated"

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
            Height =2520
            BackColor =13301748
            Name ="Detail"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =240
                    Top =840
                    Width =843
                    Height =324
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    Name ="btnBrowse"
                    Caption ="Browse"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Browse to a new back-end file"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =0
                    BorderWidth =3
                    OverlapFlags =93
                    BackStyle =0
                    Left =180
                    Top =120
                    Width =5334
                    Height =252
                    ColumnWidth =3090
                    FontSize =9
                    FontWeight =700
                    ForeColor =16711680
                    Name ="tbxLink_db"
                    ControlSource ="Link_db"
                    StatusBarText ="Linked database name"
                    FontName ="Arial"
                    ControlTipText ="Linked database name"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =93
                    Left =1380
                    Top =480
                    Width =7917
                    Height =252
                    ColumnWidth =6630
                    FontSize =9
                    TabIndex =3
                    Name ="tbxDb_desc"
                    ControlSource ="Db_desc"
                    StatusBarText ="Brief description of the type of database - e.g., data tables, lookup tables, sy"
                        "stems tables, etc."
                    FontName ="Arial"
                    ControlTipText ="Brief description of the type of database - e.g., data tables, lookup tables, sy"
                        "stems tables, etc."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =240
                            Top =480
                            Width =1065
                            Height =240
                            FontSize =9
                            Name ="lblDb_desc"
                            Caption ="Description:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    BackStyle =0
                    Left =1080
                    Top =1740
                    Width =3060
                    Height =252
                    ColumnWidth =2520
                    FontSize =9
                    TabIndex =11
                    Name ="tbxServer"
                    ControlSource ="Server"
                    StatusBarText ="Server name (for ODBC links)"
                    FontName ="Arial"
                    ControlTipText ="Server name (for ODBC links)"

                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    BackStyle =0
                    Left =1680
                    Top =900
                    Width =8997
                    Height =252
                    ColumnWidth =2205
                    TabIndex =7
                    Name ="tbxFile_path"
                    ControlSource ="File_path"
                    StatusBarText ="Full path to back-end file (for Access back-ends)"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Full path to back-end file (for Access back-ends)"

                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =114
                    Top =63
                    Width =10623
                    Height =2400
                    Name ="Box82"
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =247
                    Left =1680
                    Top =1260
                    Width =8997
                    Height =252
                    TabIndex =9
                    Name ="tbxNew_path"
                    ControlSource ="New_path"
                    StatusBarText ="New path to back-end file (for Access back-ends)"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="New path to back-end file (for Access back-ends)"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =247
                    Left =6480
                    Top =2100
                    Width =3723
                    Height =252
                    FontSize =9
                    TabIndex =15
                    Name ="tbxNew_db"
                    ControlSource ="New_db"
                    StatusBarText ="New database name"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="New database name"
                    OnDirty ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =7080
                    Top =150
                    ColumnWidth =960
                    TabIndex =2
                    Name ="cbxIs_ODBC"
                    ControlSource ="Is_ODBC"
                    StatusBarText ="Indicates whether this back-end is linked through ODBC (e.g., SQL Server)"
                    BeforeUpdate ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"
                    ControlTipText ="Indicates whether this back-end is linked through ODBC (e.g., SQL Server)"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =150
                    LayoutCachedWidth =7340
                    LayoutCachedHeight =390
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7305
                            Top =120
                            Width =1695
                            Height =240
                            FontSize =9
                            Name ="lblIs_ODBC"
                            Caption ="ODBC / SQL Server"
                            FontName ="Arial"
                            LayoutCachedLeft =7305
                            LayoutCachedTop =120
                            LayoutCachedWidth =9000
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =9240
                    Top =150
                    ColumnWidth =900
                    TabIndex =4
                    Name ="cbxBackups"
                    ControlSource ="Backups"
                    StatusBarText ="Indicates whether this back-end gets backed up by user-initiated front-end backu"
                        "ps"
                    BeforeUpdate ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"
                    ControlTipText ="Indicates whether this back-end gets backed up by user-initiated front-end backu"
                        "ps"

                    LayoutCachedLeft =9240
                    LayoutCachedTop =150
                    LayoutCachedWidth =9500
                    LayoutCachedHeight =390
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =9465
                            Top =120
                            Width =1155
                            Height =240
                            FontSize =9
                            Name ="lblBackups"
                            Caption ="File backups"
                            FontName ="Arial"
                            LayoutCachedLeft =9465
                            LayoutCachedTop =120
                            LayoutCachedWidth =10620
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =247
                    Left =6480
                    Top =1740
                    Width =3240
                    Height =252
                    FontSize =9
                    TabIndex =13
                    Name ="tbxNew_server"
                    ControlSource ="New_server"
                    StatusBarText ="New server name (for ODBC links)"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="New server name (for ODBC links)"
                    OnDirty ="[Event Procedure]"

                End
                Begin Line
                    OverlapFlags =119
                    Left =360
                    Top =1620
                    Width =9420
                    Name ="Line101"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =1500
                    Top =2040
                    Width =1923
                    Height =324
                    FontSize =9
                    FontWeight =700
                    TabIndex =16
                    Name ="btnTestODBC"
                    Caption ="Test connection"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Test the new ODBC connection"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =6240
                    Top =120
                    Width =420
                    FontSize =9
                    TabIndex =1
                    Name ="tbxSort_order"
                    ControlSource ="Sort_order"
                    StatusBarText ="Display sort order"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5700
                            Top =120
                            Width =480
                            Height =240
                            FontSize =9
                            Name ="lblSort_order"
                            Caption ="Sort:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Line
                    BorderWidth =4
                    OverlapFlags =85
                    Width =10800
                    Name ="Line104"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =1740
                    Width =840
                    FontSize =9
                    TabIndex =10
                    Name ="lblServer"
                    ControlSource ="='Server:'"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x0100000080000000010000000100000000000000000000000f00000001010000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490073005f004f004400420043005d003d00540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ffffff000e0000005b00 ,
                        0x490073005f004f004400420043005d003d005400720075006500000000000000 ,
                        0x000000000000000000000000000000
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4200
                    Top =1740
                    Width =2220
                    Height =270
                    FontSize =9
                    TabIndex =12
                    Name ="lblNew_server"
                    ControlSource ="='New server (ODBC only):'"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x0100000080000000010000000100000000000000000000000f00000001010000 ,
                        0x0000ff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490073005f004f004400420043005d003d00540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010100000000ff00ffffff000e0000005b00 ,
                        0x490073005f004f004400420043005d003d005400720075006500000000000000 ,
                        0x000000000000000000000000000000
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4080
                    Top =2100
                    Width =2340
                    Height =270
                    FontSize =9
                    TabIndex =14
                    Name ="lblNew_db"
                    ControlSource ="='New db name (ODBC only):'"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x0100000080000000010000000100000000000000000000000f00000001010000 ,
                        0x0000ff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490073005f004f004400420043005d003d00540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010100000000ff00ffffff000e0000005b00 ,
                        0x490073005f004f004400420043005d003d005400720075006500000000000000 ,
                        0x000000000000000000000000000000
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Top =900
                    Width =495
                    Height =270
                    FontSize =9
                    TabIndex =6
                    Name ="lblFile_path"
                    ControlSource ="='Path:'"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x0100000082000000010000000100000000000000000000001000000001010000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490073005f004f004400420043005d003d00460061006c00730065000000 ,
                        0x0000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ffffff000f0000005b00 ,
                        0x490073005f004f004400420043005d003d00460061006c007300650000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =660
                    Top =1260
                    Width =960
                    Height =255
                    FontSize =9
                    TabIndex =8
                    Name ="lblNew_path"
                    ControlSource ="='New path:'"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x0100000082000000010000000100000000000000000000001000000001010000 ,
                        0x0000ff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490073005f004f004400420043005d003d00460061006c00730065000000 ,
                        0x0000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010100000000ff00ffffff000f0000005b00 ,
                        0x490073005f004f004400420043005d003d00460061006c007300650000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =11258796
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
' FORM NAME:    frm_Connect_Dbs
' Description:  Standard form for updating back-end db connections
' Data source:  tsys_Link_Dbs
' Data access:  edit only, no additions, moving between records, or deletions
' Pages:        none
' Functions:    none
' References:   fxnAppSetup, fxnFileExists, fxnFormIsLoaded, fxnGetFile, fxnParseConnectionStr,
'                   fxnParseFileName, fxnParsePath, fxnRefreshLinks, fxnReplaceString,
'                   fxnSwitchboardIsOpen, fxnTestODBCConnection, fxnVerifyLinkTableInfo
' Source/date:  Susan Huse, MonitoringSM.mdb v 7/28/2004
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    John R. Boetsch, May 2005 - minor edits
' Revisions:    JRB, 5/24/2006 - documentation, added error trapping, fixed specification
'                   of initial directory to current directory, simplified a little
'               JRB, 7/27/2006 - added code to cmdUpdateLinks_Click to close and reopen the
'                   always-open back-end connection form upon successfully reconnecting
'               JRB, 7/9/2009 - significant changes to accommodate .accdb back-ends, ODBC
'                   connections, and multiple back-ends
'               JRB, 8/6/2009 - fixed a minor glitch in cmdTestODBC
'               BLC, 6/16/2014 - updated to use TempVars.Item("UserAccessLevel") vs. cAppMode ,
'                    documentation
'               -------------------------
'               BLC, 5/21/2015 - added from NCPN WQ Utilities to invasives reporting tool
'                    updated documentation & error handlers to reflect module & function/sub
'                    updated control prefixes
' =================================

' ---------------------------------
' SUB:     Form_Open
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2005
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & updated fxnSwitchboardIsOpen to FormIsOpen function
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    If DB_ADMIN_CONTROL Then
        ' Close the form if the switchboard is not open
        If FormIsOpen("frm_Switchboard") = False Then
            MsgBox "The main database switchboard must be" & vbCrLf & _
                "open for this form to function properly.", , "Cannot open the form ..."
            DoCmd.CancelEvent
            GoTo Exit_Sub
        End If
    End If

    ' Enable/disable subforms depending on user access level
    Select Case TempVars("UserAccessLevel")
      Case "admin"
        Me.cbxIs_ODBC.Locked = False
      Case Else
        Me.cbxIs_ODBC.Locked = True
    End Select

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxIs_ODBC_KeyDown
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, May, 2015 for NCPN invasives reporting tool
' Revisions:    BLC, 5/21/2015 - initial version
' ---------------------------------
Private Sub cbxIs_ODBC_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler

    Dim strMsgConfirm As String

    If cbxIs_ODBC = False Then MsgBox "false", vbCritical, "Is ODBC checked"

    strMsgConfirm = "To change the connection type, you must delete and relink" & vbCrLf & _
        "the tables for this database manually ... are you sure?"
    If MsgBox(strMsgConfirm, vbOKCancel, "Confirm change") = vbCancel Then
        Cancel = True
        Me.ActiveControl.Undo
    Else
        'disable cbxBackups control
        ToggleControl Me.name, "cbxBackups"
        ToggleControl Me.name, "lblBackups", lngLtGray
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxIs_ODBC_KeyDown[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxIs_ODBC_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed chk prefix to cbx
' ---------------------------------
Private Sub cbxIs_ODBC_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim strMsgConfirm As String

    'any modifications
    strMsgConfirm = "To change the connection type, you must delete and relink" & vbCrLf & _
        "the tables for this database manually ... are you sure?"
    
    If MsgBox(strMsgConfirm, vbOKCancel, "Confirm change") = vbCancel Then
        Cancel = True
        Me.ActiveControl.Undo
    End If

    'check for not ODBC (false)
    If Me.ActiveControl = False Then
        
        'enable cbxBackups control
        ToggleControl Me.name, "cbxBackups"
        ToggleControl Me.name, "lblBackups", lngBlack
    
    Else
    
    'ODBC (true)
    
        'disable cbxBackups control
        ToggleControl Me.name, "cbxBackups"
        ToggleControl Me.name, "lblBackups", lngLtGray
    
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxIs_ODBC_BeforeUpdate[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub

End Sub

' ---------------------------------
' SUB:          cbxBackups_KeyDown
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, May, 2015 for NCPN invasives reporting tool
' Revisions:    BLC, 5/21/2015 - initial version
' ---------------------------------
Private Sub cbxBackups_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler

    Dim strMsgConfirm As String

    If Me.cbxBackups = False Then MsgBox "False", vbCritical, "Backups checked"

    If Me.ActiveControl = False Then
        If Me.cbxIs_ODBC = False Then
            strMsgConfirm = "Unchecking this box means that the application won't " & _
                vbCrLf & "make file backups for this database ... are you sure?"
            If MsgBox(strMsgConfirm, vbYesNo + vbDefaultButton2, _
                "Confirm change") = vbNo Then
                DoCmd.CancelEvent
            Else
                'disable cbxIs_ODBC control
                ToggleControl Me.name, "cbxIs_ODBC"
                ToggleControl Me.name, "lblIs_ODBC", lngLtGray
                Me.Requery
            End If
        End If
    ElseIf Me.cbxIs_ODBC Then
        MsgBox "This option does not apply to ODBC connections", , "File backups"
        DoCmd.CancelEvent
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxIs_ODBC_BeforeUpdate[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxBackups_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed chk prefix to cbx
'                                added toggle to turn off ODBC checkbox if Backups is selected
' ---------------------------------
Private Sub cbxBackups_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim strMsgConfirm As String

    If Me.ActiveControl = False Then
        If Me.cbxIs_ODBC = False Then
            strMsgConfirm = "Unchecking this box means that the application won't " & _
                vbCrLf & "make file backups for this database ... are you sure?"
            If MsgBox(strMsgConfirm, vbYesNo + vbDefaultButton2, _
                "Confirm change") = vbNo Then
                DoCmd.CancelEvent
                'disable cbxIs_ODBC control
                ToggleControl Me.name, "cbxIs_ODBC"
                ToggleControl Me.name, "lblIs_ODBC", lngLtGray
            Else
                'enable cbxIs_ODBC control
                ToggleControl Me.name, "cbxIs_ODBC"
                ToggleControl Me.name, "lblIs_ODBC", lngBlack
            End If
        End If
    Else
        'disable cbxIs_ODBC control
        ToggleControl Me.name, "cbxIs_ODBC"
        ToggleControl Me.name, "lblIs_ODBC", lngLtGray
'    ElseIf Me.cbxIs_ODBC = True Then
'        MsgBox "This option does not apply to ODBC connections", , "File backups"
'        DoCmd.CancelEvent
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxBackups_BeforeUpdate[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxFile_path_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
' ---------------------------------
Private Sub tbxFile_path_Click()
    On Error GoTo Err_Handler

    ' Use the zoom feature to display long strings
    SendKeys "+{F2}"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxFile_path_Click[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxNew_path_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
' ---------------------------------
Private Sub tbxNew_path_Click()
    On Error GoTo Err_Handler

    ' Use the zoom feature to display long strings
    SendKeys "+{F2}"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNew_path_Click[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnBrowse_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed cmd prefix to btn
' ---------------------------------
Private Sub btnBrowse_Click()
    On Error GoTo Err_Handler

    ' Browse to a new back-end file (Access only)
    Dim strCurrentDir As String
    Dim varFilePath As Variant
    Dim strNewDb As String
    Dim varResponse As VbMsgBoxResult

    ' Verify that it is an Access back-end and that the file path is present
    If Me.cbxIs_ODBC = True Then
        MsgBox "This option is only valid for non-ODBC connections", , _
            "Cannot browse to an ODBC database"
        Me.btnTestODBC.SetFocus
        GoTo Exit_Sub
    ElseIf IsNull(Me.tbxFile_path) Then
        ' Use an empty string as the initial directory if it is somehow missing
        strCurrentDir = ""
    Else
        ' Clip the path to indicate just the folder of the current back-end
        strCurrentDir = ParsePath(Me!tbxFile_path)
    End If

Get_new_file:
    ' Select the file, and start the search in the current back-end folder
    varFilePath = GetFile(strCurrentDir, "Microsoft Access (*.mdb, *.accdb)", _
        "*.mdb;*.accdb", "Connect to Back-end File")

    ' Exit if the user didn't specify a file
    If IsNull(varFilePath) Then GoTo Exit_Sub

    ' Verify that file actually exists, warn the user and loop back if not
    If FileExists(varFilePath) = False Then
        MsgBox "Unable to locate the file you selected:" & vbCrLf & vbCrLf & _
            varFilePath, , "Please try again"
        GoTo Get_new_file
    End If

    ' Update the new path and file name controls
    Me.tbxNew_path = varFilePath
    ' Update the new file name
    Me.tbxNew_db = ParseFileName(varFilePath)

    If Me.tbxNew_db <> Me.tbxLink_db Then
        ' Confirm if user is pointing to a db with a different name
        varResponse = MsgBox("Are you sure you indicated the correct back-end?" & _
            vbCrLf & vbCrLf & "Current db:" & vbTab & Me.tbxLink_db & vbCrLf & _
            "New db:" & vbTab & vbTab & Me.tbxNew_db, vbYesNo + vbDefaultButton2, _
            "Database name does not match")
        Select Case varResponse
          Case vbYes
            ' Check for duplicate database names and exit if it is a duplicate
            If DCount("*", "tsys_Link_dbs", "[Link_db]=""" & Me.tbxNew_db & """") > 0 Then
                MsgBox "There is already another linked database with the same name", , _
                    "Unable to link to this database"
                Me.tbxNew_path = Null
                Me.tbxNew_db = Null
                GoTo Exit_Sub
            End If
          Case Else
            Me.tbxNew_path = Null
            Me.tbxNew_db = Null
            GoTo Get_new_file
        End Select
    End If

    ' Enable the update button
    Me.btnUpdateLinks.Enabled = True

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3078   ' Can't find the system table
        MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
            vbCrLf & "the database administrator before using this application.", _
            vbCritical, "System Table " & _
            "Error encountered (#" & Err.Number & " - btnBrowse_Click[frm_Connect_Dbs])"
      Case 2001   ' Field name in DLookup improperly specified
        MsgBox "Error #" & Err.Number & ":  System table field not found." & _
            vbCrLf & "Please notify the database administrator before using " & _
            "this application.", vbCritical, "System Table " & _
            "Error encountered (#" & Err.Number & " - btnBrowse_Click[frm_Connect_Dbs])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnBrowse_Click[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxNew_server_Dirty
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
' ---------------------------------
Private Sub tbxNew_server_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Make sure it is an ODBC connection before allowing revisions to server name
    If Me.cbxIs_ODBC = False Then
        MsgBox "This option is only valid for ODBC connections", , "Not an ODBC database"
        Cancel = True
        Me.tbxNew_db.TabStop = False
    Else
        Me.tbxNew_db.TabStop = True
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNew_server_Dirty[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxNew_server_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
' ---------------------------------
Private Sub tbxNew_server_AfterUpdate()
On Error GoTo Err_Handler

    ' Enable the update button
    Me.btnUpdateLinks.Enabled = True

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNew_server_AfterUpdate[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:     tbxNew_db_Dirty
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
' ---------------------------------
Private Sub tbxNew_db_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Make sure it is an ODBC connection before allowing revisions to server name
    If Me.cbxIs_ODBC = False Then
        If IsNull(Me.tbxNew_path) Then
            MsgBox "This option is only valid for ODBC connections", , "Not an ODBC database"
            Cancel = True
        Else
            MsgBox "Use the Browse button to select the database", , "Change database name"
            Cancel = True
            Me.btnBrowse.SetFocus
        End If
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNew_db_Dirty[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:     tbxNew_db_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
' ---------------------------------
Private Sub tbxNew_db_AfterUpdate()
   On Error GoTo Err_Handler

    ' Enable the update button
    Me.btnUpdateLinks.Enabled = True

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNew_db_AfterUpdate[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub

End Sub

' ---------------------------------
' SUB:          btnTestODBC_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed cmd prefix to btn
' ---------------------------------
Private Sub btnTestODBC_Click()
    On Error GoTo Err_Handler

    Dim varResponse As VbMsgBoxResult
    Dim rst As DAO.Recordset
    Dim intNRecs As Integer
    Dim strCurrentDb As String
    Dim strTable As String
    Dim strNewDb As String
    Dim strNewServer As String
    Dim strNewConn As String

    ' Test the revised ODBC connections
    If Me.cbxIs_ODBC = False Then
        MsgBox "This option is only valid for ODBC connections", , "Not an ODBC database"
        GoTo Exit_Sub
    ElseIf IsNull(Me.tbxNew_server) Then
        MsgBox "Please indicate the new server name", , "Server name missing"
        Me.tbxNew_server.SetFocus
        GoTo Exit_Sub
    ElseIf IsNull(Me.tbxNew_db) Then
        ' See if user wants to use the same server
        varResponse = MsgBox("Do you want to use the current database name?" & vbCrLf & _
            vbCrLf & vbTab & Me.tbxLink_db, vbYesNo, "Confirm database name")
        Select Case varResponse
          Case vbYes
            Me.tbxNew_db = Me.tbxLink_db
          Case Else
            Me.tbxNew_db.SetFocus
            GoTo Exit_Sub
        End Select
    End If

    ' Check for valid linked table information first
    If VerifyLinkTableInfo = False Then GoTo Exit_Sub

    ' Set the recordset to the systems table, where objects are ODBC-linked tables
    '   associated with current database name
    Set rst = CurrentDb.OpenRecordset("SELECT Name, Connect, " & _
        "ParseConnectionStr([Connect]) AS CurrDb " & _
        "FROM MSysObjects " & _
        "WHERE ((MSysObjects.Name) Not Like '~*') AND " & _
        "((MSysObjects.Type) = 4) AND " & _
        "((ParseConnectionStr([Connect]) = """ & _
        Me.tbxLink_db & """));", _
        dbOpenSnapshot)

    ' Counts the number of ODBC-linked tables in the database
    rst.MoveLast    ' Need to do this to make the record count accurate
    intNRecs = rst.RecordCount
    If intNRecs = 0 Then    ' No records in the recordset
        MsgBox "There are no ODBC-linked tables to test ...", , _
            "Unable to test/update this connection"
        GoTo Exit_Sub
    End If

    ' Loop through the recordset and test each table if it matches the current db name
    rst.MoveFirst
    Do Until rst.EOF
        strNewConn = rst![Connect]
        ' Add the ODBC prefix to the connection string if not present
        If Left(strNewConn, 5) <> "ODBC;" Then
            strNewConn = "ODBC;" & strNewConn
        End If
        strTable = rst![name]
        strNewDb = Trim(Me.tbxNew_db)
        strNewConn = ReplaceString(strNewConn, rst![CurrDb], strNewDb)
        strNewServer = Trim(Me.tbxNew_server)
        strNewConn = ReplaceString(strNewConn, Me.tbxServer, strNewServer)
        If TestODBCConnection(strTable, strNewConn, , False) = False Then
            MsgBox "The following table connection failed:" & vbCrLf & vbCrLf & _
                vbTab & strTable, , "Test connection failed"
            GoTo Exit_Sub
        End If
        rst.MoveNext
    Loop

    MsgBox "Test connection succeeded!", , "Success"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnTestODBC_Click[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub

End Sub

' ---------------------------------
' SUB:          btnUpdateLinks_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/31/2014 - changed gvars to TempVars
'               BLC, 9/5/2014  - set added Is_Network_Db value for linked Dbs
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed cmd prefix to btn
'               BLC, 5/28/2015 - fixed bug resulting in DSN dialog when calling RefreshLinks
'                                strComponent was being passed in as False due to missing
'                                comma after "strNewConnStr," new call is
'                                RefreshLinks(strDbName, strNewConnStr, , False)
'                                added frm_Main_Menu restore on exit
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub btnUpdateLinks_Click()
    On Error GoTo Err_Handler

    Dim rst As DAO.Recordset
    Dim strDbName As String
    Dim strTable As String
    Dim strServer As String
    Dim strNewDb As String          ' New name for the linked database
    Dim strNewPath As String        ' New path to the linked database (Access)
    Dim strNewServer As String      ' New server name for the linked database (ODBC)
    Dim strNewConnStr As String     ' New connection string to update tables with
    Dim bHasError As Boolean

    ' Set a loop in case of multiple back-ends.  If errors are encountered on one,
    '   go to the next loop rather than exit
    Set rst = Me.Recordset
    rst.MoveFirst

    bHasError = False       ' Default until an error is encountered
    TempVars("HasAccessBE") = False

    ' If it is open, close the always-open form (used to maintain the connection to the
    '   back-end and avoid unnecessary create/delete/updates to its .ldb lock file)
    If FormIsLoaded("frm_Lock_BE") Then
        DoCmd.Close acForm, "frm_Lock_BE", acSaveNo
    End If

    Do Until rst.EOF
        strDbName = rst.Fields("Link_db")
        If rst.Fields("Is_ODBC") = True Then
            ' ODBC connection
            ' Move to next back-end if server or db name are blank
            If IsNull(rst.Fields("New_server")) Or IsNull(rst.Fields("New_db")) Then _
                GoTo NextBackEnd
            strNewDb = rst.Fields("New_db")
            strServer = rst.Fields("Server")
            strNewServer = rst.Fields("New_server")
            ' Make sure at least 1 table is associated with this back-end
            If DCount("*", "tsys_Link_Tables", "[Link_db]=""" & strDbName & """") = 0 Then
                MsgBox "There are no linked tables associated with this database", _
                    vbExclamation, strDbName
                bHasError = True
                GoTo NextBackEnd
            End If
            ' Get the first table in the list for this back-end
            strTable = DFirst("[Link_table]", "tsys_Link_Tables", _
                "[Link_db]=""" & strDbName & """")
            ' Start with the current connection string
            strNewConnStr = CurrentDb.tabledefs(strTable).Connect
            ' Update the connection string with the new server and db name
            strNewConnStr = ReplaceString(strNewConnStr, strDbName, strNewDb)
            strNewConnStr = ReplaceString(strNewConnStr, strServer, strNewServer)
            ' Update the links to the selected database
            If RefreshLinks(strDbName, strNewConnStr, True) = False Then
                ' A linking error was encountered
                MsgBox "Links to this database were not updated or only partially updated", _
                    vbExclamation, strDbName
                bHasError = True
                GoTo NextBackEnd
            End If
        Else
            ' Access back-end
            TempVars("HasAccessBE") = True
            ' If the user didn't specify a different database,
            '   refresh the links to the current linked file
            If IsNull(rst.Fields("New_path")) Then
                strNewPath = rst.Fields("File_path")
            Else
                strNewPath = rst.Fields("New_path")
            End If
            strNewConnStr = ";DATABASE=" & strNewPath
            ' Update the links to the selected database
            If RefreshLinks(strDbName, strNewConnStr, , False) = False Then
                ' A linking error was encountered
                MsgBox "Links to this database were not updated or only partially updated", _
                    vbExclamation, strDbName
                bHasError = True
                GoTo NextBackEnd
            End If
        End If

        ' Move to next back end without updating the record if no new info was entered
        If IsNull(rst.Fields("New_db")) Then GoTo NextBackEnd
        ' If no error on this back end then update this record with the new info
        With rst
            .Edit
            !Link_db = rst.Fields("New_db").Value
            !File_path = rst.Fields("New_path").Value
            !Server = rst.Fields("New_server").Value
            !New_db = Null
            !New_path = Null
            !New_server = Null
            '!Is_Network_Db = IsNetworkFile(rst.Fields("New_path").Value) ' <<<<< ERROR 94 TRIGGERED HERE
            .Update
            .Bookmark = .lastModified
        End With

NextBackEnd:
        On Error Resume Next
        If Err > 0 Then
            MsgBox "Error #" & Err.Number & ": " & Err.Description, _
                vbCritical, "Error encountered while updating database links"
            bHasError = True
        End If
        Err = 0
        rst.MoveNext
    Loop
    ' End the loop accommodating multiple back-end files here

    ' If no connection errors, then notify the user and reset the application
    If bHasError = False Then
        ' Verify and clean up the link table info
        If VerifyLinkTableInfo = True Then
            ' If no problems persist ...
            MsgBox "Update complete!", vbExclamation, "Back-end database connections"
            TempVars("Connected") = True
            DoCmd.Close , , acSaveNo
            ' Call the function to set up the application using new connection info
            Call AppSetup
        End If
    End If

Exit_Sub:
    ' restore main menu
    DoCmd.SelectObject acForm, Forms(MAIN_APP_MENU), False
    DoCmd.Restore
    
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3078   ' Can't find the system table
        MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
            vbCrLf & "the database administrator before using this application.", _
            vbCritical, "Application error" & _
            "(#" & Err.Number & " - btnUpdateLinks_Click[frm_Connect_Dbs])"
      Case 3265   ' Field name in the system table improperly specified
        MsgBox "Error #" & Err.Number & ":  System table field not found." & _
            vbCrLf & "Please notify the database administrator before using " & _
            "this application.", vbCritical, "Application error" & _
            "(#" & Err.Number & " - btnUpdateLinks_Click[frm_Connect_Dbs])"
      Case 94    ' Missing information in the system table
        MsgBox "Error #" & Err.Number & ":  Missing system table info. Please notify" & _
            vbCrLf & "the database administrator before using this application.", _
            vbCritical, "Application error" & _
            "(#" & Err.Number & " - btnUpdateLinks_Click[frm_Connect_Dbs])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered while updating database links " & _
            "(#" & Err.Number & " - btnUpdateLinks_Click[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub

End Sub

' ---------------------------------
' SUB:          btnClose_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed cmd prefix to btn,
'                                added DB_ADMIN_CONTROL check for partial DbAdmin form implementations
' ---------------------------------
Private Sub btnClose_Click()
    On Error GoTo Err_Handler

    If Me.btnUpdateLinks.Enabled Then
        If MsgBox("Close without updating links?", vbOKCancel, _
            "Back-end links not updated") = vbCancel Then Exit Sub
    End If

    Dim rst As DAO.Recordset

    ' Set the recordset to the form's underlying table
    Set rst = Me.Recordset
    rst.MoveFirst
    ' Loop through records and blank out the new file name & path before closing
    Do Until rst.EOF
        With rst
            .Edit
            !New_db = Null
            !New_path = Null
            !New_server = Null
            .Update
            .Bookmark = .lastModified
        End With
        rst.MoveNext
    Loop

    If DB_ADMIN_CONTROL Then
        ' Requery the control that shows the linked back-ends
        Forms!frm_Switchboard!lbxLinkedDbs.Requery
    End If
    
Exit_Sub:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    DoCmd.Close , , acSaveNo
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Connect_Dbs])"
    End Select
    Resume Exit_Sub
End Sub
