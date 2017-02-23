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
    ItemSuffix =111
    Left =3855
    Top =2430
    Right =23490
    Bottom =15015
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xfe27f5b1c5c3e440
    End
    RecordSource ="SELECT tsys_Link_Dbs.* FROM tsys_Link_Dbs ORDER BY tsys_Link_Dbs.[SortOrder], ts"
        "ys_Link_Dbs.[LinkDb]; "
    Caption =" Update Database Connections"
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
                    Left =2700
                    Top =120
                    Width =2814
                    Height =252
                    ColumnWidth =3090
                    FontSize =9
                    FontWeight =700
                    ForeColor =16711680
                    Name ="tbxLinkDb"
                    ControlSource ="LinkDb"
                    StatusBarText ="Linked database name"
                    FontName ="Arial"
                    ControlTipText ="Linked database name"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =120
                    LayoutCachedWidth =5514
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =93
                    Left =1680
                    Top =480
                    Width =7917
                    Height =252
                    ColumnWidth =6630
                    FontSize =9
                    TabIndex =3
                    Name ="tbxDb_desc"
                    ControlSource ="DbDesc"
                    StatusBarText ="Brief description of the type of database - e.g., data tables, lookup tables, sy"
                        "stems tables, etc."
                    FontName ="Arial"
                    ControlTipText ="Brief description of the type of database - e.g., data tables, lookup tables, sy"
                        "stems tables, etc."

                    LayoutCachedLeft =1680
                    LayoutCachedTop =480
                    LayoutCachedWidth =9597
                    LayoutCachedHeight =732
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =720
                            Top =480
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="lblDbDesc"
                            Caption ="Description:"
                            FontName ="Arial"
                            LayoutCachedLeft =720
                            LayoutCachedTop =480
                            LayoutCachedWidth =1605
                            LayoutCachedHeight =720
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
                    Name ="tbxFilePath"
                    ControlSource ="FilePath"
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
                    Name ="tbxNewPath"
                    ControlSource ="NewPath"
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
                    Name ="tbxNewDb"
                    ControlSource ="NewDb"
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
                    Name ="cbxIsODBC"
                    ControlSource ="IsODBC"
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
                            Name ="lblIsODBC"
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
                    Name ="tbxNewServer"
                    ControlSource ="NewServer"
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
                    Name ="tbxSortOrder"
                    ControlSource ="SortOrder"
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
                        0x0100000080000000010000000100000000000000000000000e00000001010000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490073004f004400420043005d003d005400720075006500000000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ffffff000d0000005b00 ,
                        0x490073004f004400420043005d003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
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
                    Name ="lblNewServer"
                    ControlSource ="='New server (ODBC only):'"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x0100000080000000010000000100000000000000000000000e00000001010000 ,
                        0x0000ff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490073004f004400420043005d003d005400720075006500000000000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010100000000ff00ffffff000d0000005b00 ,
                        0x490073004f004400420043005d003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
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
                    Name ="lblNewDb"
                    ControlSource ="='New db name (ODBC only):'"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x0100000080000000010000000100000000000000000000000e00000001010000 ,
                        0x0000ff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490073004f004400420043005d003d005400720075006500000000000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010100000000ff00ffffff000d0000005b00 ,
                        0x490073004f004400420043005d003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
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
                    Name ="lblFilePath"
                    ControlSource ="='Path:'"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x0100000082000000010000000100000000000000000000000f00000001010000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490073004f004400420043005d003d00460061006c007300650000000000 ,
                        0x0000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ffffff000e0000005b00 ,
                        0x490073004f004400420043005d003d00460061006c0073006500000000000000 ,
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
                    Left =660
                    Top =1260
                    Width =960
                    Height =255
                    FontSize =9
                    TabIndex =8
                    Name ="lblNewPath"
                    ControlSource ="='New path:'"
                    FontName ="Arial"
                    ConditionalFormat = Begin
                        0x0100000082000000010000000100000000000000000000000f00000001010000 ,
                        0x0000ff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490073004f004400420043005d003d00460061006c007300650000000000 ,
                        0x0000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010100000000ff00ffffff000e0000005b00 ,
                        0x490073004f004400420043005d003d00460061006c0073006500000000000000 ,
                        0x000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =0
                    BorderWidth =3
                    OverlapFlags =247
                    BackStyle =0
                    Left =180
                    Top =120
                    Width =2394
                    Height =252
                    FontSize =9
                    FontWeight =700
                    TabIndex =17
                    Name ="tbxLinkType"
                    ControlSource ="=[LinkType] & \":\""
                    StatusBarText ="Linked database name"
                    FontName ="Arial"
                    ControlTipText ="Linked database name"

                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =2574
                    LayoutCachedHeight =372
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
' FORM NAME:    ConnectDbs
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
'               BLC, 12/3/2015 - modified btnUpdateLinks_Click to handle updating database names for
'                    the same type w/ a different file name
'               BLC, 6/4/2016 - adapted for Big Rivers Application & re-named to ConnectDbs
'               BLC, 2/22/2017 - added CallingForm property & Form_Close event
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_CallingForm As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidCallingForm(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let CallingForm(Value As String)
    If Len(Value) > 0 Then
        m_CallingForm = Value
    Else
        RaiseEvent InvalidCallingForm(Value)
    End If
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'---------------------
' Methods
'---------------------

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
'               BLC, 6/5/2016 - removed underscores from field names
'               BLC, 6/20/2016 - revised from "frm_Switchboard" to MAIN_APP_FORM
'               BLC, 6/24/2016 - minimized Main form
'               BLC, 2/22/2017 - added CallingForm property
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "DbAdmin"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize calling form
    ToggleForm Me.CallingForm, -1

    If DB_ADMIN_CONTROL Then
        ' Close the form if the switchboard is not open
        If FormIsOpen(MAIN_APP_FORM) = False Then
            MsgBox "The main database switchboard must be" & vbCrLf & _
                "open for this form to function properly.", , "Cannot open the form ..."
            DoCmd.CancelEvent
            GoTo Exit_Sub
        End If
    End If

    ' Enable/disable subforms depending on user access level
    Select Case TempVars("UserAccessLevel")
      Case "admin"
        Me.cbxIsODBC.Locked = False
      Case Else
        Me.cbxIsODBC.Locked = True
    End Select

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[ConnectDbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxIsODBC_KeyDown
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, May, 2015 for NCPN invasives reporting tool
' Revisions:    BLC, 5/21/2015 - initial version
'               BLC, 6/5/2016 - removed underscores from field names
' ---------------------------------
Private Sub cbxIsODBC_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler

    Dim strMsgConfirm As String

    If cbxIsODBC = False Then MsgBox "false", vbCritical, "Is ODBC checked"

    strMsgConfirm = "To change the connection type, you must delete and relink" & vbCrLf & _
        "the tables for this database manually ... are you sure?"
    If MsgBox(strMsgConfirm, vbOKCancel, "Confirm change") = vbCancel Then
        'vbCancel = True
        Me.ActiveControl.Undo
    Else
        'disable cbxBackups control
        ToggleControl Me.Name, "cbxBackups"
        ToggleControl Me.Name, "lblBackups", lngLtGray
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxIsODBC_KeyDown[ConnectDbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxIsODBC_BeforeUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed chk prefix to cbx
'               BLC, 6/5/2016 - removed underscores from field names
' ---------------------------------
Private Sub cbxIsODBC_BeforeUpdate(Cancel As Integer)
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
        ToggleControl Me.Name, "cbxBackups"
        ToggleControl Me.Name, "lblBackups", lngBlack
    
    Else
    
    'ODBC (true)
    
        'disable cbxBackups control
        ToggleControl Me.Name, "cbxBackups"
        ToggleControl Me.Name, "lblBackups", lngLtGray
    
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxIsODBC_BeforeUpdate[ConnectDbs])"
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
'               BLC, 6/5/2016 - removed underscores from field names
' ---------------------------------
Private Sub cbxBackups_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler

    Dim strMsgConfirm As String

    If Me.cbxBackups = False Then MsgBox "False", vbCritical, "Backups checked"

    If Me.ActiveControl = False Then
        If Me.cbxIsODBC = False Then
            strMsgConfirm = "Unchecking this box means that the application won't " & _
                vbCrLf & "make file backups for this database ... are you sure?"
            If MsgBox(strMsgConfirm, vbYesNo + vbDefaultButton2, _
                "Confirm change") = vbNo Then
                DoCmd.CancelEvent
            Else
                'disable cbxIs_ODBC control
                ToggleControl Me.Name, "cbxIsODBC"
                ToggleControl Me.Name, "lblIsODBC", lngLtGray
                Me.Requery
            End If
        End If
    ElseIf Me.cbxIsODBC Then
        MsgBox "This option does not apply to ODBC connections", , "File backups"
        DoCmd.CancelEvent
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxIsODBC_BeforeUpdate[ConnectDbs])"
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
'               BLC, 6/5/2016 - removed underscores from field names
' ---------------------------------
Private Sub cbxBackups_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim strMsgConfirm As String

    If Me.ActiveControl = False Then
        If Me.cbxIsODBC = False Then
            strMsgConfirm = "Unchecking this box means that the application won't " & _
                vbCrLf & "make file backups for this database ... are you sure?"
            If MsgBox(strMsgConfirm, vbYesNo + vbDefaultButton2, _
                "Confirm change") = vbNo Then
                DoCmd.CancelEvent
                'disable cbxIs_ODBC control
                ToggleControl Me.Name, "cbxIsODBC"
                ToggleControl Me.Name, "lblIsODBC", lngLtGray
            Else
                'enable cbxIs_ODBC control
                ToggleControl Me.Name, "cbxIsODBC"
                ToggleControl Me.Name, "lblIsODBC", lngBlack
            End If
        End If
    Else
        'disable cbxIs_ODBC control
        ToggleControl Me.Name, "cbxIsODBC"
        ToggleControl Me.Name, "lblIsODBC", lngLtGray
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
            "Error encountered (#" & Err.Number & " - cbxBackups_BeforeUpdate[ConnectDbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxFilePath_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
'               BLC, 6/5/2016 - removed underscores from field names
' ---------------------------------
Private Sub tbxFilePath_Click()
    On Error GoTo Err_Handler

    ' Use the zoom feature to display long strings
    SendKeys "+{F2}"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxFilePath_Click[ConnectDbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxNewPath_Click
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
'               BLC, 6/5/2016 - removed underscores from field names
' ---------------------------------
Private Sub tbxNewPath_Click()
    On Error GoTo Err_Handler

    ' Use the zoom feature to display long strings
    SendKeys "+{F2}"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNewPath_Click[ConnectDbs])"
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
'               BLC, 6/5/2016 - removed underscores from field names
' ---------------------------------
Private Sub btnBrowse_Click()
    On Error GoTo Err_Handler

    ' Browse to a new back-end file (Access only)
    Dim strCurrentDir As String
    Dim varFilePath As Variant
    Dim strNewDb As String
    Dim varResponse As VbMsgBoxResult

    ' Verify that it is an Access back-end and that the file path is present
    If Me.cbxIsODBC = True Then
        MsgBox "This option is only valid for non-ODBC connections", , _
            "Cannot browse to an ODBC database"
        Me.btnTestODBC.SetFocus
        GoTo Exit_Sub
    ElseIf IsNull(Me.tbxFilePath) Then
        ' Use an empty string as the initial directory if it is somehow missing
        strCurrentDir = ""
    Else
        ' Clip the path to indicate just the folder of the current back-end
        strCurrentDir = ParsePath(Me!tbxFilePath)
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
    Me.tbxNewPath = varFilePath
    ' Update the new file name
    Me.tbxNewDb = ParseFileName(varFilePath)

    If Me.tbxNewDb <> Me.tbxLinkDb Then
        ' Confirm if user is pointing to a db with a different name
        varResponse = MsgBox("Are you sure you indicated the correct back-end?" & _
            vbCrLf & vbCrLf & "Current db:" & vbTab & Me.tbxLinkDb & vbCrLf & _
            "New db:" & vbTab & vbTab & Me.tbxNewDb, vbYesNo + vbDefaultButton2, _
            "Database name does not match")
        Select Case varResponse
          Case vbYes
            ' Check for duplicate database names and exit if it is a duplicate
            If DCount("*", "tsys_Link_Dbs", "[LinkDb]=""" & Me.tbxNewDb & """") > 0 Then
                MsgBox "There is already another linked database with the same name", , _
                    "Unable to link to this database"
                Me.tbxNewPath = Null
                Me.tbxNewDb = Null
                GoTo Exit_Sub
            End If
          Case Else
            Me.tbxNewPath = Null
            Me.tbxNewDb = Null
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
            "Error encountered (#" & Err.Number & " - btnBrowse_Click[ConnectDbs])"
      Case 2001   ' Field name in DLookup improperly specified
        MsgBox "Error #" & Err.Number & ":  System table field not found." & _
            vbCrLf & "Please notify the database administrator before using " & _
            "this application.", vbCritical, "System Table " & _
            "Error encountered (#" & Err.Number & " - btnBrowse_Click[ConnectDbs])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnBrowse_Click[ConnectDbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxNewServer_Dirty
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
'               BLC, 6/5/2016 - removed underscores from field names
' ---------------------------------
Private Sub tbxNewServer_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Make sure it is an ODBC connection before allowing revisions to server name
    If Me.cbxIsODBC = False Then
        MsgBox "This option is only valid for ODBC connections", , "Not an ODBC database"
        Cancel = True
        Me.tbxNewDb.TabStop = False
    Else
        Me.tbxNewDb.TabStop = True
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNewServer_Dirty[ConnectDbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxNewServer_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
'               BLC, 6/5/2016 - removed underscores from field names
' ---------------------------------
Private Sub tbxNewServer_AfterUpdate()
On Error GoTo Err_Handler

    ' Enable the update button
    Me.btnUpdateLinks.Enabled = True

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNewServer_AfterUpdate[ConnectDbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:     tbxNewDb_Dirty
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
'               BLC, 6/5/2016 - removed underscores from field names
' ---------------------------------
Private Sub tbxNewDb_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Make sure it is an ODBC connection before allowing revisions to server name
    If Me.cbxIsODBC = False Then
        If IsNull(Me.tbxNewPath) Then
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
            "Error encountered (#" & Err.Number & " - tbxNewDb_Dirty[ConnectDbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:     tbxNewDb_AfterUpdate
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 6/13/2014 - added to WQ utilities tool
'               BLC, 5/21/2015 - added to invasives reporting tool & renamed txt prefix to tbx
'               BLC, 6/5/2016 - removed underscores from field names
' ---------------------------------
Private Sub tbxNewDb_AfterUpdate()
   On Error GoTo Err_Handler

    ' Enable the update button
    Me.btnUpdateLinks.Enabled = True

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNewDb_AfterUpdate[ConnectDbs])"
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
'               BLC, 6/5/2016 - removed underscores from field names & use GetDbTemplates()
' ---------------------------------
Private Sub btnTestODBC_Click()
    On Error GoTo Err_Handler

    Dim varResponse As VbMsgBoxResult
    Dim rs As DAO.Recordset
    Dim intNRecs As Integer
    Dim strCurrentDb As String
    Dim strTable As String
    Dim strNewDb As String
    Dim strNewServer As String
    Dim strNewConn As String

    ' Test the revised ODBC connections
    If Me.cbxIsODBC = False Then
        MsgBox "This option is only valid for ODBC connections", , "Not an ODBC database"
        GoTo Exit_Sub
    ElseIf IsNull(Me.tbxNewServer) Then
        MsgBox "Please indicate the new server name", , "Server name missing"
        Me.tbxNewServer.SetFocus
        GoTo Exit_Sub
    ElseIf IsNull(Me.tbxNewDb) Then
        ' See if user wants to use the same server
        varResponse = MsgBox("Do you want to use the current database name?" & vbCrLf & _
            vbCrLf & vbTab & Me.tbxLinkDb, vbYesNo, "Confirm database name")
        Select Case varResponse
          Case vbYes
            Me.tbxNewDb = Me.tbxLinkDb
          Case Else
            Me.tbxNewDb.SetFocus
            GoTo Exit_Sub
        End Select
    End If

    ' Check for valid linked table information first
    If VerifyLinkTableInfo = False Then GoTo Exit_Sub

    ' Set the recordset to the systems table, where objects are ODBC-linked tables
    '   associated with current database name
'    Set rs = CurrentDb.OpenRecordset("SELECT Name, Connect, " & _
'        "ParseConnectionStr([Connect]) AS CurrDb " & _
'        "FROM MSysObjects " & _
'        "WHERE ((MSysObjects.Name) Not Like '~*') AND " & _
'        "((MSysObjects.Type) = 4) AND " & _
'        "((ParseConnectionStr([Connect]) = """ & _
'        Me.tbxLink_db & """));", _
'        dbOpenSnapshot)

    Set rs = CurrentDb.OpenRecordset(GetTemplate("s_msysobjects_connect", "LinkDb" & PARAM_SEPARATOR & Me.tbxLinkDb), dbOpenSnapshot)

    ' Counts the number of ODBC-linked tables in the database
    rs.MoveLast    ' Need to do this to make the record count accurate
    intNRecs = rs.RecordCount
    If intNRecs = 0 Then    ' No records in the recordset
        MsgBox "There are no ODBC-linked tables to test ...", , _
            "Unable to test/update this connection"
        GoTo Exit_Sub
    End If

    ' Loop through the recordset and test each table if it matches the current db name
    rs.MoveFirst
    Do Until rs.EOF
        strNewConn = rs![Connect]
        ' Add the ODBC prefix to the connection string if not present
        If Left(strNewConn, 5) <> "ODBC;" Then
            strNewConn = "ODBC;" & strNewConn
        End If
        strTable = rs![Name]
        strNewDb = Trim(Me.tbxNewDb)
        strNewConn = ReplaceString(strNewConn, rs![CurrDb], strNewDb)
        strNewServer = Trim(Me.tbxNewServer)
        strNewConn = ReplaceString(strNewConn, Me.tbxServer, strNewServer)
        If TestODBCConnection(strTable, strNewConn, , False) = False Then
            MsgBox "The following table connection failed:" & vbCrLf & vbCrLf & _
                vbTab & strTable, , "Test connection failed"
            GoTo Exit_Sub
        End If
        rs.MoveNext
    Loop

    MsgBox "Test connection succeeded!", , "Success"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnTestODBC_Click[ConnectDbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnUpdateLinks_Click
' Description:  Update new linked database tsys_Link_Dbs, tsys_Link_Files, tsys_Link_Tables
' Assumptions:  Tables (tsys_Link_Dbs, tsys_Link_Files & tsys_Link_Dbs) exist with fields as noted
'               Linked database file types are set for a given application based on its needs.
'               Link_type is the primary key (and unchangeable) for tsys_Link_Dbs & tsys_Link_Files.
'               tsys_Link_Files is present for backward compatibility in some applications.
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
'               BLC, 12/2/2015 - revised to handle new database names
'               BLC, 6/4/2016  - adapted to big rivers app & updated frm_Lock_BE to LockBE
'               BLC, 6/5/2016 - removed underscores from field names & use GetTemplate()
'                               close MAIN_APP_MENU form only when open
' ---------------------------------
Private Sub btnUpdateLinks_Click()
    On Error GoTo Err_Handler

    Dim rs As DAO.Recordset
    Dim strDbName As String
    Dim strDbType As String         ' Database type (e.g. back-end database)
    Dim strTable As String
    Dim strServer As String
    Dim strNewDb As String          ' New name for the linked database
    Dim strNewPath As String        ' New path to the linked database (Access)
    Dim strNewServer As String      ' New server name for the linked database (ODBC)
    Dim strNewConnStr As String     ' New connection string to update tables with
    Dim bHasError As Boolean, bIsODBC As Boolean
    Dim strSQL As String

    ' -------------------------------
    ' Handle multiple back-ends --> On error, next loop rather than exit
    ' -------------------------------
    Set rs = Me.Recordset      'backend: tsys_Link_Dbs
    rs.MoveFirst

    bIsODBC = False       ' Default
    bHasError = False       ' Default until an error is encountered
    TempVars("HasAccessBE") = False

    ' If open, close always-open form (used to maintain back-end connection &
    '   avoid unnecessary create/delete/updates to its .ldb lock file)
    If FormIsLoaded("LockBE") Then
        DoCmd.Close acForm, "LockBE", acSaveNo
    End If

' --- multiple back-end files loop ---
    Do Until rs.EOF
        strDbName = rs.Fields("LinkDb")
        strDbType = rs.Fields("LinkType")

    '---------------------
    ' ODBC Connected Back-ends
    '---------------------
        If rs.Fields("IsODBC") = True Then
            
            ' Server or db name are blank? --> Move to next back-end
            If IsNull(rs.Fields("NewServer")) Or IsNull(rs.Fields("NewDb")) Then _
                GoTo NextBackEnd
            
            strNewDb = rs.Fields("New_db")
            strServer = rs.Fields("Server")
            strNewServer = rs.Fields("New_server")
            
            ' Check for associated tables (must be @ least 1 for this back-end)
            If DCount("*", "tsys_Link_Tables", "[LinkDb]=""" & strDbName & """") = 0 Then
                MsgBox "There are no linked tables associated with this database", _
                    vbExclamation, strDbName
                bHasError = True
                GoTo NextBackEnd
            End If

            ' Get the first table in the list for this back-end
            strTable = DFirst("[LinkTable]", "tsys_Link_Tables", _
                "[LinkDb]=""" & strDbName & """")
            
            ' Start with the current connection string
            strNewConnStr = CurrentDb.TableDefs(strTable).Connect

            ' Update connection string with new server & db name
            strNewConnStr = ReplaceString(strNewConnStr, strDbName, strNewDb)
            strNewConnStr = ReplaceString(strNewConnStr, strServer, strNewServer)

            ' Update selected database links
            bIsODBC = True

    '---------------------
    ' Access Connected Back-ends
    '---------------------
        Else

            TempVars("HasAccessBE") = True

'            ' Different database? --> if so, go to next backend
'            ' -----------------------------
'            If (rs.Fields("Link_db") <> strNewDb) and Then GoTo NextBackEnd

            ' Same database? --> refresh links to current linked file
            ' -----------------------------
            If IsNull(rs.Fields("NewPath")) Then
                strNewPath = rs.Fields("FilePath")
                strNewDb = rs.Fields("LinkDb")
            Else
                strNewPath = rs.Fields("NewPath")
                strNewDb = rs.Fields("NewDb")
            End If
            strNewConnStr = ";DATABASE=" & strNewPath

            ' Verify file & update links to it
            bIsODBC = False

        End If

    '---------------------
    ' ODBC or Access Connected Back-ends
    '---------------------
        ' Update selected database links
        If RefreshLinks(strDbName, strNewConnStr, , bIsODBC) = False Then  '(strDbName, strNewConnStr, , bIsODBC, strNewDb) = False Then
            ' A linking error was encountered
            MsgBox "Links to this database were not updated or only partially updated", _
                vbExclamation, strDbName
            bHasError = True
            GoTo NextBackEnd
        End If

        ' Move to next back end without updating the record if no new info was entered
        If IsNull(rs.Fields("NewDb")) Then GoTo NextBackEnd
'-------------------------
        'No Linking Errors on this back end & new file path --> update current path and file
'--- ADD? ----------------
'        ElseIf IsNull(rs.Fields("New_db")) = False Then
'-------------------------

        With rs
            .Edit
            !LinkDb = rs.Fields("NewDb").Value
            !FilePath = rs.Fields("NewPath").Value
            !Server = rs.Fields("NewServer").Value
            !NewDb = Null
            !NewPath = Null
            !NewServer = Null
            '!Is_Network_Db = IsNetworkFile(rs.Fields("New_path").Value) ' <<<<< ERROR 94 TRIGGERED HERE
            .Update
            .Bookmark = .LastModified
        End With
        
        ' update tsys_Link_Dbs & tsys_Link_Files database name & paths
        DoCmd.SetWarnings False 'hide the append dialog
        
'        strSQL = "UPDATE tsys_Link_Files SET Link_File_Path = '" & strNewPath & "', " & _
'                 "Link_file_name = '" & strNewDb & "' " & _
'                 "WHERE Link_file_name = '" & strDbName & "';"
        strSQL = GetTemplate("u_tsys_Link_Files_new_db", "NewPath" & PARAM_SEPARATOR & strNewPath & "|NewDb" & PARAM_SEPARATOR & strNewDb & "|DbName" & PARAM_SEPARATOR & strNewDb)
        DoCmd.RunSQL strSQL
        DoCmd.SetWarnings True

NextBackEnd:
        On Error Resume Next
        If Err > 0 Then
            MsgBox "Error #" & Err.Number & ": " & Err.Description, _
                vbCritical, "Error encountered while updating database links"
            bHasError = True
        End If
        Err = 0
        rs.MoveNext
    Loop
' --- multiple back-end files loop ---

    ' If no connection errors --> notify the user and reset the application
    If bHasError = False Then
        ' Verify and clean up the link table info
        If VerifyLinkTableInfo = True Then
            ' If no problems persist ...
            MsgBox "Update complete!", vbExclamation, "Back-end database connections"
            TempVars("Connected") = True
            DoCmd.Close , , acSaveNo
            ' Call function to set up the application using new connection info
            Call AppSetup
        End If
    End If

Exit_Sub:
    ' restore main menu (if open)
    If FormIsOpen(MAIN_APP_MENU) Then
        DoCmd.SelectObject acForm, Forms(MAIN_APP_MENU), False
        DoCmd.Restore
    End If
    
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3078   ' Can't find the system table
        MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
            vbCrLf & "the database administrator before using this application.", _
            vbCritical, "Application error" & _
            "(#" & Err.Number & " - btnUpdateLinks_Click[ConnectDbs])"
      Case 3265   ' Field name in the system table improperly specified
        MsgBox "Error #" & Err.Number & ":  System table field not found." & _
            vbCrLf & "Please notify the database administrator before using " & _
            "this application.", vbCritical, "Application error" & _
            "(#" & Err.Number & " - btnUpdateLinks_Click[ConnectDbs])"
      Case 94    ' Missing information in the system table
        MsgBox "Error #" & Err.Number & ":  Missing system table info. Please notify" & _
            vbCrLf & "the database administrator before using this application.", _
            vbCritical, "Application error" & _
            "(#" & Err.Number & " - btnUpdateLinks_Click[ConnectDbs])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered while updating database links " & _
            "(#" & Err.Number & " - btnUpdateLinks_Click[ConnectDbs])"
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
'               BLC, 6/5/2016 - removed underscores from field names
'               BLC, 2/22/2017 - added CallingForm toggle
' ---------------------------------
Private Sub btnClose_Click()
    On Error GoTo Err_Handler

    If Me.btnUpdateLinks.Enabled Then
        If MsgBox("Close without updating links?", vbOKCancel, _
            "Back-end links not updated") = vbCancel Then Exit Sub
    End If

    Dim rs As DAO.Recordset

    ' Set the recordset to the form's underlying table
    Set rs = Me.Recordset
    rs.MoveFirst
    ' Loop through records and blank out the new file name & path before closing
    Do Until rs.EOF
        With rs
            .Edit
            !NewDb = Null
            !NewPath = Null
            !NewServer = Null
            .Update
            .Bookmark = .LastModified
        End With
        rs.MoveNext
    Loop

'    If DB_ADMIN_CONTROL Then
'        ' Requery the control that shows the linked back-ends
'        Forms!frm_Switchboard!lbxLinkedDbs.Requery
'    End If
    
Exit_Sub:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    DoCmd.Close , , acSaveNo
    
    'restore calling form
    ToggleForm Me.CallingForm, 0
    
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[ConnectDbs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' Sub:          Form_Close
' Description:  form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 22, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/22/2017 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore calling form
    ToggleForm Me.CallingForm, 0
    
    'set next action
    SetTempVar "GoToSplash", True
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[ConnectDbs])"
    End Select
    Resume Exit_Handler
End Sub
