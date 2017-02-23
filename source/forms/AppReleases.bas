Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =8100
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =3975
    Top =3510
    Right =22110
    Bottom =13815
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1719458615c5e440
    End
    RecordSource ="SELECT tsys_App_Releases.* FROM tsys_App_Releases ORDER BY tsys_App_Releases.Rel"
        "ease_date, tsys_AppReleases.VersionNumber; "
    Caption =" Application Releases"
    AfterUpdate ="[Event Procedure]"
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
        Begin Section
            CanGrow = NotDefault
            Height =9000
            BackColor =9677753
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1560
                    Top =120
                    Width =5340
                    Height =252
                    ColumnWidth =1440
                    Name ="tbxReleaseID"
                    ControlSource ="ID"
                    StatusBarText ="Unique identifier for the release"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1032
                            Height =252
                            FontWeight =700
                            Name ="lblReleaseID"
                            Caption ="Release ID"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =1200
                    Width =1680
                    Height =252
                    ColumnWidth =1140
                    TabIndex =4
                    Name ="tbxRelease_date"
                    ControlSource ="ReleaseDate"
                    Format ="Short Date"
                    StatusBarText ="Date of the release"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1200
                            Width =1185
                            Height =270
                            FontWeight =700
                            Name ="lblReleaseDate"
                            Caption ="Release date"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =840
                    Width =1557
                    Height =252
                    ColumnWidth =972
                    TabIndex =2
                    Name ="tbxVersionNumber"
                    ControlSource ="VersionNumber"
                    Format ="General Number"
                    StatusBarText ="Version control number"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =1452
                            Height =252
                            FontWeight =700
                            Name ="lblVersionNumber"
                            Caption ="Version number"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =1560
                    Width =3300
                    Height =252
                    ColumnWidth =2568
                    TabIndex =6
                    Name ="tbxFileName"
                    ControlSource ="FileName"
                    StatusBarText ="Filename, used to identify older versions of the database"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =924
                            Height =252
                            FontWeight =700
                            Name ="lblFileName"
                            Caption ="File name"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3600
                    Left =4440
                    Top =1200
                    Width =2106
                    Height =252
                    ColumnWidth =2568
                    TabIndex =5
                    Name ="cbxReleaseBy"
                    ControlSource ="ReleaseBy"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Project_Crew.Contact_ID, IIf([Contact_is_active],'Active','') AS Is_a"
                        "ctive FROM tlu_Project_Crew ORDER BY IIf([Contact_is_active],'Active','') DESC ,"
                        " tlu_Project_Crew.Contact_ID; "
                    ColumnWidths ="2448;1152"
                    StatusBarText ="Person who made the release"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3420
                            Top =1200
                            Width =1044
                            Height =252
                            FontWeight =700
                            Name ="lblReleaseBy"
                            Caption ="Release by"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    AllowAutoCorrect = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =1920
                    Width =6480
                    Height =2400
                    ColumnWidth =3000
                    TabIndex =7
                    Name ="tbxReleaseNotes"
                    ControlSource ="ReleaseNotes"
                    StatusBarText ="Release notes, which may include a summary of revisions"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1920
                            Width =1332
                            Height =252
                            FontWeight =700
                            Name ="lblReleaseNotes"
                            Caption ="Release notes"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =480
                    Width =6480
                    ColumnWidth =2568
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"200\""
                    Name ="btnDatabaseTitle"
                    ControlSource ="DatabaseTitle"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tsys_App_Releases.DatabaseTitle FROM tsys_App_Releases GROUP BY tsys_App_"
                        "Releases.DatabaseTitle ORDER BY tsys_App_Releases.DatabaseTitle; "
                    StatusBarText ="Title of the database"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1272
                            Height =252
                            FontWeight =700
                            Name ="lblDatabaseTitle"
                            Caption ="Database title"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Subform
                    Locked = NotDefault
                    OverlapFlags =215
                    Left =120
                    Top =4500
                    Width =7920
                    Height =4500
                    TabIndex =8
                    Name ="Bugs"
                    SourceObject ="Form.BugReports"
                    LinkChildFields ="Release_ID"
                    LinkMasterFields ="Release_ID"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =4260
                            Width =1116
                            Height =252
                            FontWeight =700
                            Name ="lblBugs"
                            Caption ="Known bugs"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5100
                    Top =840
                    Height =270
                    TabIndex =3
                    Name ="cbxIsSupported"
                    ControlSource ="IsSupported"
                    RowSourceType ="Value List"
                    RowSource ="0;Not supported;1;Supported;2;Current"
                    ColumnWidths ="0;1728"
                    StatusBarText ="Indicates the level of support for this version"
                    DefaultValue ="2"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3240
                            Top =840
                            Width =1785
                            Height =255
                            FontWeight =700
                            Name ="lblIsSupported"
                            Caption ="Current/supported?"
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
' FORM NAME:    AppReleases
' Level:        Framework form
' Version:      1.04
'
' Description:  Standard form for viewing and entering release information
' Data source:  In-line SQL statement based on tsys_App_Releases
'   --> SELECT tsys_App_Releases.* FROM tsys_App_Releases ORDER BY tsys_App_Releases.ReleaseDate, tsys_App_Releases.VersionNumber;
' Data access:  edits, add, no deletions
' Pages:        none
' Functions:    none
' References:   none
' Source/date:  John R. Boetsch, September 2008
' Adapted:      Bonnie Campbell, June 2014
' Revisions:    JRB, 9/26/2008 - 1.00 - updated cmbIs_supported to allow 3 values instead of true/false
'               JRB, 10/6/2008 - 1.01 - updated to unlock subform if in admin mode
'               BLC, 6/12/2014 - 1.02 - Revised to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 6/12/2016 - 1.03 - Adapted to big rivers
'               BLC, 6/24/2016 - 1.04 - updated error handling, form minimize/restore
' =================================

' ---------------------------------
' SUB:     Form_Open
' Description: Opens sub form & sets controls based on UserAccessLevel
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Created John R. Boetsch, September 2008
' Adapted:      Bonnie Campbell, June 2014
' Revisions:    BLC, 6/12/2014 - Revised to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 8/5/2014 - changed to use setUserAccess for initializing control settings based on app mode
'               BLC, 6/13/2016 - adapted for big rivers
'               BLC, 6/24/2016 - updated error handling, added DbAdmin minimize
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'minimize DbAdmin
    ToggleForm "DbAdmin", -1
    
    DoCmd.GoToRecord , , acLast
    If SwitchboardIsOpen Then
            'initialize controls based on app mode
            setUserAccess Me
    End If
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[AppReleases form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:     Form_AfterUpdate
' Description: Requeries form after updates
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Created John R. Boetsch, September 2008
' Adapted:      Bonnie Campbell, June 2014
' Revisions:    BLC, 6/12/2014 - Revised to use TempVars.Item("UserAccessLevel") vs. cAppMode
'               BLC, 6/13/2016 - adapted for big rivers
'               BLC, 6/24/2016 - updated error handling
' ---------------------------------
Private Sub Form_AfterUpdate()
On Error GoTo Err_Handler

    If SwitchboardIsOpen Then Forms!frm_Switchboard!cmbVersion.Requery

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_AfterUpdate[AppReleases form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Close
' Description:  Closes form
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, June 24, 2016
' Adapted:      -
' Revisions:    BLC, 6/24/2014 - initial
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore DbAdmin
    ToggleForm "DbAdmin", 0

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[AppReleases form])"
    End Select
    Resume Exit_Handler
End Sub
