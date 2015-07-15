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
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9300
    DatasheetFontHeight =10
    ItemSuffix =96
    Left =4770
    Top =2895
    Right =14070
    Bottom =8400
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb46db5e5f0f8e240
    End
    RecordSource ="tsys_Link_Files"
    Caption =" Update Data Table Connections"
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
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
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
            Height =960
            BackColor =10402763
            Name ="FormHeader"
            Begin
                Begin Label
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =2
                    Left =2580
                    Top =60
                    Width =3924
                    Height =276
                    FontSize =10
                    FontWeight =700
                    Name ="labTitle"
                    Caption ="Update links to back end database tables"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =420
                    Width =7500
                    Height =495
                    FontSize =9
                    Name ="labFormDesc"
                    Caption ="Data tables are stored in one or more separate database files.  Check the filena"
                        "me and location on your computer for the following and use the browse button to "
                        "change the file location(s)."
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7920
                    Top =60
                    Width =1263
                    FontSize =9
                    FontWeight =700
                    Name ="cmdClose"
                    Caption ="Close form"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =7920
                    Top =540
                    Width =1263
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdUpdateLinks"
                    Caption ="Update links"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Update links to the file(s) indicated"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2280
            BackColor =12902115
            Name ="Detail"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =7740
                    Top =1440
                    Width =888
                    Height =300
                    FontSize =9
                    FontWeight =700
                    Name ="cmdBrowse"
                    Caption ="Browse"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    BorderWidth =3
                    OverlapFlags =93
                    BackStyle =0
                    Left =180
                    Top =120
                    Width =2754
                    Height =252
                    ColumnWidth =3090
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="txtLinkType"
                    ControlSource ="Link_type"
                    StatusBarText ="code for the data source - NCBN tables, NER tables, etc."
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =3000
                    Top =120
                    Width =6057
                    Height =480
                    ColumnWidth =6630
                    FontSize =9
                    TabIndex =2
                    Name ="txtLinkDescription"
                    ControlSource ="Link_description"
                    StatusBarText ="Describes the types of data tables included in the link"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =1020
                    Top =720
                    Width =6183
                    ColumnWidth =2520
                    FontSize =9
                    TabIndex =3
                    Name ="txtCurrentName"
                    ControlSource ="Link_file_name"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =300
                            Top =540
                            Width =696
                            Height =444
                            FontSize =9
                            Name ="labCurrentName"
                            Caption ="Current name:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =95
                    Left =1020
                    Top =1080
                    Width =8037
                    ColumnWidth =2205
                    FontSize =9
                    TabIndex =4
                    Name ="txtCurrentPath"
                    ControlSource ="Link_file_path"
                    StatusBarText ="Current linked file path"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =480
                            Top =1080
                            Width =540
                            Height =240
                            FontSize =9
                            Name ="labCurrentPath"
                            Caption ="Path:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =114
                    Top =63
                    Width =9063
                    Height =2160
                    Name ="Box82"
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =1020
                    Top =1860
                    Width =8037
                    Height =252
                    FontSize =9
                    TabIndex =6
                    Name ="txtNewPath"
                    ControlSource ="New_file_path"
                    StatusBarText ="New linked file path"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =420
                            Top =1860
                            Width =540
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labNewPath"
                            Caption ="Path:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =1020
                    Top =1500
                    Width =6183
                    Height =252
                    FontSize =9
                    TabIndex =5
                    Name ="txtNewName"
                    ControlSource ="New_file_name"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =120
                            Top =1500
                            Width =840
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labNewName"
                            Caption ="New file:"
                            FontName ="Arial"
                        End
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
' FORM NAME:    frm_Connect_Tables
' Description:  Standard module to update back-end db connections
' Data source:  tsys_Link_Files
' Data access:  edit only, no additions, moving between records, or deletions
' Pages:        none
' Functions:    none
' References:   GetLinkFile, RefreshLinks, SwitchboardIsOpen
' Source/date:  Susan Huse, MonitoringSM.mdb v 7/28/2004
' Revisions:    John R. Boetsch, May 2005 - minor edits
' Revisions:    JRB, May 24, 2006 - documentation, added error trapping, fixed specification
'               of initial directory to current directory, simplified a little
' =================================

Private Sub txtCurrentPath_Click()
    On Error GoTo Err_Handler

    SendKeys "+{F2}"

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub txtNewPath_Click()
    On Error GoTo Err_Handler

    SendKeys "+{F2}"

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo Err_Handler

    Dim strCurrentFile As String
    Dim strCurrentDir As String
    Dim varFilePath As Variant
    Dim arrFile() As String

    strCurrentFile = Me!txtCurrentName
    strCurrentDir = Me!txtCurrentPath

    ' Clip to indicate just the folder of the current back-end
    strCurrentDir = Left(strCurrentDir, Len(strCurrentDir) - Len(strCurrentFile) - 1)

    ' Select the file, and start the search in the current back-end folder
    ' -------------------------------------------------------
    ' BLC, 5/19/2015 - revised to use GetFile vs GetLinkFile
    varFilePath = GetFile(strCurrentDir)
    ' -------------------------------------------------------
    ' Exit if the user didn't specify a file
    If IsNull(varFilePath) Then GoTo Exit_Procedure

    ' Update the new path and file name controls
    Me!txtNewPath = varFilePath
    ' Update the new file name after first storing the path components in an array
    arrFile = Split(varFilePath, "\")
    Me!txtNewName = arrFile(UBound(arrFile))
    Me!cmdUpdateLinks.Enabled = True

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case 3078   ' Can't find the system table
            MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_Link_Files)"
            Resume Exit_Procedure
        Case 2001   ' Field name in DLookup improperly specified
            MsgBox "Error #" & Err.Number & ":  System table field not found." & _
                vbCrLf & "Please notify the database administrator before using " & _
                "this application.", vbCritical, "System table error (tsys_Link_Files)"
            Resume Exit_Procedure
        Case 94    ' Missing information in the systems table
            MsgBox "Error #" & Err.Number & ":  Missing system table info. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_Link_Files)"
            Resume Exit_Procedure
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
            Resume Exit_Procedure
    End Select
    
End Sub

Private Sub cmdUpdateLinks_Click()
    On Error GoTo Err_Handler

    Dim rst As DAO.Recordset
    Dim strSysTable As String       ' Name of the system table listing linked tables
    Dim strLinkName As String
    Dim strFilePath As String       ' Path of the new database
    Dim strSQL As String
    Dim bHasError As Boolean
    Dim strLinkDb As String         ' Name of linked database
    Dim strConnString As String     ' Linked db connection string

    strSysTable = "[tsys_Link_Tables]"  ' Set the name of the system table

    ' Set a loop in case of multiple back-ends.  If errors are encountered on one,
    '   go to the next loop rather than exit
    Set rst = Me.Recordset
    rst.MoveFirst

    bHasError = False   ' Default until an error is encountered

    Do Until rst.EOF
        strLinkDb = rst.Fields("Link_file_name")
        strLinkName = rst.Fields("Link_type")
        ' If the user didn't specify a different database,
        '   refresh the links to the current linked file
        If IsNull(rst.Fields("New_file_path")) Then
            strFilePath = rst.Fields("Link_file_path")
        Else
            strFilePath = rst.Fields("New_file_path")
        End If

        ' Build a query statement identifying the tables that should be in the file
        strSQL = "SELECT * FROM " & strSysTable & " WHERE " & _
            strSysTable & "![Link_type] = '" & strLinkName & "'"

        ' Verify the file and update the links to the selected file
    ' -------------------------------------------------------
    '   BLC, 5/20/2015 - updated to new RefreshLinks version
        
        'If RefreshLinks(strSQL, strFilePath) = False Then

        'ODBC; DATABASE=database; UID=user; PWD=password; DSN= datasourcename;
        'strConnString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};" & _
        '            "Dbq=" & strFilePath & ";Uid=Admin;Pwd=;"
        strConnString = ";DATABASE=" & strFilePath

                    
        If RefreshLinks(strLinkDb, strConnString) = False Then
    ' -------------------------------------------------------
            ' An error was encountered
            MsgBox "Links to this file were not updated or only partially updated", _
                vbExclamation, strLinkName
            bHasError = True
            GoTo NextBackEnd
        ' If no linking error on this back end then update the current path and file
        ElseIf IsNull(rst.Fields("New_file_path")) = False Then
            With rst
                .Edit
                !Link_file_name = rst.Fields("New_file_name").Value
                !Link_file_path = rst.Fields("New_file_path").Value
                !New_file_name = Null
                !New_file_path = Null
                .Update
                .Bookmark = .lastModified
            End With
            
    ' -------------------------------------------------------
    '   BLC, 5/20/2015 - added update for tsys_Link_Dbs paths
    
            'update tsys_Link_Dbs paths
            strSQL = "UPDATE tsys_Link_Dbs SET File_Path = '" & strFilePath & "' " & _
                     "WHERE Link_db = '" & strLinkDb & "';"
            DoCmd.SetWarnings False 'hide the append dialog
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
    ' -------------------------------------------------------
        End If

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

    ' If no connection errors, then notify the user and close
    If bHasError = False Then
        MsgBox "Update complete!", vbExclamation, "Update Back-end Data Connections"
        DoCmd.Close , , acSaveNo
    End If

Exit_Procedure:
    ' -------------------------------------------------------
    '   BLC, 5/28/2015 - added frm_Main_Menu restore
            DoCmd.SelectObject acForm, Forms(MAIN_APP_MENU), False
            DoCmd.Restore
    ' -------------------------------------------------------
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case 3078   ' Can't find the system table
            MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_Link_Tables)"
        Case 3265   ' Field name in the tsys_Link_Files improperly specified
            MsgBox "Error #" & Err.Number & ":  System table field not found." & _
                vbCrLf & "Please notify the database administrator before using " & _
                "this application.", vbCritical, "System table error (tsys_Link_Tables)"
        Case 94    ' Missing information in the systems table
            MsgBox "Error #" & Err.Number & ":  Missing system table info. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_Link_Tables)"
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, _
                vbCritical, "Error encountered while updating database links"
    End Select
    Resume Exit_Procedure
    
End Sub

Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

DoCmd.Close acForm, Me.name, acSaveNo
'clear new file name, new file path
CurrentDb.Execute "UPDATE tsys_Link_Files SET New_file_name=null, New_file_path=null;"

Exit_Procedure:
    On Error Resume Next
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
