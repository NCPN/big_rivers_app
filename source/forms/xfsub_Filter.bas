Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4140
    DatasheetFontHeight =9
    ItemSuffix =50
    Left =6615
    Top =7275
    Right =10755
    Bottom =10140
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1385341e7574e340
    End
    Caption =" "
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ToggleButton
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackColor =-1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-1
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverColor =0
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedColor =0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeColor =0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeColor =0
            PressedForeThemeColorIndex =1
        End
        Begin Section
            Height =2880
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =975
                    Left =1200
                    Top =1200
                    Width =1740
                    Height =360
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="lbxYear"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Visit_Year.Visit_Year FROM qry_Visit_Year ORDER BY qry_Visit_Year.Vis"
                        "it_Year DESC; "
                    ColumnWidths ="975"
                    StatusBarText ="Choose the year (10/1-9/31/XXXX)"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Choose the year (10/1-9/31/XXXX)"

                    LayoutCachedLeft =1200
                    LayoutCachedTop =1200
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1560
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1200
                            Width =465
                            Height =240
                            FontWeight =700
                            Name ="lblYear"
                            Caption ="Year"
                            LayoutCachedLeft =120
                            LayoutCachedTop =1200
                            LayoutCachedWidth =585
                            LayoutCachedHeight =1440
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2280
                    Top =2100
                    Width =1740
                    Height =600
                    FontSize =12
                    FontWeight =600
                    TabIndex =2
                    ForeColor =-2147483617
                    Name ="btnRun"
                    Caption ="RUN >>"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
                    VerticalAnchor =1

                    LayoutCachedLeft =2280
                    LayoutCachedTop =2100
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =2700
                    UseTheme =1
                    Bevel =-1
                    Gradient =25
                    BackColor =5880731
                    BackThemeColorIndex =6
                    BorderColor =5880731
                    BorderThemeColorIndex =6
                    HoverColor =4763790
                    HoverThemeColorIndex =6
                    HoverShade =90.0
                    PressedColor =4234622
                    PressedThemeColorIndex =6
                    PressedShade =80.0
                    HoverForeColor =9974127
                    PressedForeColor =62207
                    Shadow =-1
                    QuickStyle =39
                    QuickStyleMask =-513
                    WebImagePaddingTop =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =7200
                    Left =1200
                    Top =660
                    Width =2820
                    Height =360
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"16\""
                    Name ="lbxProjectID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [tblProjects].[ProjectID], [tblProjects].[ProjectName] FROM tblProjects O"
                        "RDER BY [ProjectID]; "
                    ColumnWidths ="1008;6192"
                    StatusBarText ="Choose the desired project"
                    ControlTipText ="Choose the desired project"

                    LayoutCachedLeft =1200
                    LayoutCachedTop =660
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =1020
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =780
                            Width =690
                            Height =240
                            FontWeight =700
                            Name ="lblProjectID"
                            Caption ="Project"
                            LayoutCachedLeft =120
                            LayoutCachedTop =780
                            LayoutCachedWidth =810
                            LayoutCachedHeight =1020
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
' FORM NAME:    fsub_Filter
' Description:  Subform - provides values to queries & reports run from switchboard form
' Data source:  various
' Data access:  read only
' Pages:        -
' Functions:    none
' Subroutines:  -
' References:   -
' Source/date:  Bonnie L. Campbell, June 12, 2014
' Adapted/date: -
' Revisions:    BLC, 6/xx/2014 - XX
' =================================

' ---------------------------------
' SUB:     btnRun_Click
' Description:  Prepares parameters & runs query or report
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Adapted from John Boetsch
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - XX
' ---------------------------------
Private Sub btnRun_Click()
    Dim dictParams As New Dictionary
    
    'pass selected values to query/report
    Select Case TempVars.item("analysis")
    'analysis
        Case "Outliers", "MissingData", "Duplicates"
        
            dictParams.Add "ProjectID", lbxProjectID.ListIndex
            dictParams.Add "Project", lbxProjectID.Value
            dictParams.Add "Year", lbxYear.ListIndex
            dictParams.Add "qry", "qry_" & TempVars.item("analysis")

            RunReport dictParams

        'suspect values
        Case "SuspectValues", "SuspectDO", "SuspectpH", "SuspectSC", "SuspectWT"
        
            dictParams.Add "ProjectID", lbxProjectID.ListIndex
            dictParams.Add "Project", lbxProjectID.Value
            dictParams.Add "Year", lbxYear.ListIndex
            dictParams.Add "qry", "qry_" & TempVars.item("analysis")

            RunReport dictParams
        
        'reports
        Case "Precision", _
             "Effectiveness", _
             "Bias", _
             "Stage", _
             "Flow" ' Representativeness > Stage & Flow
            
            dictParams.Add "ProjectID", lbxProjectID.ListIndex
            dictParams.Add "Project", lbxProjectID.Value
            dictParams.Add "Year", lbxYear.ListIndex
            dictParams.Add "qry", "qry_" & TempVars.item("analysis")

            RunReport dictParams
            
    'export
    End Select
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    'prepare breadcrumbs
    'MsgBox Me.OpenArgs, vbCritical, "crumbs"
    'Dim aryCrumbs As Variant
    
    'aryCrumbs = fxnCrumbsToArray(Me.OpenArgs)

    
End Sub

' ---------------------------------
' SUB:          RunReport
' Description:  Runs query or report based on parameters (site, year)
' Assumptions:  -
' Parameters:   TempVars.Item("params") - parameters as a string
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, June 17, 2014 - for NCPN WQ Utilities tool
' Revisions:    6/17/2014 - BLC - XX
' ---------------------------------
Public Sub RunReport(varParams As Variant)
On Error GoTo Err_Handler:

    Dim Response As String

    'pass selected values to query/report
    Select Case TempVars.item("analysis")
        
    ' --------------
    ' analysis
    ' --------------
    ' - field data -
        Case "Outliers"
            varParams("qry") = "qry_Outlier_Count"
            
        Case "MissingData"
            'qry = Missing vs. MissingData
            varParams("qry") = "qry_Missing"

        'suspect values
        Case "SuspectValues"
            'qry = Suspect vs. SuspectValues
            varParams("qry") = "qry_Suspect"
        Case "SuspectDO"
        Case "SuspectpH"
        Case "SuspectSC"
        Case "SuspectWT"
    
    ' - lab data -
        Case "Duplicates"
            'qry = Lab_Duplicates vs. Duplicates
            varParams("qry") = "qry_Lab_Duplicates"
    
    ' --------------
    ' reports
    ' --------------
        Case "Precision"
        Case "Effectiveness"
        Case "Bias"
        
        Case "Stage" ' Representativeness > Stage
            varParams("qry") = "qry_Representativeness_Stage"
            
        Case "Flow" ' Representativeness > Flow
            varParams("qry") = "qry_Representativeness_Flow"
    ' --------------
    ' exports
    ' --------------
        Case "UtahLab"
        
    End Select
        
    'prepare output
    Dim strQuery As String
    
    If Len(varParams("qry")) > 0 Then
        strQuery = varParams("qry")
    Else
        GoTo Not_Found
    End If
    
    Select Case TempVars.item("action")
    ' --------------
    ' analysis (queries)
    ' --------------
        Case "qry"
            'handle Suspect_XX queries
            'strQuery = Replace(strQuery, "Suspect_", "Suspect")
            If Len(Replace(strQuery, "Suspect", "")) > 4 Then
                strQuery = Replace(strQuery, "Suspect", "Suspect_")
            End If
            
    ' --------------
    ' reports
    ' --------------
        Case "rpt"
            
    ' --------------
    ' exports
    ' --------------
        Case "exp"

    End Select

' --------------
'  Run Report
' --------------
    'process query
    If qryExists(strQuery) Then
        DoCmd.SetWarnings False
        DoCmd.Hourglass True
        SysCmd acSysCmdSetStatus, "Running " & strQuery & " query... "
    
        ' Get basic result info
        Dim rstQuery As DAO.Recordset
        
        Set rstQuery = dbCurrent.OpenRecordset(strQuery, dbOpenDynaset)
        
        ' clear statusbar
        SysCmd acSysCmdSetStatus, "Calculations complete!"
        
        If rstQuery.EOF Then
            MsgBox "Sorry, no valid " & TempVars.item("analysis") & " records were found" & vbCrLf & vbCrLf & _
             "for the park(s)/year(s) selected when running the query.", vbOKOnly, _
             "No Records for " & TempVars.item("analysis") & " Analysis"
            GoTo Exit_Procedure
        Else
            'present user with the choice of viewing the records or not
             Response = MsgBox("Finished calculating!" & vbCrLf & vbCrLf & _
                               "Do you want to view your results in the " & _
                               vbCrLf & vbCrLf & TempVars.item("analysis") & " query?" & _
                               vbCrLf & vbCrLf & "If not, they'll be there until you run this calculation again.", _
                               vbYesNo, StrConv(TempVars.item("analysis"), vbProperCase) & " Complete!")
             
             If Response = vbYes Then    ' User chose Yes.
                'open the query
                DoCmd.OpenQuery strQuery, acViewNormal, acEdit
             Else    ' User chose No.
                'do nothing
             End If
        
        End If

    Else
        GoTo Not_Found
    End If

    'cleanup
    Set rstQuery = Nothing

Exit_Procedure:
    DoCmd.Hourglass False
    SysCmd acSysCmdSetStatus, " "
    DoCmd.SetWarnings True
    Exit Sub
    
Not_Found:
    MsgBox "Sorry, the query could not be found.", vbCritical, "Missing Query"
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RunReport[Form_fsub_Filter])"
    End Select
    GoTo Exit_Procedure
End Sub

' ---------------------------------
' SUB:          Form_Close
' Description:  Cleanup
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, June 17, 2014
' Revisions:    6/17/2014 - BLC - XX
' ---------------------------------
Public Sub Form_Close()
    MsgBox "close"

    'clear controls
    With Me
        .Controls.item("lbxProjectID").ListIndex = 0
        .Controls.item("lbxYear").ListIndex = 0
        .Parent.Controls.item("tbxInstructions").Caption = " "
    End With
    
End Sub
