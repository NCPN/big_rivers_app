Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4140
    DatasheetFontHeight =11
    ItemSuffix =16
    Left =1320
    Top =1890
    Right =6210
    Bottom =6000
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x9f832d99b891e440
    End
    RecordSource ="tbl_Target_Species"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    SplitFormSplitterBar =0
    SplitFormSplitterBar =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =2880
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =1740
                    Top =2280
                    Width =2220
                    ForeColor =16711680
                    Name ="btnContinue"
                    Caption ="Continue >>"
                    StatusBarText ="Continue to choose activities"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1740
                    LayoutCachedTop =2280
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =2640
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =6750156
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =52377
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =3407769
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =52224
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =2375487
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =6750156
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ListBox
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =300
                    Top =480
                    Width =2880
                    Height =1620
                    ColumnOrder =0
                    TabIndex =1
                    BoundColumn =2
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxTgtLists"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tbl_Target_List.Park_Code, tbl_Target_List.Target_Year, tbl_Targ"
                        "et_List.Park_Code & \"-\" & tbl_Target_List.Target_Year AS ParkYear FROM tbl_Tar"
                        "get_List ORDER BY tbl_Target_List.[Park_Code], tbl_Target_List.[Target_Year];"
                    ColumnWidths ="0;0;1440"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =480
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =2100
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =60
                            Width =1176
                            Height =314
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblParkYear"
                            Caption ="Target Lists"
                            GridlineColor =10921638
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =1236
                            LayoutCachedHeight =374
                        End
                    End
                End
            End
        End
        Begin Section
            Height =0
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' FORM:         Form_fsub_Select_Tgt_Lists
' Description:  Choose target list(s) for reporting functions & procedures
'
' Source/date:  Bonnie Campbell, 5/1/2015
' Revisions:    BLC - 5/1/2015 - initial version
'               BLC - 6/12/2015 - added Continue button enable,
'                                 replaced TempVars.item("... with TempVars("...
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  Actions for form loading
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 1, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/1/2015 - initial version
'   BLC - 6/12/2015 - disabled Continue button to start
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
    
    Initialize
    
    'disable continue to start
    btnContinue.Enabled = False
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[form_fsub_Select_Tgt_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtLists_Click
' Description:  Determine selected target list(s)
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/1/2015 - initial version
'   BLC - 6/12/2015 - added logic to enable Continue button
' ---------------------------------
Private Sub lbxTgtLists_Click()
On Error GoTo Err_Handler
Dim strTgtLists As String, strComma As String
Dim item As Variant

    'determine the selected list(s)
    For Each item In lbxTgtLists.ItemsSelected
        
        strTgtLists = strTgtLists & "'" & lbxTgtLists.ItemData(item) & "',"

    Next
    
    'trim last comma
    strTgtLists = IIf(Right(strTgtLists, 1) = ",", Left(strTgtLists, Len(strTgtLists) - 1), strTgtLists)
    
    TempVars.Add "TgtLists", strTgtLists
    
    'enable Continue button
    If Len(strTgtLists) > 0 Then
        btnContinue.Enabled = True
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtLists_Click[form_fsub_Select_Tgt_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnContinue_Click
' Description:  Continue to report displays
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC, 5/1/2015 - initial version
'   BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub btnContinue_Click()
On Error GoTo Err_Handler
    Dim strReport As String, strWhere As String
    
    Select Case TempVars("rpt")
        
        Case "CrewSpeciesList" ' Reports > Field Crew Species List
            strReport = "rpt_Tgt_Species_List"
            strWhere = "TgtList IN (" & TempVars("TgtLists") & ")"
        
        Case "SpeciesListByPark" ' Reports > Species List By Park
            strReport = "rpt_Tgt_Species_List_By_Park"
            strWhere = "TgtList IN (" & TempVars("TgtLists") & ")"
        
        Case "TgtListAnnualSummary" ' Reports > Annual Species List Summary
            strReport = "rpt_Tgt_Species_List_Annual_Summary"
            strWhere = ""
            
    End Select
    
    'open target species list
    DoCmd.OpenReport strReport, acViewReport, , strWhere
        
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnContinue_Click[form_fsub_Select_Tgt_Lists])"
    End Select
    Resume Exit_Sub
End Sub
