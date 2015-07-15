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
    Width =3960
    DatasheetFontHeight =11
    ItemSuffix =14
    Left =7350
    Top =5430
    Right =11490
    Bottom =8295
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x2cb45f3fbd91e440
    End
    RecordSource ="qry_Park_Tgt_Species_Lists"
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
            Height =1920
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =1140
                    Width =2220
                    ForeColor =16711680
                    Name ="btnContinue"
                    Caption ="Continue >>"
                    StatusBarText ="Continue to choose activities"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =1500
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
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =480
                    Top =540
                    Width =1980
                    Height =300
                    ColumnOrder =0
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="cbxYear"
                    RowSourceType ="Value List"
                    RowSource ="'SEL';'Select Year';'2017';'2017';'2015';'2015';'2014';'2014';'2013';'2013';'201"
                        "2';'2012';'2011';'2011';'2010';'2010';'2009';'2009';'2008';'2008';'';'';"
                    ColumnWidths ="0;1440"
                    DefaultValue ="\"SEL\""
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =480
                    LayoutCachedTop =540
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Width =1176
                    Height =314
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear"
                    Caption ="Year"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =1296
                    LayoutCachedHeight =434
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
' FORM:         Form_fsub_Select_Year
' Description:  Target species functions & procedures
'
' Source/date:  Bonnie Campbell, 5/1/2015
' Revisions:    BLC - 5/1/2015 - initial version
'               BLC - 6/12/2015 - added Continue button enable,
'                                 replaced TempVars.item("... with TempVars("...
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  actions for select year form load
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/1/2015 - initial version
'   BLC - 6/12/2015 - disabled Continue button to start
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String, strValueList As String
    Dim i As Integer, count As Integer
    
    Initialize
    
    'prepare value list
    strSQL = "SELECT DISTINCT TgtYear FROM qry_Park_Tgt_Species_Lists ORDER BY TgtYear DESC;"
    
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)

    strValueList = "'SEL';'Select Year';"

    If Not rs.BOF And Not rs.EOF Then
        rs.MoveLast
        count = rs.RecordCount
        rs.MoveFirst
        For i = 0 To count - 1
            strValueList = strValueList & "'" & rs("TgtYear") & "';'" & rs("TgtYear") & "';"
            rs.MoveNext
        Next
    End If
    
    cbxYear.RowSource = strValueList
    cbxYear.Value = "SEL"
    
    'disable continue to start
    btnContinue.Enabled = False
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[form_fsub_Select_Year])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxYear_Change
' Description:  Actions to take when a task action is selected
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 1, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/1/2015 - initial version
'   BLC - 6/12/2015 - added enable Continue button when valid year value is selected,
'                     replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub cbxYear_Change()
On Error GoTo Err_Handler

    If Len(Trim(cbxYear)) > 0 Then
        'set year
        TempVars("TgtYear") = cbxYear.Value
        'enable continue
        btnContinue.Enabled = True
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxYear_Change[form_fsub_Select_Year])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnContinue_Click
' Description:  Continue to park actions
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/12/2015 - initial version
'   BLC, 5/1/2015  - switched from frmActions to launching popup frm_Tgt_Species form for Invasive Species Reporting tool
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'                     added catch for Error #94 Invalid Use of Null which occasionally
'                     happens w/ debugging, TempVars("TgtYear") is somehow lost
' ---------------------------------
Private Sub btnContinue_Click()
On Error GoTo Err_Handler
       
    TempVars("TgtYear") = cbxYear.Value
    
    If TempVars("TgtYear") > 0 Then
    
        'open report
        DoCmd.OpenReport "rpt_Tgt_Species_List_Annual_Summary", acViewReport, , "TgtYear=" & CInt(TempVars("TgtYear"))
        
    End If
        
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 94 'Invalid Use of NULL
        MsgBox "Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
            "Re-select the year you desire. I've somehow forgotten it." & vbCrLf & vbCrLf & _
            "Selected Target Year: " & TempVars("TgtYear"), vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[form_fsub_Select_Year])"
        Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnContinue_Click[form_fsub_Select_Year])"
    End Select
    Resume Exit_Sub
End Sub
