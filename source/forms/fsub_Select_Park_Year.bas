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
    Width =4320
    DatasheetFontHeight =11
    ItemSuffix =15
    Left =10104
    Top =5004
    Right =14256
    Bottom =7620
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xc1f3db6ed487e440
    End
    RecordSource ="tbl_Target_Areas"
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
                    Left =1800
                    Top =1440
                    Width =2220
                    ForeColor =16711680
                    Name ="btnContinue"
                    Caption ="Continue >>"
                    StatusBarText ="Continue to choose activities"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1800
                    LayoutCachedTop =1440
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =1800
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
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5040
                    Left =300
                    Top =360
                    Width =2160
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"10\""
                    Name ="cbxPark"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.ParkCode, tlu_Parks.ParkName FROM tlu_Parks WHERE tlu_Parks.Par"
                        "kCode IN ('BLCA','CARE','COLM','CURE','DINO','FOBU','GOSP','ZION') ORDER BY tlu_"
                        "Parks.[ParkName];"
                    ColumnWidths ="1080;3960"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="Choose a park."
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =360
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =660
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =300
                    Top =960
                    Width =2220
                    Height =300
                    ColumnOrder =0
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="cbxYear"
                    RowSourceType ="Value List"
                    RowSource ="'SEL';'Select Year';'2017';'2017';'2016';'2016';'2015';'2015';'2014';'2014';'201"
                        "3';'2013';'2012';'2012';"
                    ColumnWidths ="0;1440"
                    DefaultValue ="\"SEL\""
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =960
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =1260
                End
                Begin Label
                    OverlapFlags =247
                    Left =60
                    Top =60
                    Width =1176
                    Height =314
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPark"
                    Caption ="Park"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1236
                    LayoutCachedHeight =374
                End
                Begin Label
                    OverlapFlags =247
                    Left =60
                    Top =660
                    Width =1176
                    Height =314
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear"
                    Caption ="Year"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =660
                    LayoutCachedWidth =1236
                    LayoutCachedHeight =974
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
' FORM:         Form_fsub_Select_Park_Year
' Description:  Target species functions & procedures
'
' Source/date:  Bonnie Campbell, 2/11/2015
' Revisions:    BLC - 2/11/2015 - initial version
'               BLC - 6/12/2015 - added Continue button enable,
'                                 replaced TempVars.item("... with TempVars("...
'               BLC - 7/7/2015  - investigated bug causing debugger to open on clicking btnContinue
'                                 for *some* park/year combos (DINO-2015, BLCA-2016, CARE-2016, COLM-2016)
'                                 appears related more to IDE debug error handling (Tools>Options>General)?
'                                 setting to "Break in Class Module" then back to "Break on Unhandled Errors"
'                                 seems to fix? no subroutine code changes made
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
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
'   BLC - 6/12/2015 - disabled Continue button to start
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
    Dim i As Integer, iYear As Integer
    Dim strValueList As String
    
    Initialize
    
    'prepare value list
    strValueList = "'SEL';'Select Year';"

    iYear = Year(Now()) + 2

    For i = 1 To 6
        strValueList = strValueList & "'" & iYear & "';'" & iYear & "';"
        iYear = iYear - 1
    Next
    
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
            "Error encountered (#" & Err.Number & " - Form_Load[form_fsub_Select_Park_Year])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxPark_Change
' Description:  Actions to take when a park is selected
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
'   BLC - 6/12/2015 - changed the check from 0 to 3 (Park_Code = 4 characters) and
'                     added enabling continue button, changed TempVars.item("... to TempVars("...
' ---------------------------------
Private Sub cbxPark_Change()
On Error GoTo Err_Handler
    
    'set park & enable continue when a 4-letter park code is selected
    If Len(cbxPark.Value) > 3 Then
        'set park
        TempVars("park") = Trim(cbxPark.Value)
        
        'enable the continue button
        If Len(cbxPark) > 3 And TempVars("TgtYear") > 0 Then
            btnContinue.Enabled = True
        End If
    End If
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPark_Change[form_fsub_Select_Park_Year])"
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
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015 - initial version
'   BLC - 6/12/2015 - added enabling continue button, changed TempVars.item("... to TempVars("...
' ---------------------------------
Private Sub cbxYear_Change()
On Error GoTo Err_Handler

    If Len(Trim(cbxYear)) > 0 Then
        'set year
        TempVars("TgtYear") = cbxYear.Value
        
        'enable the continue button
        If Len(cbxPark) > 3 And TempVars("TgtYear") > 0 Then
            btnContinue.Enabled = True
        End If
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxYear_Change[form_fsub_Select_Park_Year])"
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
'   BLC, 5/10/2015 - cleared park & year cbx values to prevent NULL errors & force user to re-select park before clicking continue
'   BLC, 7/7/2015  - investigated bug causing debugger to open on clicking btnContinue for *some* park/year combos
'                    reported combos were: DINO-2015, BLCA-2016, CARE-2016, COLM-2016
'                    appears this is related more to IDE debug error handling (Tools>Options>General)?
'                    setting to "Break in Class Module" then back to "Break on Unhandled Errors" seems to fix?
'                    no changes were made to this subroutine
' ---------------------------------
Private Sub btnContinue_Click()
On Error GoTo Err_Handler
       
    'clear year & park (prevents NULL errors & click continue if values aren't set)
    cbxYear.Value = "SEL"
    cbxPark.Value = ""
       
    'open target species list
    DoCmd.OpenForm "frm_Tgt_Species", acNormal, , , , , TempVars("TgtYear")
        
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnContinue_Click[form_fsub_Select_Park_Year])"
    End Select
    Resume Exit_Sub
End Sub
