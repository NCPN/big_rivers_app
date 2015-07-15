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
    Width =6912
    DatasheetFontHeight =11
    ItemSuffix =54
    Left =1980
    Top =1524
    Right =9468
    Bottom =7524
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xc1f3db6ed487e440
    End
    RecordSource ="tbl_Target_Areas"
    Caption ="Photo Details"
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
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            Height =4260
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Width =6912
                    Height =360
                    BackColor =15266810
                    BorderColor =10921638
                    Name ="rctPhotogHdr"
                    GridlineColor =10921638
                    LayoutCachedWidth =6912
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4920
                    Top =3792
                    Width =1740
                    Height =372
                    ForeColor =16711680
                    Name ="btnNext"
                    Caption ="Save && Next >>"
                    StatusBarText ="Save photo details & move to next photo"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4920
                    LayoutCachedTop =3792
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =4164
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
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1260
                    Top =480
                    Width =3180
                    Height =315
                    ColumnOrder =0
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPhotog"
                    GridlineColor =10921638

                    LayoutCachedLeft =1260
                    LayoutCachedTop =480
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =795
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =480
                            Width =690
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPhotog"
                            Caption ="Name"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =480
                            LayoutCachedWidth =1050
                            LayoutCachedHeight =795
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =120
                    Top =60
                    Width =1266
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPhotogHdr"
                    Caption ="Photographer"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1386
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1260
                    Top =1440
                    Width =3180
                    Height =315
                    ColumnOrder =1
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPhotoNum"
                    GridlineColor =10921638

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1440
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =1755
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1440
                            Width =780
                            Height =300
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPhotoNum"
                            Caption ="Photo #"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1440
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =1740
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =600
                    Top =2280
                    Width =3888
                    Height =1560
                    ColumnOrder =2
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDescription"
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =2280
                    LayoutCachedWidth =4488
                    LayoutCachedHeight =3840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1860
                            Width =1080
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblDescription"
                            Caption ="Description"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1860
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =2175
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4560
                    Top =1440
                    Width =2220
                    Height =720
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblPhotoNumHint"
                    Caption ="Photo # hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =4560
                    LayoutCachedTop =1440
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =2160
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4560
                    Top =2280
                    Width =2220
                    Height =660
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblDescriptionHint"
                    Caption ="Description hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =4560
                    LayoutCachedTop =2280
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =2940
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =960
                    Width =6912
                    Height =360
                    BackColor =16381933
                    BorderColor =16381933
                    Name ="rctPhotoDetailHdr"
                    GridlineColor =10921638
                    LayoutCachedTop =960
                    LayoutCachedWidth =6912
                    LayoutCachedHeight =1320
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =120
                    Top =1020
                    Width =1266
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPhotoDetail"
                    Caption ="Photo Details"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1386
                    LayoutCachedHeight =1335
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
' FORM:         Form_fsub_Photo_FTOR_Details
' Description:  Photo detail functions & procedures for feature, transect, overview & reference photos
'
' Source/date:  Bonnie Campbell, 7/13/2015
' Revisions:    BLC - 7/13/2015 - initial version
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
' Adapted:      Bonnie Campbell, July 13, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/13/2015 - initial version
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
    Dim i As Integer, iYear As Integer
    Dim strValueList As String
    
'    Initialize
    
    'prepare value list
 '   strValueList = "'SEL';'Select Year';"

 '   iYear = Year(Now()) + 2

 '   For i = 1 To 6
 '       strValueList = strValueList & "'" & iYear & "';'" & iYear & "';"
 '       iYear = iYear - 1
 '   Next
    
 '   cbxYear.RowSource = strValueList
 '   cbxYear.Value = "SEL"

    'disable next to start
    btnNext.Enabled = False

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[form_fsub_Photo_Other_Details])"
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
' SUB:          btnNext_Click
' Description:  Save photo info & go to next actions
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, July 13, 2015 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 7/13/2015 - initial version
' ---------------------------------
Private Sub btnNext_Click()
On Error GoTo Err_Handler
       
    'clear year & park (prevents NULL errors & click continue if values aren't set)
'    cbxYear.Value = "SEL"
'    cbxPark.Value = ""
       
    'open target species list
'    DoCmd.OpenForm "frm_Tgt_Species", acNormal, , , , , TempVars("TgtYear")
    
    ' save & move to next photo in tree
    
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNext_Click[form_fsub_Photo_Other_Details])"
    End Select
    Resume Exit_Sub
End Sub
