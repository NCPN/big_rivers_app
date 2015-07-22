Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10935
    DatasheetFontHeight =11
    ItemSuffix =30
    Right =11184
    Bottom =5988
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x72574db34b86e440
    End
    Caption ="Create Target Species List"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnActivate ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =6000
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =5760
                    Top =1080
                    Width =4320
                    Height =4032
                    FontSize =10
                    BoundColumn =2
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxTgtSpecies"
                    RowSourceType ="Value List"
                    RowSource ="Code;Species;LUCode;;;ABAR;Abronia argillosa;ABRARG;0;0;ABTH;Abutilon theophrast"
                        "i;ABUTHE;0;0;ACSP;Acamptopappus sphaerocephalus;ACASPH;0;0;ACNE2;Acer negundo;AC"
                        "ENEG;0;0"
                    ColumnWidths ="1440;2520;720;288;288"
                    OnDblClick ="[Event Procedure]"
                    OnKeyUp ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Target species"
                    GridlineColor =10921638

                    LayoutCachedLeft =5760
                    LayoutCachedTop =1080
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =5112
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =720
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblTgtSpecies"
                            Caption ="Target Species"
                            GridlineColor =10921638
                            LayoutCachedLeft =5580
                            LayoutCachedTop =720
                            LayoutCachedWidth =7020
                            LayoutCachedHeight =1035
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =840
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblParkHdr"
                    Caption ="CARE"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =900
                    LayoutCachedHeight =432
                End
                Begin Subform
                    OverlapFlags =85
                    Left =420
                    Top =1080
                    Width =3960
                    Height =4032
                    TabIndex =1
                    BorderColor =10921638
                    Name ="fsub_Species_Listbox"
                    SourceObject ="Form.fsub_Species_Listbox"
                    GridlineColor =10921638

                    LayoutCachedLeft =420
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =5112
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =660
                            Width =1200
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSpeciesListbox"
                            Caption ="UT Species"
                            GridlineColor =10921638
                            LayoutCachedLeft =180
                            LayoutCachedTop =660
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =975
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =8340
                    Top =780
                    Width =1440
                    Height =276
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTgtSpeciesCount"
                    Caption ="4 species"
                    ControlTipText ="Number of species in the current list"
                    GridlineColor =10921638
                    LayoutCachedLeft =8340
                    LayoutCachedTop =780
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =1056
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =3000
                    Top =780
                    Width =1440
                    Height =276
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSfrmSpeciesCount"
                    Caption ="3195 species"
                    ControlTipText ="Number of species in the current list"
                    GridlineColor =10921638
                    LayoutCachedLeft =3000
                    LayoutCachedTop =780
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =1056
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =2880
                    Top =120
                    Width =4320
                    Height =315
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear"
                    Caption ="Target Species List for 2017"
                    GridlineColor =10921638
                    LayoutCachedLeft =2880
                    LayoutCachedTop =120
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =435
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7560
                    Top =180
                    Width =1560
                    Height =405
                    TabIndex =2
                    ForeColor =16711680
                    Name ="btnLoad"
                    Caption ="Load List"
                    StatusBarText ="Continue to choose activities"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =7560
                    LayoutCachedTop =180
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =585
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =8640
                    Top =5220
                    Width =1560
                    Height =405
                    TabIndex =3
                    ForeColor =16711680
                    Name ="btnSaveList"
                    Caption ="Save List"
                    StatusBarText ="Save the current list"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =8640
                    LayoutCachedTop =5220
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =5625
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =1380
                    Top =5220
                    Width =1560
                    Height =405
                    TabIndex =4
                    ForeColor =16711680
                    Name ="btnSearch"
                    Caption ="Find Species"
                    StatusBarText ="Find a species..."
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1380
                    LayoutCachedTop =5220
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =5625
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =4380
                    Top =5220
                    Width =1320
                    Height =405
                    TabIndex =5
                    ForeColor =16711680
                    Name ="btnReset"
                    Caption ="Reset List"
                    StatusBarText ="Reset lists to their original state"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =5220
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =5625
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
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =4740
                    Top =1200
                    Width =540
                    Height =480
                    FontWeight =600
                    TabIndex =6
                    ForeColor =8224125
                    Name ="btnAddAll"
                    Caption =">>"
                    StatusBarText ="Reset lists to their original state"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =1680
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =13882323
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =8224125
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =4560
                    Width =540
                    Height =480
                    FontWeight =600
                    TabIndex =7
                    ForeColor =16711680
                    Name ="btnRemoveAll"
                    Caption ="<<"
                    StatusBarText ="Remove all"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =4560
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =5040
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =52479
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =0
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =0
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =3960
                    Width =540
                    Height =480
                    FontWeight =600
                    TabIndex =8
                    ForeColor =16711680
                    Name ="btnRemove"
                    Caption ="<"
                    StatusBarText ="Remove all"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Remove selected"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =3960
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =4440
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =52479
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =0
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =0
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =3
                    WebImagePaddingBottom =3
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =1860
                    Width =540
                    Height =480
                    FontWeight =600
                    TabIndex =9
                    ForeColor =16711680
                    Name ="btnAdd"
                    Caption =">"
                    StatusBarText ="Remove all"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add selected"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =1860
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =2340
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =6750156
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
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
                    WebImagePaddingRight =3
                    WebImagePaddingBottom =3
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6240
                    Top =5220
                    Width =1560
                    Height =405
                    TabIndex =10
                    ForeColor =16711680
                    Name ="btnPreviewList"
                    Caption ="Preview List"
                    StatusBarText ="Save the current list"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6240
                    LayoutCachedTop =5220
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =5625
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
' MODULE:       Form_frm_Tgt_Species
' Description:  Species selction functions & procedures
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC, 2/9/2015 - initial version
'               BLC, 4/30/2015 - integrated into Invasives Reporting tool
'               BLC, 7/7/2015  - btnAdd() bug fix to avoid lbxSpecies compiler error (should have been lbxTgtSpecies)
' =================================

'=================================================================
'  Properties
'=================================================================
' ---------------------------------
' PROPERTY:     Maximized
' Description:  Indicates if form is maximized or not by checking IsZoomed()
' Assumptions:  none
' Parameters:   N/A
' Returns:      True(1) - form is maximized
'               False(0) - form is not maximized
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Get Maximized() As Boolean
     Maximized = IsZoomed(Me.hwnd) * 1
End Property

' ---------------------------------
' PROPERTY:     Minimized
' Description:  Indicates if form is minimized or not by checking IsIconic()
' Assumptions:  none
' Parameters:   N/A
' Returns:      True(1) - form is minimized
'               False(0) - form is not minimized
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Get Minimized() As Boolean
     Minimized = IsIconic(Me.hwnd) * 1
End Property

' ---------------------------------
' PROPERTY LET: Maximized
' Description:  Sets custom form property 'Maximized'
' Assumptions:
' Note:         The IsMax argument must be defined as the same data type
'               returned by the corresponding Property Get procedure for
'               the same custom property.
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Let Maximized(IsMax As Boolean)
     If IsMax Then
         Me.SetFocus
         DoCmd.Maximize
     Else
         Me.SetFocus
         DoCmd.Restore
     End If
End Property

' ---------------------------------
' PROPERTY LET: Minimized
' Description:  Sets custom form property 'Minimized'
' Assumptions:
' Note:         The IsMin argument must be defined as the same data type
'               returned by the corresponding Property Get procedure for
'               the same custom property.
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Let Minimized(IsMin As Boolean)
     If IsMin Then
         Me.SetFocus
         DoCmd.Minimize
     Else
         Me.SetFocus
         DoCmd.Restore
     End If
End Property

'=================================================================
'  Subroutines & Functions
'=================================================================

' ---------------------------------
' SUB:          Form_Load
' Description:  Form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/9/2015 - initial version
'   BLC, 5/1/2015 - integrated into Invasives Reporting tool, removed frmSelectYear closure since that form
'                   is no longer needed, added check for species number to ensure >= 0
'   BLC, 5/13/2015 - disabled Remove All button to start & recaptioned btnReset to "Reset List" vs. "Reset Lists"
'                    set btnAdd to enabled to start vs disabled
'   BLC - 6/9/2015 - toggle preview & save list buttons (enabled if lbx has species)
'   BLC - 6/10/2015 - added toggle for reset button
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
    Dim intSpecies As Integer, iSpeciesCount As Integer
    
    Initialize
       
    'set state
    TempVars("state") = getParkState(TempVars("park"))
    
    'set year
    TempVars("TgtYear") = Form.OpenArgs
    
    'prep headers
    lblParkHdr.Caption = TempVars("park")
    lblYear.Caption = "Target Species List for " & Form.OpenArgs
    lblSpeciesListbox.Caption = TempVars("state") & " Species"
    
    'clear headers
    lbxTgtSpecies.RowSource = ""
    
    'initial listbox fill
     fillList Me, lbxTgtSpecies

    'Enable move items lbls (or not)
    btnAddAll.Enabled = False
    
    'Set counts
    iSpeciesCount = GetListCount(lbxTgtSpecies, True)
    lblTgtSpeciesCount.Caption = iSpeciesCount & " species"
    
    'Set Reset button label (reset list vs. lists)
    btnReset.Caption = "Reset List"
    
    DisableControl btnAddAll
    EnableControl btnAdd, lngLtLime, lngBlue, lngDkLime, lngBrtLime, lngLtGreen, lngDkGray, lngLtLime
    DisableControl btnRemove
    DisableControl btnRemoveAll
    
    If iSpeciesCount > 0 Then
        'enable reset, preview & save buttons
        btnReset.Enabled = True
        btnPreviewList.Enabled = True
        btnSaveList.Enabled = True
    Else
        'Disable reset, preview & save buttons
        btnReset.Enabled = False
        btnPreviewList.Enabled = False
        btnSaveList.Enabled = False
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Form_Activate
' Description:  Sets tbxTgtSpecies value
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
'   BLC - 5/1/2015 - added check for no species to prevent # = -1
'   BLC - 5/10/2015 - revised to include generic count function
'   BLC - 6/9/2015 - toggle preview & save list buttons (enabled if lbx has species)
'   BLC - 6/10/2015 - added toggle for reset button
' ---------------------------------
Private Sub Form_Activate()

On Error GoTo Err_Handler
    
    Dim iSpeciesCount As Integer
    
    'set species count
    iSpeciesCount = GetListCount(lbxTgtSpecies, True)
    lblTgtSpeciesCount.Caption = iSpeciesCount & " species"
    
    If iSpeciesCount > 0 Then
        btnReset.Enabled = True
        btnPreviewList.Enabled = True
        btnSaveList.Enabled = True
    Else
        btnReset.Enabled = False
        btnPreviewList.Enabled = False
        btnSaveList.Enabled = False
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Activate[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnLoad_Click
' Description:  Load list from previous year
' Assumptions:  -
' Parameters:   none
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC, 3/5/2015 - initial version
'   BLC, 5/1/2015 - updated frmSelectList to frm_Select_List to conform to standards
' ---------------------------------
Private Sub btnLoad_Click()

On Error GoTo Err_Handler

Dim aryFields() As String
Dim aryFieldTypes() As Variant
Dim strCode As String, strSpecies As String, strLUCode As String
Dim iRow As Integer, iTransectOnly As Integer, iTgtAreaID As Integer
Dim rs As DAO.Recordset
    
    iRow = Me.Controls("lbxTgtSpecies").ListCount - 1
    
    ReDim Preserve aryFields(0 To iRow)
        
    'header row (iRow = 0)
    aryFields(0) = "Code;Species;LUCode;Transect_Only;Target_Area_ID"   'iRow = 0
    aryFieldTypes = Array(dbText, dbText, dbText, dbInteger, dbInteger)

    'data rows (iRow > 0)
    For iRow = 1 To lbxTgtSpecies.ListCount - 1
        
        ' ---------------------------------------------------
        '  NOTE: listbox column MUST have a non-zero width to retrieve its value
        ' ---------------------------------------------------
         strCode = lbxTgtSpecies.Column(0, iRow) 'column 0 = Master_PLANT_Code (Code)
         strSpecies = lbxTgtSpecies.Column(1, iRow) 'column 1 = Species name (Species)
         strLUCode = lbxTgtSpecies.Column(2, iRow) 'column 2 = LU_Code (LUCode)
         iTransectOnly = Nz(lbxTgtSpecies.Column(3, iRow), 0) 'column 3 = Transect_Only (TransectOnly)
         iTgtAreaID = Nz(lbxTgtSpecies.Column(4, iRow), 0) 'column 4 = Target_Area_ID (TgtAreaID)
        
        aryFields(iRow) = strCode & ";" & strSpecies & ";" & strLUCode & ";" & iTransectOnly & ";" & iTgtAreaID
        
    Next
    
    'save the existing records to temp_Listbox_Recordset & replace any existing records
    SetListRecordset lbxTgtSpecies, True, aryFields, aryFieldTypes, "temp_Listbox_Recordset", True

    'open tgt species list form
    DoCmd.OpenForm "frm_Select_List", acNormal, , , , acWindowNormal, Me.name

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLoad_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnReset_Click
' Description:  Reset lists to their original state
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
' ---------------------------------
Private Sub btnReset_Click()
On Error GoTo Err_Handler

    'go back to initial state
    Form_Load

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReset_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtSpecies_Click
' Description:  click event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Allen Browne, June 2006
' http://allenbrowne.com/func-12.html
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 5/10/2015 - added species count update
'   BLC - 5/13/2015 - revised to use global constants vs. tempvars for enabled control
'   BLC - 6/9/2015 - enable preview & save list buttons if species in list
'   BLC - 6/10/2015 - added toggle for reset button
' ---------------------------------
Private Sub lbxTgtSpecies_Click()
On Error GoTo Err_Handler
    
    Dim varItem As Variant
    
   'check for selected items --> if present, enable btnRemove
    If lbxTgtSpecies.ItemsSelected.count > 0 Then
        If btnRemove.backcolor <> CTRL_REMOVE_ENABLED Then
            EnableControl btnRemove, CTRL_REMOVE_ENABLED, TEXT_ENABLED
            EnableControl btnRemoveAll, CTRL_REMOVE_ENABLED, TEXT_ENABLED
        End If
        'enable reset, preview & save
        btnReset.Enabled = True
        btnPreviewList.Enabled = True
        btnSaveList.Enabled = True
    Else
        DisableControl btnRemove
    End If
    
    'set species count
    lblTgtSpeciesCount.Caption = GetListCount(lbxTgtSpecies, True) & " species"
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtSpecies_DblClick
' Description:  Removes items from lbxTgSpecies
' Assumptions:  -
' Parameters:   Cancel - if true cancels action, false runs removal
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
'   BLC - 5/10/2015 - changed from MoveSingleItem to RemoveSelectedItems to handle
'                     removing species versus populating them back to the original species list
'                     added count update
'   BLC - 6/9/2015 - added toggle for save & preview buttons
'   BLC - 6/10/2015 - added toggle for reset button
' ---------------------------------
Private Sub lbxTgtSpecies_DblClick(Cancel As Integer)
    
On Error GoTo Err_Handler
    Dim iSpeciesCount As Integer


    'MoveSingleItem Me, "lbxTgtSpecies", "lbxTgtSpecies"
    RemoveSelectedItems lbxTgtSpecies

    'set species count
    iSpeciesCount = GetListCount(lbxTgtSpecies, True)
    lblTgtSpeciesCount.Caption = iSpeciesCount & " species"

    If iSpeciesCount > 0 Then
        btnReset.Enabled = True
        btnPreviewList.Enabled = True
        btnSaveList.Enabled = True
    Else
        btnReset.Enabled = True
        btnPreviewList.Enabled = False
        btnSaveList.Enabled = False
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_DblClick[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtSpecies_KeyUp
' Description:  Deselects items after control update
' Assumptions:  -
' Parameters:   KeyCode - keystroke code
'               Shift - if shift key has been pressed
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 5/13/2015 - revised to use global constants vs. tempvars for enabled control
' ---------------------------------
Private Sub lbxTgtSpecies_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

'    If lbxSpecies.ItemsSelected.Count > 0 And lblRemove.backcolor <> TempVars.item("ctrlRemoveEnabled") Then
    If btnRemove.backcolor <> CTRL_REMOVE_ENABLED Then
        EnableControl btnRemove, CTRL_REMOVE_ENABLED, TEXT_ENABLED
        EnableControl btnRemoveAll, CTRL_REMOVE_ENABLED, TEXT_ENABLED
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_KeyUp[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnAdd_Click
' Description:  Add selected items to list
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
'   BLC - 5/10/2015 - added update for species count
'   BLC - 6/9/2015  - enable preview and save list buttons if species in list
'   BLC - 6/10/2015 - added toggle for reset button
'   BLC - 7/7/2015  - fixed bug bringing up compiler error (lbxSpecies not defined),
'                     should have been lbxTgtSpecies vs. lbxSpecies
' ---------------------------------
Private Sub btnAdd_Click()
On Error GoTo Err_Handler
    
    'ignore if 'disabled'
    'If lblAdd.backcolor = lngGray Then GoTo Exit_Sub
    If btnAdd.backcolor = lngGray Then GoTo Exit_Sub
    
    'MoveSingleItem Me, "lbxSpecies", "lbxTgtSpecies"
    MoveSingleItem Me, "fsub_Species_Listbox", "lbxTgtSpecies"

    'update count
    lblTgtSpeciesCount.Caption = GetListCount(lbxTgtSpecies, True) & " species"
    
    'enable reset, preview & save
    If GetListCount(lbxTgtSpecies, True) > 0 Then
        btnReset.Enabled = True
        btnPreviewList.Enabled = True
        btnSaveList.Enabled = True
    End If

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAdd_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnRemove_Click
' Description:  Remove selected items from list
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
'   BLC - 5/10/2015 - changed from MoveSingleItem to RemoveSelectedItems to handle
'                     removing species from list vs. adding back to original list
'                     added update for species count
'   BLC - 5/13/2015 - revised to use global constants vs. tempvars for disabled control
'                     disabled btnRemove, btnRemoveAll when target species count = 0
'   BLC - 6/9/2015 - disable preview and save list buttons if no species in list
'   BLC - 6/10/2015 - disable reset button if no species present
' ---------------------------------
Private Sub btnRemove_Click()
On Error GoTo Err_Handler
    
    'ignore if 'disabled'
    If btnRemove.backcolor = CTRL_DISABLED Then GoTo Exit_Sub
    
    'MoveSingleItem Me, "lbxTgtSpecies", "fsub_Species_Listbox"
    RemoveSelectedItems lbxTgtSpecies
    
    'update count
    Dim iSpeciesCount As Integer
    iSpeciesCount = GetListCount(lbxTgtSpecies, True)
    lblTgtSpeciesCount.Caption = iSpeciesCount & " species"
    
    'turn off Remove buttons if iCount = 0
    If iSpeciesCount = 0 Then
        DisableControl btnRemove
        DisableControl btnRemoveAll
        btnReset.Enabled = False
        btnPreviewList.Enabled = False
        btnSaveList.Enabled = False
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRemove_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnAddAll_Click
' Description:  Add all items to list
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
'   BLC - 5/10/2015 - added update for species count
'   BLC - 6/9/2015 - enable preview and save list buttons if species in list
'   BLC - 6/10/2015 - enable reset button if species in list
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub btnAddAll_Click()
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'fetch recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset(TempVars("strSQL"))
    
    'MoveAllItems Me, "lbxSpecies", "lbxTgtSpecies"
    MoveAllItems Me, "", "lbxTgtSpecies"

    'update count
    lblTgtSpeciesCount.Caption = GetListCount(lbxTgtSpecies, True) & " species"
    
    'enable reset, preview & save
    btnReset.Enabled = True
    btnPreviewList.Enabled = True
    btnSaveList.Enabled = True

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddAll_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnRemoveAll_Click
' Description:  Remove all items from list
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
'   BLC - 5/10/2015 - changed from MoveAllItems to Form_Load to handle
'                     removing all species vs. moving them to original listbox
' ---------------------------------
Private Sub btnRemoveAll_Click()
On Error GoTo Err_Handler
    
    'MoveAllItems Me, "lbxTgtSpecies", "fsub_Species_Listbox"
    'go back to initial state
    Form_Load

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRemoveAll_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnSaveList_Click
' Description:  Save list items
' Assumptions:  Current and future year's lists can be amended by adding/deleting species.
'               Prior year lists can only add new species to the list.
'               Species deletions for prior years can only be done via the backend database.
'               This is to prevent inadvertent list deletions.
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
'   BLC - 5/13/2015 - added LU code and fixed MasterCode bug which substituted LU_Code as Master_Code from tlu_NPCN_Plants
'   BLC - 5/20/2015 - reverted to LUCode for col 2, Code (Master_Plant_Code) in col 0
'   BLC - 6/3/2015 - added ability to delete from previously created lists for current or future years,
'                    prior years cannot delete species - done by first deleting then inserting
'                    list for the park/year
'                    added message alert for prior year list changes to ensure the user really wants to add
'                    new species to them (deletions are done via the BE)
'   BLC - 6/4/2015 - renamed tempTgtSpecies to temp_Tgt_Species
'   BLC - 6/10/2015 - adjusted SQL strings to accommodate new tbl_Target_List and changes to tbl_Target_Species
'                     (park & year shift to tbl_Target_List)
'   BLC - 6/11/2015 - added check to retrieve Tgt_List_ID when list previously existed to populate new species records
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub btnSaveList_Click()
On Error GoTo Err_Handler

    Dim iRow As Integer, i As Integer, iTransectOnly As Integer, iTgtAreaID As Integer, iResponse As Integer, tgtListID As Integer
    Dim strMasterCode As String, strSpecies As String, strLUCode As String
    Dim strSQL As String, strInsert As String
    Dim varReturn As Variant
    Dim blnAddToList As Boolean
    Dim wrkCurrent As DAO.Workspace
    
    'default
    blnAddToList = False
    
    'show action
    DoCmd.Hourglass True
    
    'delete the full list for current or future years
    If CInt(TempVars("TgtYear")) > 0 And CInt(TempVars("TgtYear")) > Year(Now()) Then
    
        MsgBox "Removing previously saved " & TempVars("park") & " - " & TempVars("TgtYear") & _
                " species. " & vbCrLf & vbCrLf & _
                "Your new list will be saved shortly.", vbInformation, _
                "Deleting Old " & TempVars("park") & " - " & TempVars("TgtYear") & " List"
                
        'set statusbar notice
        varReturn = SysCmd(acSysCmdSetStatus, "Removing old list...")
        
        'remove the old list
        strSQL = "DELETE DISTINCTROW  tbl_Target_Species.* " & _
                "FROM tbl_Target_Species " & _
                "INNER JOIN tbl_Target_List ON tbl_Target_Species.Tgt_List_ID_FK = tbl_Target_List.Tgt_List_ID " & _
                "WHERE tbl_Target_List.Park_Code = '" & TempVars("park") & "' " & _
                "AND tbl_Target_List.Target_Year = " & TempVars("TgtYear") & ";"
        
        CurrentDb.Execute strSQL, dbFailOnError
        
         'pause to view status bar
        Pause 3
        
        blnAddToList = True
    Else
    
        'warn the user, but allow them to choose to add to the previous year list (or not)
        Dim strCurrPrev As String
        strCurrPrev = IIf(CInt(TempVars("TgtYear")) = Year(Now()), "the current", "a previous")

        iResponse = MsgBox("The list you are saving is for " & strCurrPrev & " year ( " & _
                TempVars("park") & " - " & TempVars("TgtYear") & " )." & vbCrLf & vbCrLf & _
                "If necessary, species can be added to current/prior year lists, but they cannot be removed." & vbCrLf & vbCrLf & _
                "If the list you are saving has new species, they will be added." & vbCrLf & vbCrLf & _
                "Removed species will be ignored." & vbCrLf & vbCrLf & _
                "Do you really want to add species to the " & _
                TempVars("park") & " - " & TempVars("TgtYear") & "?", _
                vbYesNoCancel, "Altering List for a Current/Previous Year!")
        
        'check response - vbOK(1), vbCancel(2), vbAbort(3), vbRetry(4), vbIgnore(5), vbYes(6), vbNo(7)
        'allow addition only if user says "Yes!"
        If iResponse = 6 Then blnAddToList = True
        
    End If
    
    'skip it?
    If blnAddToList = False Then GoTo Exit_Sub
    
    'set status bar
    varReturn = SysCmd(acSysCmdSetStatus, "Preparing new list... ")
    
    '-------------------------------------------------
    ' add new list record if it doesn't exist
    '-------------------------------------------------
    strSQL = "SELECT * " & _
            "FROM tbl_Target_List " & _
            "WHERE tbl_Target_List.Park_Code = '" & TempVars("park") & "' " & _
            "AND tbl_Target_List.Target_Year = " & TempVars("TgtYear") & ";"
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'default
    tgtListID = 0
    
    Set rs = CurrentDb.OpenRecordset(strSQL)
    
    'check if list exists
    If rs.BOF And rs.EOF Then
        
        'insert & retrieve ID
        strSQL = "INSERT INTO tbl_Target_List(Park_Code, Target_Year, Created) " & _
                 "VALUES ('" & TempVars("park") & "', " & TempVars("TgtYear") & ", Now() );"
        
        CurrentDb.Execute strSQL, dbFailOnError
        
        tgtListID = CurrentDb.OpenRecordset("SELECT @@IDENTITY")(0)
    
    Else
    
        'update the last modified date & retrieve the list ID
        strSQL = "UPDATE tbl_Target_List " & _
                "SET Last_Modified = Now() " & _
                "WHERE Park_Code = '" & TempVars("park") & "' AND Target_Year = " & TempVars("TgtYear") & ";"

        CurrentDb.Execute strSQL, dbFailOnError
        
        'fetch the ID
        strSQL = "SELECT TOP 1 tbl_Target_List.Tgt_List_ID FROM tbl_Target_List " & _
                "WHERE Park_Code = '" & TempVars("park") & "' AND Target_Year = " & TempVars("TgtYear") & ";"

        Set rs = CurrentDb.OpenRecordset(strSQL)

        If Not (rs.BOF And rs.EOF) Then tgtListID = rs("Tgt_List_ID")
    
    End If
    
    'start @ row 1 (headers = row 0)
    For iRow = 1 To lbxTgtSpecies.ListCount - 1
       
       ' ---------------------------------------------------
       '  NOTE: listbox column MUST have a non-zero width to retrieve its value
       ' ---------------------------------------------------
        strMasterCode = lbxTgtSpecies.Column(0, iRow) 'column 0 = Master_PLANT_Code
        strSpecies = lbxTgtSpecies.Column(1, iRow) 'column 1 = Species name
        strLUCode = lbxTgtSpecies.Column(2, iRow) 'column 2 = LU_Code
        iTransectOnly = Nz(lbxTgtSpecies.Column(3, iRow), 0) 'column 3 = Transect_Only
        iTgtAreaID = Nz(lbxTgtSpecies.Column(4, iRow), 0) 'column 4 = Target_Area_ID
        
       ' ---------------------------------------------------
       '  Check if item exists in tbl_TgtSpecies for Park, Year, Species combo
       ' ---------------------------------------------------
        strSQL = "SELECT * FROM tbl_Target_Species " & _
                 "INNER JOIN tbl_Target_List ON tbl_Target_Species.Tgt_List_ID_FK = tbl_Target_List.Tgt_List_ID " & _
                 "WHERE tbl_Target_Species.Master_PLANT_Code_FK ='" & strMasterCode & "' " & _
                 "AND tbl_Target_List.Park_Code = '" & TempVars("park") & "' " & _
                 "AND tbl_Target_List.Target_Year = " & TempVars("TgtYear") & ";"

        Set rs = CurrentDb.OpenRecordset(strSQL) 'CurrentDb.Execute(strSQL, dbFailOnError) >> doesn't compile expected function or variable
      
        'check if there are no records (rs.BOF & rs.EOF are both true)
        If rs.BOF And rs.EOF Then
            
            'set statusbar notice
            varReturn = SysCmd(acSysCmdSetStatus, "Saving " & strSpecies & "...")
            
            'prepare SQL
            strSQL = "INSERT INTO tbl_Target_Species" _
                    & "(Tgt_List_ID_FK, Master_Plant_Code_FK, Species_Name, LU_Code, " _
                    & "Transect_Only, Target_Area_ID) " _
                    & "VALUES "

            'prepare insert value
            strInsert = "(" & tgtListID & ",'" & strMasterCode & "','" & strSpecies & "','" & strLUCode _
                        & "'," & iTransectOnly & "," & iTgtAreaID & ");"
            
            'add comma if more than one row to insert
            'If (lbxTgtSpecies.ListCount - 1) > 1 And iRow < (lbxTgtSpecies.ListCount - 1) Then strInsert = strInsert & ","
            
            'finalize SQL
            strSQL = strSQL & strInsert
            
            'save full target list (insert value) [NOTE: MS Access does not support multiple insert statements, must go 1 @ a time]
            CurrentDb.Execute strSQL, dbFailOnError
            
        End If
        
    Next

    ' check for temp query & clear if it exists
    If QueryExists("temp_Tgt_Species") Then
        CurrentDb.QueryDefs.Delete "temp_Tgt_Species"
    End If
    
    'open target list
    Dim qdf As QueryDef
    
    Set qdf = CurrentDb.QueryDefs("qry_Tgt_Species_List")
    
    strSQL = qdf.sql
    
    strSQL = "SELECT tbl_Target_List.Park_Code AS Park, " & _
             "tbl_Target_List.Target_Year AS TgtYear, " & _
             "Master_Plant_Code_FK, Species_Name, LU_Code, " & _
             "Priority, Transect_Only, Target_Area_ID " & _
             "FROM tbl_Target_Species " & _
             "INNER JOIN tbl_Target_List ON tbl_Target_Species.Tgt_List_ID_FK = tbl_Target_List.Tgt_List_ID " & _
             "WHERE (((tbl_Target_List.Target_Year) = CInt(tgtYear)) " & _
             "And ((LCase([tbl_Target_List].[Park_Code])) = LCase(park))) " & _
             "ORDER BY tbl_Target_Species.Species_Name;"
    
    'replace values
    strSQL = Replace(strSQL, "(park)", "('" & TempVars("park") & "')")
    strSQL = Replace(strSQL, "(tgtYear)", "(" & TempVars("TgtYear") & ")")
    
    'DoCmd.OpenQuery "qryTgtSpeciesList", acViewNormal, acReadOnly
    'DoCmd.RunSQL strSQL <=== NO! not on a SELECT...
    
    CurrentDb.CreateQueryDef("temp_Tgt_Species").sql = strSQL
    DoCmd.OpenQuery "temp_Tgt_Species"
    
    'set statusbar notice
    varReturn = SysCmd(acSysCmdSetStatus, "Targetlist save complete.")
    
    'pause to view status bar
    Pause 4
    
    'reset status bar
    varReturn = SysCmd(acSysCmdSetStatus, " ")

    'close form
    DoCmd.Close acForm, Me.name

Exit_Sub:
    DoCmd.Hourglass False
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSaveList_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnSearch_Click
' Description:  Opens species search to find species for populating target list
' Description:  Reset lists to their original state
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC, 3/3/2015  - initial version
'   BLC, 4/30/2015 - integrated into Invasives Reporting tool & updated form naming
' ---------------------------------
Private Sub btnSearch_Click()
On Error GoTo Err_Handler
    Dim originForm As String
    
    originForm = Me.name
    
    'open species search form
    DoCmd.OpenForm "frm_Species_Search", acNormal, , , , acWindowNormal, originForm
    If Forms("frm_Species_Search").Minimized Then DoCmd.Restore

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSearch_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnPreviewList_Click
' Description:  Open a preview report listing of the current target list (based on species in the target species listbox)
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/3/2015 - initial version
' ---------------------------------
Private Sub btnPreviewList_Click()
On Error GoTo Err_Handler

    Dim rs As DAO.Recordset

    'prepare the recordset for the report (populate temp_Listbox_Recordset)
    'SetListRecordset lbxTgtSpecies, True, aryFields, aryFieldTypes, "temp_Listbox_Recordset", True
    AddListToTable lbxTgtSpecies
    
    'open the report in preview mode
    DoCmd.OpenReport "rpt_Tgt_Species_List", acViewReport, , , acWindowNormal, "preview"
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnPreviewList_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Form_Close
' Description:  Actions for closing form
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015 - initial version
'   BLC - 3/4/2015  - closed species search form
'   BLC - 4/30/2015 - integrated into Invasives Reporting tool & updated form naming
'   BLC - 5/27/2015 - added clear temp_Listbox_Recordset table
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'clear tempvars
    TempVars.Remove ("park")
    TempVars.Remove ("state")

    'clear temp_Listbox_Recordset table
    ClearTable "temp_Listbox_Recordset"

    'close frmSpeciesSearch if open
    DoCmd.Close acForm, "frm_Species_Search"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub
