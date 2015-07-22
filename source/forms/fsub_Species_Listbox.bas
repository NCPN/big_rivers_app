Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9120
    DatasheetFontHeight =11
    ItemSuffix =17
    Left =570
    Top =1560
    Right =4260
    Bottom =5580
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x5897ca243799e440
    End
    RecordSource ="SELECT tlu_NCPN_Plants.Master_PLANT_Code AS Code, tlu_NCPN_Plants.Master_Species"
        " AS Species, Switch(tlu_NCPN_Plants.LU_Code Is Null,\" \",tlu_NCPN_Plants.LU_Cod"
        "e<>\"\",tlu_NCPN_Plants.LU_Code) AS LUCode FROM tlu_NCPN_Plants ORDER BY tlu_NCP"
        "N_Plants.Master_Species;"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyDown ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
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
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
        Begin PageHeader
            DisplayWhen =1
            Height =1320
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =840
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblResultsHdr"
                    Caption ="Results"
                    GridlineColor =10921638
                    LayoutCachedWidth =840
                    LayoutCachedHeight =372
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =600
                    Width =4800
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblChooseSpeciesType"
                    Caption ="Double click the species to add it to your target list."
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =600
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =915
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =1020
                    Width =1440
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCodeHdr"
                    Caption ="Code"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =1320
                End
                Begin Label
                    OverlapFlags =85
                    Left =1680
                    Top =1020
                    Width =2520
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesHdr"
                    Caption ="Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =1680
                    LayoutCachedTop =1020
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =1320
                End
                Begin Line
                    OverlapFlags =85
                    Left =1620
                    Top =1020
                    Width =0
                    Height =299
                    Name ="lineHdrSeparator"
                    GridlineColor =10921638
                    LayoutCachedLeft =1620
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =1319
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7200
                    Top =960
                    Width =1800
                    Height =300
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCurrentRecord"
                    GridlineColor =10921638

                    LayoutCachedLeft =7200
                    LayoutCachedTop =960
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =1260
                End
            End
        End
        Begin Section
            Height =300
            Name ="Detail"
            OnPaint ="[Event Procedure]"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Height =300
                    FontSize =10
                    BackColor =9699294
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCode"
                    ControlSource ="Code"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =300
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Width =2520
                    Height =300
                    FontSize =10
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSpecies"
                    ControlSource ="Species"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1680
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =300
                End
                Begin Line
                    OverlapFlags =85
                    Left =1620
                    Width =0
                    Height =299
                    Name ="lineListSeparator"
                    GridlineColor =10921638
                    LayoutCachedLeft =1620
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =299
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4320
                    Width =1980
                    Height =300
                    FontSize =10
                    TabIndex =2
                    BackColor =9699294
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLUCode"
                    ControlSource ="LUCode"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4320
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =300
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6420
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =3
                    BackColor =9699294
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTransectOnly"
                    ControlSource ="Transect_Only"
                    GridlineColor =10921638

                    LayoutCachedLeft =6420
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =300
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7560
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =4
                    BackColor =9699294
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTgtAreaID"
                    ControlSource ="Target_Area_ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =7560
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =300
                    BackThemeColorIndex =-1
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =360
            Name ="PageFooterSection"
            AutoHeight =1
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
' MODULE:       Form_fsub_Species_Listbox
' Description:  Species selction functions & procedures
'               and for lists which exceed standard listbox capacity
'
' Source/date:  Bonnie Campbell, 2/18/2015
' Revisions:    BLC - 2/18/2015 - initial version
'               BLC, 5/1/2015 - renamed from sfrmSpeciesListbox to fsub_Species_Listbox
'               BLC, 6/30/2015 - removed unused private version of tbxCode_DblClick (public sub used)
' =================================

'=================================================================
'  Declarations
'=================================================================
Dim curID As String 'Integer

' ---------------------------------
' SUB:          Form_Load
' Description:  Form loading routine
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 18, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/18/2015 - initial version
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler

    'initial data fill
    fillList Me.Parent, Me.Parent.Controls("fsub_Species_Listbox"), Forms("frm_Tgt_Species")!lbxTgtSpecies
    'fillList Forms("frm_Tgt_Species"), Me, Forms("frm_Tgt_Species")!lbxTgtSpecies

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_fsub_Species_Listbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Form_Current
' Description:  Actions for current detail record
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Rabbit July 11, 2011
' http://bytes.com/topic/access/answers/914781-set-colour-current-record
' March 6, 2010
' http://www.upsizing.co.uk/Art53_Highlight.aspx
' Adapted:      Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/4/2015 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    'set selected record ID
    curID = Me.tbxCode 'Nz(Me.tbxMasterCode, 0)
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[Form_fsub_Species_Listbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Detail_Paint
' Description:  Actions for clicking tbxCode
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Rabbit July 11, 2011
' http://bytes.com/topic/access/answers/914781-set-colour-current-record
' Adapted:      Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/4/2015 - initial version
' ---------------------------------
Private Sub Detail_Paint()
On Error GoTo Err_Handler

    'set selected record backcolor
    If Me.tbxLUCode = curID Then
        Me.Detail.backcolor = lngYelLime
        Me.tbxLUCode.backcolor = lngYelLime
        'Me.tbxSpecies.backcolor = lngYelLime
        Me.tbxLUCode.backcolor = lngYelLime
        
    Else
        Me.Detail.backcolor = lngWhite
        'Me.tbxCode.backcolor = lngWhite
    End If
       
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Paint[Form_fsub_Species_Listbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxCode_Click
' Description:  Actions for clicking tbxCode
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015 - initial version
'   BLC - 5/20/2015 - changed to Me.tbxLUCode vs. Me.tbxMasterCode
' ---------------------------------
Private Sub tbxCode_Click()
On Error GoTo Err_Handler

    'set selected record ID
    curID = Me.tbxLUCode
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxCode_Click[Form_fsub_Species_Listbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxSpecies_Click
' Description:  Actions for clicking tbxSpecies
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/4/2015 - initial version
'   BLC - 5/20/2015 - switched from tbxMasterCode to tbxLUCode
' ---------------------------------
Private Sub tbxSpecies_Click()
On Error GoTo Err_Handler

    'set selected record ID
    curID = Me.tbxLUCode 'Nz(Me.tbxMasterCode, 0)
       
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxSpecies_Click[Form_fsub_Species_Listbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxLUCode_Click
' Description:  Actions for clicking tbxCode (was tbxMasterCode)
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
'   BLC - 5/20/2015 - changed to tbxLUCode from tbxMasterCode
' ---------------------------------
Private Sub tbxLUCode_Click()
On Error GoTo Err_Handler

    'set selected record ID
    curID = Me.tbxLUCode

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxLUCode_Click[Form_fsub_Species_Listbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxCode_DblClick
' Description:  Actions for clicking tbxCode
' Assumptions:  Species with empty lookup codes must first be fixed before being added to a list.
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015 - initial version
'   BLC - 2/23/2015 - added lblTgtSpeciesCount update
'   BLC - 5/10/2015 - exposed event as Public to allow calls from main form
'   BLC - 5/20/2015 - switched from tbxMasterCode to tbxLUCode,
'                     added transect only & tgt area ID
'   BLC - 5/26/2015 - added 0 for passing base value for TgtAreaID to target species listbox
'   BLC - 5/27/2015 - added check for missing LU Codes
'                     (species w/ missing codes cannot be added to target list)
'   BLC - 6/9/2015 -  enable preview and save list buttons on species double click
'   BLC - 6/10/2015 - enable reset button on species double click
' ---------------------------------
Public Sub tbxCode_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    Dim item As String
    Dim lbx As ListBox
    
    'check for empty Lookup code (LUCode)
    If IsNull(tbxLUCode) Or IsEmpty(tbxLUCode) Or Len(Trim(tbxLUCode)) = 0 Then
        
        MsgBox "Species " & tbxSpecies & " is missing a lookup code (LUCode). " & _
            vbCrLf & vbCrLf & "This code is required before the species can be added to a target list. " & _
            vbCrLf & vbCrLf & "Please determine the appropriate code and enter it into the master " & _
            "plant species list." & _
            vbCrLf & vbCrLf & "Contact the project ecologist/data manager to add the species. ", _
            vbExclamation, "Missing Lookup Code!"

        'email species desired
        
        GoTo Exit_Sub
    End If
    
    'add components of item (code, species (UT or whatever), & ITIS) to listbox

    'prepare item for listbox value
    item = tbxCode & ";" & tbxSpecies & ";" & tbxLUCode & ";0;0;" '& tbxTransectOnly & ";" & tbxTgtAreaID & ";" 'tbxMasterCode
    
    'check listbox for duplicate & skip if already present (col 0 vs 2)
    If IsListDuplicate(Forms("frm_Tgt_Species").Controls("lbxTgtSpecies"), 2, tbxLUCode) Then
        'duplicate, so exit
        GoTo Exit_Sub
    End If

    Set lbx = Forms("frm_Tgt_Species").Controls("lbxTgtSpecies")
    
    With lbx
        'add item if not duplicate
        .AddItem item
    
        'update target species count
        Forms("frm_Tgt_Species").Controls("lblTgtSpeciesCount").Caption = .ListCount - 1 & " species"
        
    End With
    
    'enable reset, preview & save
    With Forms("frm_Tgt_Species")
        .Controls("btnReset").Enabled = True
        .Controls("btnPreviewList").Enabled = True
        .Controls("btnSaveList").Enabled = True
    End With
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxCode_DblClick[Form_fsub_Species_Listbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Form_KeyDown
' Description:  Respond to Up/Down in a continuous form by moving to next record
' Assumptions:  Active control's EnterKeyBehaviro is OFF
' Parameters:   frm - form for key behavior
'               KeyCode - code for key being pressed (integer)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Allen Browne via Jeanette Cunningham, Apr 13, 2010
' http://www.pcreview.co.uk/threads/need-to-get-the-up-down-arrow-keys-working.3995845/
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015  - initial version
' ---------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

    Call ContinuousUpDown(Me, KeyCode)
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_KeyDown[Form_fsub_Species_Listbox])"
    End Select
    Resume Exit_Sub
End Sub
