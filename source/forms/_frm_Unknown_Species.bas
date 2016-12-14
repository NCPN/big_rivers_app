Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =9
    ItemSuffix =46
    Left =3825
    Top =3750
    Right =18345
    Bottom =14745
    DatasheetGridlinesColor =12632256
    Filter ="[Unknown_ID]='20110415113257-756092607.975006'"
    RecSrcDt = Begin
        0xedfd33e8cd12e340
    End
    RecordSource ="tbl_Unknown_Species"
    Caption ="frm_Unknown_Species"
    BeforeInsert ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyDown ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =7200
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =10140
                    Top =120
                    Width =570
                    ColumnWidth =2310
                    Name ="Unknown_ID"
                    ControlSource ="Unknown_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10800
                    Top =120
                    Width =570
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Species_ID"
                    ControlSource ="Species_ID"
                    StatusBarText ="Foreign key to tbl_Quadrat_Species"

                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1980
                    Top =1620
                    Width =6000
                    TabIndex =6
                    Name ="Plant_Description"
                    ControlSource ="Plant_Description"
                    StatusBarText ="General description"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =180
                            Top =1620
                            Width =1800
                            Height =240
                            FontWeight =700
                            Name ="Plant_Description_Label"
                            Caption ="General Description"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2040
                    Top =2100
                    Width =7320
                    TabIndex =8
                    Name ="Salient_Feature"
                    ControlSource ="Salient_Feature"
                    StatusBarText ="Most salient feature"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =180
                            Top =2100
                            Width =1860
                            Height =240
                            FontWeight =700
                            Name ="Salient_Feature_Label"
                            Caption ="Most Salient Feature"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1140
                    Top =2580
                    Width =3180
                    ColumnWidth =2310
                    TabIndex =9
                    Name ="Leaf_Type"
                    ControlSource ="Leaf_Type"
                    StatusBarText ="Leaf type: compound/simple, arrangement"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =180
                            Top =2580
                            Width =960
                            Height =240
                            FontWeight =700
                            Name ="Leaf_Type_Label"
                            Caption ="Leaf Type"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1260
                    Top =3060
                    Width =3060
                    ColumnWidth =2310
                    TabIndex =10
                    Name ="Margin"
                    ControlSource ="Margin"
                    StatusBarText ="Leaf margin"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =180
                            Top =3060
                            Width =1110
                            Height =240
                            FontWeight =700
                            Name ="Margin_Label"
                            Caption ="Leaf Margin"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2520
                    Top =3540
                    Width =6840
                    TabIndex =11
                    Name ="Other_Characteristics"
                    ControlSource ="Other_Characteristics"
                    StatusBarText ="Other leaf characteristics:  pubescence, sap, stipules"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =180
                            Top =3540
                            Width =2340
                            Height =240
                            FontWeight =700
                            Name ="Other_Characteristics_Label"
                            Caption ="Other Leaf Characteristics"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2040
                    Top =4020
                    Width =7320
                    TabIndex =12
                    Name ="Stem_Characteristics"
                    ControlSource ="Stem_Characteristics"
                    StatusBarText ="Stem characteristics: shape, pubescence, bud"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =180
                            Top =4020
                            Width =1860
                            Height =240
                            FontWeight =700
                            Name ="Stem_Characteristics_Label"
                            Caption ="Stem Characteristics"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2160
                    Top =4500
                    Width =7200
                    TabIndex =13
                    Name ="Flower_Characteristics"
                    ControlSource ="Flower_Characteristics"
                    StatusBarText ="Flower characteristics: color location floral formula"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =180
                            Top =4500
                            Width =1980
                            Height =240
                            FontWeight =700
                            Name ="Flower_Characteristics_Label"
                            Caption ="Flower Characteristics"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =3720
                    Top =4980
                    Width =5640
                    TabIndex =14
                    Name ="General_Characteristics"
                    ControlSource ="General_Characteristics"
                    StatusBarText ="General and microhabitat characteristics"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =180
                            Top =4980
                            Width =3540
                            Height =240
                            FontWeight =700
                            Name ="General_Characteristics_Label"
                            Caption ="General and Microhabitat Characteristics"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1200
                    Top =5520
                    Width =735
                    Height =300
                    TabIndex =15
                    Name ="Collected"
                    ControlSource ="Collected"
                    StatusBarText ="Was plant collected"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            Left =180
                            Top =5460
                            Width =960
                            Height =240
                            FontWeight =700
                            Name ="Collected_Label"
                            Caption ="Collected?"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1260
                    Top =5940
                    Width =2310
                    ColumnWidth =2310
                    TabIndex =19
                    Name ="Best_Guess"
                    ControlSource ="Best_Guess"
                    StatusBarText ="Best guess species name"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =180
                            Top =5940
                            Width =1080
                            Height =240
                            FontWeight =700
                            Name ="Best_Guess_Label"
                            Caption ="Best Guess"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4020
                    Top =120
                    Width =3420
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label28"
                    Caption ="Unknown Plant Species"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListWidth =675
                    Left =1200
                    Top =1140
                    Width =1380
                    TabIndex =3
                    Name ="Plant_Type"
                    ControlSource ="Plant_Type"
                    RowSourceType ="Value List"
                    RowSource ="\"herb\";\"shrub\";\"tree\";\"grass\";\"sedge\";\"other\""
                    ColumnWidths ="675"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =1140
                            Width =1020
                            Height =245
                            FontWeight =700
                            Name ="Plant Type_Label"
                            Caption ="Plant Type"
                            EventProcPrefix ="Plant_Type_Label"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =9300
                    Top =180
                    Width =1020
                    Height =300
                    TabIndex =24
                    Name ="ButtonClose"
                    Caption ="Close Form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1620
                    Top =660
                    TabIndex =2
                    Name ="Unknown_Code"
                    ControlSource ="Unknown_Code"
                    StatusBarText ="Temporary code for unknown species"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =660
                            Width =1440
                            Height =240
                            FontWeight =700
                            Name ="Label32"
                            Caption ="Unknown_Code:"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =1740
                    Top =6420
                    Width =2580
                    TabIndex =21
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Confirmed_Code"
                    ControlSource ="Confirmed_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Utah_Species FROM tlu_"
                        "NCPN_Plants WHERE (((tlu_NCPN_Plants.Utah_PLANT_Code) Is Not Null)) ORDER BY tlu"
                        "_NCPN_Plants.Utah_Species; "
                    ColumnWidths ="0;3600"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            Left =180
                            Top =6420
                            Width =1545
                            Height =245
                            FontWeight =700
                            Name ="Confirmed To Be_Label"
                            Caption ="Confirmed To Be"
                            EventProcPrefix ="Confirmed_To_Be_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =9240
                    Top =6420
                    Width =1260
                    TabIndex =23
                    Name ="Identified_Date"
                    ControlSource ="Identified_Date"
                    Format ="Short Date"
                    StatusBarText ="Date of identification - Line point form"
                    InputMask ="99/99/0000;0;_"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =7860
                            Top =6420
                            Width =1380
                            Height =240
                            FontWeight =700
                            Name ="Label38"
                            Caption ="Identified Date"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListWidth =795
                    Left =4560
                    Top =1140
                    Width =1200
                    TabIndex =4
                    Name ="Forb_Grass_Type"
                    ControlSource ="Forb_Grass_Type"
                    RowSourceType ="Value List"
                    RowSource ="\"annual\";\"perennial\""
                    ColumnWidths ="795"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2880
                            Top =1140
                            Width =1680
                            Height =245
                            FontWeight =700
                            Name ="Forbs and Grasses_Label"
                            Caption ="Forbs and Grasses"
                            EventProcPrefix ="Forbs_and_Grasses_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1050
                    Left =7680
                    Top =1140
                    TabIndex =5
                    Name ="Perennial_Grasses"
                    ControlSource ="Perennial_Grasses"
                    RowSourceType ="Value List"
                    RowSource ="\"bunchgrass\";\"rhizomatous\""
                    ColumnWidths ="1050"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6060
                            Top =1140
                            Width =1575
                            Height =245
                            FontWeight =700
                            Name ="Perennial Grasses_Label"
                            Caption ="Perennial Grasses"
                            EventProcPrefix ="Perennial_Grasses_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =119
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1980
                    Left =2880
                    Top =5460
                    Width =1620
                    TabIndex =16
                    Name ="Collected_by"
                    ControlSource ="Collected_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                    ColumnWidths ="0;990;990"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =1680
                            Top =5460
                            Width =1200
                            Height =245
                            FontWeight =700
                            Name ="Collected by_Label"
                            Caption ="Collected by"
                            EventProcPrefix ="Collected_by_Label"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4680
                    Top =6000
                    TabIndex =20
                    Name ="Have_Photos"
                    ControlSource ="Have_Photos"
                    StatusBarText ="Are there photos? - Line point form"

                    LayoutCachedLeft =4680
                    LayoutCachedTop =6000
                    LayoutCachedWidth =4940
                    LayoutCachedHeight =6240
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3840
                            Top =5940
                            Width =780
                            Height =240
                            FontWeight =700
                            Name ="Label42"
                            Caption ="Photos"
                            LayoutCachedLeft =3840
                            LayoutCachedTop =5940
                            LayoutCachedWidth =4620
                            LayoutCachedHeight =6180
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =10380
                    Top =1620
                    Width =420
                    TabIndex =7
                    Name ="Position"
                    ControlSource ="Position"
                    StatusBarText ="Position on Transect (m)"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =8280
                            Top =1620
                            Width =2100
                            Height =240
                            FontWeight =700
                            Name ="Label39"
                            Caption ="Position onTransect (m)"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1980
                    Left =5760
                    Top =6420
                    Width =1860
                    TabIndex =22
                    Name ="Identified_by"
                    ControlSource ="Identified_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                    ColumnWidths ="0;990;990"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4560
                            Top =6420
                            Width =1200
                            Height =245
                            FontWeight =700
                            Name ="Label41"
                            Caption ="Identified by"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =2055
                    Left =5640
                    Top =5460
                    Width =2100
                    Height =255
                    TabIndex =17
                    Name ="Method"
                    ControlSource ="Method"
                    RowSourceType ="Value List"
                    RowSource ="\"Point Intercept\";\"Exotic Frequency\";\"1-m Belt Shrubs\";\"1-m Belt Seedling"
                        "s\";\"1-m Belt Species Richness\";\"5-m Belt Saplings\";\"Overstory Census\";\"S"
                        "ite Impact Exotic Species\""
                    ColumnWidths ="2055"

                    LayoutCachedLeft =5640
                    LayoutCachedTop =5460
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =5715
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4800
                            Top =5460
                            Width =720
                            Height =245
                            FontWeight =700
                            Name ="Method_Label"
                            Caption ="Method"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =5460
                            LayoutCachedWidth =5520
                            LayoutCachedHeight =5705
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =270
                    Left =9720
                    Top =5460
                    Width =660
                    Height =255
                    TabIndex =18
                    Name ="Transect_Number"
                    ControlSource ="Transect_Number"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3"
                    ColumnWidths ="270"

                    LayoutCachedLeft =9720
                    LayoutCachedTop =5460
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =5715
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8040
                            Top =5460
                            Width =1590
                            Height =245
                            FontWeight =700
                            Name ="Transect Number_Label"
                            Caption ="Transect Number"
                            EventProcPrefix ="Transect_Number_Label"
                            LayoutCachedLeft =8040
                            LayoutCachedTop =5460
                            LayoutCachedWidth =9630
                            LayoutCachedHeight =5705
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =================================
' MODULE:       Form_UnknownSpecies
' Description:  Unknown species properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, July 5, 2016
' Revisions:    BLC - 7/5/2016 - initial version adapted from Upland frm_Unknown_Species
' =================================

'=================================================================
'  Properties
'=================================================================

'=================================================================
'  Subroutines & Functions
'=================================================================

' ---------------------------------
' SUB:          Form_Load
' Description:  Unknown species form load action
' Assumptions:  none
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, July 5, 2016 - for NCPN tools
' Revisions:
'   BLC - 7/5/2016  - initial version
'   BLC - 7/5/2016 - adjusted for big rivers & new form name SpeciesSearch vs. frmSpeciesSearch
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler
    

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[UnknownSpecies form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_KeyDown
' Description:  handles form's key down actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 2015
' Revisions:    BLC, 8/21/2014 - initial version
'               BLC, 7/5/2016 - adapted for big rivers
' ---------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

    'capture ESC & let user determine if fields should be cleared
    CaptureEscapeKey KeyCode
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_KeyDown[Form_frm_Unknown_Species])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_BeforeInsert
' Description:  handles form's pre-insert actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 2015
' Revisions:    BLC, 8/21/2014 - initial version
'               BLC, 7/5/2016 - adapted for big rivers, added documentation
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler
    
    ' Create the GUID primary key value
    If IsNull(Me!Unknown_ID) Then
        If GetDataType("tbl_Unknown_Species", "Unknown_ID") = dbText Then
            Me.Unknown_ID = fxnGUIDGen
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[UnknownSpecies form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          UnknownCode_AfterUpdate
' Description:  Handles unknown code after update actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 2015
' Revisions:    BLC, 8/21/2014 - initial version
'               BLC, 7/5/2016 - adapted for big rivers
' ---------------------------------
Private Sub tbxUnknown_Code_AfterUpdate()
On Error GoTo Err_Handler

  If Not IsNull(Me!Unknown_Code) Then
        If Left(Me!Unknown_Code, 3) <> "UNK" Then
          MsgBox "Unknown code must be prefixed by UNK.", , "Unknown Code"
 '         DoCmd.CancelEvent
 '         SendKeys "{ESC}"
          Me.Undo
        ElseIf Not IsNull(DLookup("[Unknown_ID]", "tbl_Unknown_Species", "[Unknown_Code] = '" & Me.Unknown_Code & "'")) Then
          MsgBox "Unknown code already exists in table.", , "Unknown Code"
          Me.Undo
        End If
  End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxUnknownCode_AfterUpdate[UnknownSpecies form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Close
' Description:  Handles btnClose click actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 5, 2016
' Revisions:    BLC, 7/5/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler


    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[UnknownSpecies form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          GetDataType
' Description:  Retrieves data type of table field
' Parameters:   strTableName - name of database table (string)
'               strFieldName - name of table field (string)
' Returns:      data type (integer)
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Unknown
' Revisions:    Unknown - initial version
'               BLC, 7/5/2016 - revised for big rivers
' ---------------------------------
Public Function GetDataType(strTableName As String, strFieldName As String) As Integer
On Error Resume Next
    
    Dim intResult As Integer

    intResult = CurrentDb.TableDefs(strTableName)(strFieldName).Type
    GetDataType = intResult

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[UnknownSpecies form])"
    End Select
    Resume Exit_Handler
End Function
