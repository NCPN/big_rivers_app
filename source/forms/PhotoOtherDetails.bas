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
    Width =6420
    DatasheetFontHeight =11
    ItemSuffix =56
    Left =4530
    Top =4155
    Right =11295
    Bottom =11280
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x36469deccdc4e440
    End
    Caption ="Photo Details"
    OnOpen ="[Event Procedure]"
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
            Height =375
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
                    Width =6420
                    Height =360
                    BackColor =15266810
                    BorderColor =10921638
                    Name ="rctPhotogHdr"
                    GridlineColor =10921638
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
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
            End
        End
        Begin Section
            Height =3780
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =1680
                    Top =3300
                    Width =1740
                    Height =372
                    ForeColor =16711680
                    Name ="btnNext"
                    Caption ="Save && Next >>"
                    StatusBarText ="Save photo details & move to next photo"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1680
                    LayoutCachedTop =3300
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =3672
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
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1260
                    Top =960
                    Width =3180
                    Height =315
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPhotoNum"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a6000000020000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff00010000000000000004000000220000000100000000000000 ,
                        0xfff2000000000000000000000000000000000000000000000000000000000000 ,
                        0x220022000000000049004900660028004c0065006e0028005b006c0062006c00 ,
                        0x500068006f0074006f004e0075006d005d0029003c0038002c0031002c003000 ,
                        0x290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1260
                    LayoutCachedTop =960
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =1275
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000fff200001d00000049004900660028004c0065006e0028005b ,
                        0x006c0062006c00500068006f0074006f004e0075006d005d0029003c0038002c ,
                        0x0031002c0030002900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =960
                            Width =780
                            Height =300
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPhotoNum"
                            Caption ="Photo #"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =960
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =1260
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =600
                    Top =1860
                    Width =3888
                    Height =1380
                    ColumnOrder =2
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDescription"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =1860
                    LayoutCachedWidth =4488
                    LayoutCachedHeight =3240
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1440
                            Width =1080
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblDescription"
                            Caption ="Description"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1440
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =1755
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4560
                    Top =900
                    Width =1860
                    Height =960
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblPhotoNumHint"
                    Caption ="P + Month\015\012(Jan-Sep=0-9,Oct-Dec=A-C) + day(01-31) + \015\0124-digit camera"
                        " seq# \015\012(PA010300 = Jan 1, #300)"
                    GridlineColor =10921638
                    LayoutCachedLeft =4560
                    LayoutCachedTop =900
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =1860
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4620
                    Top =1980
                    Width =1800
                    Height =660
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblDescriptionHint"
                    GridlineColor =10921638
                    LayoutCachedLeft =4620
                    LayoutCachedTop =1980
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =2640
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =480
                    Width =6420
                    Height =360
                    BackColor =16381933
                    BorderColor =16381933
                    Name ="rctPhotoDetailHdr"
                    GridlineColor =10921638
                    LayoutCachedTop =480
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =840
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =120
                    Top =540
                    Width =1266
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPhotoDetail"
                    Caption ="Photo Details"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =540
                    LayoutCachedWidth =1386
                    LayoutCachedHeight =855
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4260
                    Width =1320
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnContacts"
                    Caption ="Add Contact"
                    StatusBarText ="Add new contact"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add new contact"
                    GridlineColor =10921638

                    LayoutCachedLeft =4260
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =360
                    PictureCaptionArrangement =5
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1440
                    Width =2760
                    Height =315
                    TabIndex =4
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x01000000a2000000020000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff00010000000000000004000000200000000100000000000000 ,
                        0xfff2000000000000000000000000000000000000000000000000000000000000 ,
                        0x220022000000000049004900660028004c0065006e0028005b00630062007800 ,
                        0x500068006f0074006f0067005d0029003d0030002c0031002c00300029000000 ,
                        0x0000
                    End
                    Name ="cbxPhotog"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1440"
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =315
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000fff200001b00000049004900660028004c0065006e0028005b ,
                        0x00630062007800500068006f0074006f0067005d0029003d0030002c0031002c ,
                        0x0030002900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Width =690
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPhotog"
                            Caption ="Name"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedWidth =1050
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =4500
                    Top =3300
                    Width =1800
                    TabIndex =5
                    ForeColor =4210752
                    Name ="btnSave"
                    Caption ="Save && Next >>"
                    StatusBarText ="Save photo details & move to next photo"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save photo details & move to next photo"
                    GridlineColor =10921638

                    LayoutCachedLeft =4500
                    LayoutCachedTop =3300
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =3660
                    PictureCaptionArrangement =5
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =3420
                    Width =480
                    FontSize =8
                    TabIndex =6
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxPhotoType"
                    ConditionalFormat = Begin
                        0x0100000098000000020000000100000000000000000000000d00000001000000 ,
                        0x3333ff00ffffff0001000000000000000e0000001b0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004400450056005f004d004f00440045005d003d003100000000005b004400 ,
                        0x450056005f004d004f00440045005d003d00300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =3420
                    LayoutCachedWidth =600
                    LayoutCachedHeight =3660
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                    ConditionalFormat14 = Begin
                        0x0100020000000100000000000000010000003333ff00ffffff000c0000005b00 ,
                        0x4400450056005f004d004f00440045005d003d00310000000000000000000000 ,
                        0x000000000000000000000001000000000000000100000000000000ffffff000c ,
                        0x0000005b004400450056005f004d004f00440045005d003d0030000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =660
                    Top =3420
                    Width =480
                    FontSize =8
                    TabIndex =7
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxID"
                    ConditionalFormat = Begin
                        0x0100000098000000020000000100000000000000000000000d00000001000000 ,
                        0x3333ff00ffffff0001000000000000000e0000001b0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004400450056005f004d004f00440045005d003d003100000000005b004400 ,
                        0x450056005f004d004f00440045005d003d00300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =660
                    LayoutCachedTop =3420
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =3660
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                    ConditionalFormat14 = Begin
                        0x0100020000000100000000000000010000003333ff00ffffff000c0000005b00 ,
                        0x4400450056005f004d004f00440045005d003d00310000000000000000000000 ,
                        0x000000000000000000000001000000000000000100000000000000ffffff000c ,
                        0x0000005b004400450056005f004d004f00440045005d003d0030000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
            End
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
' MODULE:       Form_PhotoOtherDetails
' Level:        Development module
' Version:      1.02
'
' Description:  Photo detail functions & procedures for other & unclassified photos
'
' Source/date:  Bonnie Campbell, 7/13/2015
' Revisions:    BLC - 7/13/2015 - 1.00 - initial version
'               BLC - 2/21/2017 - 1.01 - added Form_Activate() event to handle photographer list updates
'                                        removed Form_Activate() event photographer list updates
'                                        require redefining RowSource for cbxPhoto control
'               BLC - 12/22/2017 - 1.02 - update hints, defaults
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 12/22/2017 - update hints, defaults
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    'set hover
    btnSave.HoverColor = lngGreen
      
    'defaults
    btnSave.Enabled = False
    cbxPhotog.BackColor = lngYellow
    tbxPhotoNum.BackColor = lngYellow
    tbxPhotoType = "OO"  'later set by image node
    tbxID = 0
  
    'set hints
    lblPhotoNumHint.ForeColor = lngBlue
    lblPhotoNumHint.Caption = "P + Month" & vbCrLf & _
                        "(Jan-Sep=0-9,Oct-Dec=A-C)" & vbCrLf & _
                        "+ day(01-31) +" & vbCrLf & _
                        "4-digit camera seq#" & vbCrLf & _
                        "(PA010300 = Jan 1, #300)"
    lblDescriptionHint.ForeColor = lngBlue
    lblDescriptionHint.Caption = ""
    
    'based on node clicked
    Dim nodeinfo() As String
    '0 - M, 1- C, 2-full file path, 3-file name w/o extension
    'nodeinfo = Split(Me.Parent!tvwTree.Object.SelectedItem.Tag, "|")
    'FilePath = nodeinfo(2)
  
    With Me.Parent!tvwTree.Object
    
        Debug.Print "tag: " & .SelectedItem.Tag
    
    End With
  
    'initialize values
    Set Me.cbxPhotog.Recordset = GetRecords("s_contact_list")
    
    
'    ClearForm Me
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[PhotoOtherDetails form])"
    End Select
    Resume Exit_Handler
End Sub

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

    'disable next to start
    btnNext.Enabled = False

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[PhotoOtherDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Activate
' Description:  Actions for form Activateing
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 21, 2017 - for NCPN tools
' Revisions:
'   BLC - 2/21/2017 - initial version
' ---------------------------------
Private Sub Form_Activate()
On Error GoTo Err_Handler

    'update Contacts
    'Me.cbxPhotog.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Activate[PhotoOtherDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPhotoNum_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/14/2016 - initial version
' ---------------------------------
Private Sub tbxPhotoNum_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPhotoNum_AfterUpdate[PhotoFTORDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxDescription_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/14/2016 - initial version
' ---------------------------------
Private Sub tbxDescription_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDescription_AfterUpdate[PhotoFTORDetails form])"
    End Select
    Resume Exit_Handler
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
           
    ' save & move to next photo in tree
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNext_Click[PhotoOtherDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnContacts_Click
' Description:  Add contact button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/14/2016 - initial version
'   BLC - 2/21/2017  - revised to use Photo vs. Tree form
' ---------------------------------
Private Sub btnContacts_Click()
On Error GoTo Err_Handler
    
    DoCmd.OpenForm "Contact", acNormal, , , , , "Photo" '"Tree"
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnContacts_Click[PhotoOther form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSave_Click
' Description:  Save button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    UpsertRecord Me
        
    'move to next node

    'MsgBox Forms("Tree").Controls("tvwTree").SelectedItem.index, vbCritical, " sub node index"

'    Me.Parent.Form.MoveToNext (tbx)

        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[PhotoOtherDetails form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ReadyForSave
' Description:  Check if form values are ready to save
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 8/23/2016 - changed ReadyForSave() to public for mod_App_Data Upsert/SetRecord()
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    'requires: direction facing, photog (comments optional)
    If Len(Nz(cbxPhotog.Value, "")) > 0 _
        And Len(Nz(tbxPhotoNum.Value, "")) > 0 Then
        isOK = True
    End If
    
'    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    btnSave.Enabled = isOK
    
    'refresh form
    Me.Requery
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[PhotoOtherDetails form])"
    End Select
    Resume Exit_Handler
End Sub
