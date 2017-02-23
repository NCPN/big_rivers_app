Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
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
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2880
    DatasheetFontHeight =11
    ItemSuffix =2
    Left =3855
    Top =2430
    Right =28545
    Bottom =15015
    TimerInterval =1000
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x89c8fece91e4e440
    End
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnTimer ="[Event Procedure]"
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
        Begin Section
            Height =540
            BackColor =0
            Name ="Detail"
            AlternateBackColor =0
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =120
                    Top =120
                    Width =2040
                    Height =285
                    BorderColor =8355711
                    ForeColor =65535
                    Name ="lblChkDb"
                    Caption ="Checking database...."
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =405
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =215
                    Left =60
                    Top =120
                    Width =180
                    Height =300
                    BackColor =0
                    BorderColor =10921638
                    Name ="rctHider"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =240
                    LayoutCachedHeight =420
                    BackThemeColorIndex =-1
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
' Form:         PreSplash
' Level:        Application form
' Version:      1.00
' Basis:        -
'
' Description:  PreSplash form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, February 22, 2017
' References:   -
' Revisions:    BLC - 2/22/2017 - 1.00 - initial version
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
' Source/date:  Bonnie Campbell, February 22, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/22/2017 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'set form timer interval (milliseconds)
    Me.TimerInterval = 1000
    
    lblChkDb.Caption = "Checking database..."
    
    'initialize app settings
    initApp
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[PreSplash form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Timer
' Description:  form periodic actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Microsoft, Unknown
'   https://msdn.microsoft.com/en-us/library/office/ff192530.aspx
' Source/date:  Bonnie Campbell, February 22, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/22/2017 - initial version
' ---------------------------------
Private Sub Form_Timer()
On Error GoTo Err_Handler

'    Dim i As Integer
'
'    i = 10000
'
'    Do While i > 0
'        Me.rctHider.Left = 0 + i * (Me.Width / 10000)
'        i = i - 1
'    Loop
    
    Me.lblChkDb.ForeColor = RandomColor(lngYellow, lngSalmon)

    'exit if
    If TempVars("Connected") Then
        DoCmd.Close
    Else
        lblChkDb.Caption = "Still not connected..."
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Timer[PreSplash form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Close
' Description:  form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 22, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/22/2017 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    If TempVars("Connected") Then _
        DoCmd.OpenForm "Splash", acNormal, , , , acWindowNormal, Me.Name

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[PreSplash form])"
    End Select
    Resume Exit_Handler
End Sub
