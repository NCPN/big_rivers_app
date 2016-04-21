Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9864
    DatasheetFontHeight =11
    ItemSuffix =5
    Right =15828
    Bottom =9408
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0x32f638d6d6a9e440
    End
    Caption ="_Default"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    FitToPage =1
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
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =360
            BackColor =15849926
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =840
                    Top =60
                    Width =4200
                    Height =300
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label4"
                    Caption ="=[Report].[Caption]"
                    GridlineColor =10921638
                    LayoutCachedLeft =840
                    LayoutCachedTop =60
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =360
                End
            End
        End
        Begin PageHeader
            Height =360
            Name ="PageHeaderSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            Height =7560
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin PageFooter
            Height =360
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    Left =5400
                    Top =60
                    Width =2460
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPages"
                    Caption ="=[Page] & \" | \" & [Pages]"
                    GridlineColor =10921638
                    LayoutCachedLeft =5400
                    LayoutCachedTop =60
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =300
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =360
            Name ="ReportFooter"
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
' Form:         _Default
' Level:        Framework form
' Version:      1.00
'
' Description:  _Default form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, November 10, 2015
' References:
'  Allen Browne, April 2010
'  http://allenbrowne.com/ser-43.html
' Revisions:    BLC - 11/10/2015 - 1.00 - initial version
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
' Events
'---------------------

' ---------------------------------
' Sub:          XX
' Description:  XX event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 10, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/10/2015 - initial version
' ---------------------------------
Private Sub XX()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - XX[Default form])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------

' ---------------------------------
' Function:     NoData
' Description:  report actions when no data is found
' Assumptions:  -
' Parameters:   rpt - report being referenced
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 10, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/10/2015 - initial version
' ---------------------------------
Public Function NoData(rpt As Report)
On Error GoTo Err_Handler

    'Purpose: Called by report's NoData event.
    'Usage: =NoData([Report])
    Dim strCaption As String   'Caption of report.
    
    strCaption = rpt.Caption
    If strCaption = vbNullString Then
        strCaption = rpt.Name
    End If
    
    DoCmd.CancelEvent
    MsgBox "There are no records to include in report """ & _
        strCaption & """.", vbInformation, "No Data..."


Exit_Function:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[_Default report])"
    End Select
    Resume Exit_Function
End Function
