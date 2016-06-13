Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    DefaultView =2
    ViewsAllowed =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7140
    ItemSuffix =17
    Left =3225
    Top =7035
    Right =14430
    Bottom =11640
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5e46d82ab6c4e440
    End
    RecordSource ="SELECT tbl_Events.Location_ID, tbl_Features.*, tbl_Events.Start_date AS Event_da"
        "te FROM tbl_Events INNER JOIN tbl_Features ON tbl_Events.Event_ID = tbl_Features"
        ".Event_ID ORDER BY tbl_Features.Feature_status, tbl_Events.Start_date DESC , tbl"
        "_Features.Feature_type; "
    Caption ="fsub_Features"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowFormView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    DatasheetForeColor12 =33554432
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
        Begin Section
            Height =2700
            BackColor =13025979
            Name ="Detail"
            Begin
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1848
                    Top =480
                    Width =5088
                    Height =285
                    ColumnWidth =4128
                    FontSize =9
                    TabIndex =3
                    Name ="txtFeature_desc"
                    ControlSource ="Feature_desc"
                    StatusBarText ="Brief description of the feature"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1272
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labFeature_desc"
                            Caption ="Travel feature"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1848
                    Top =2340
                    Height =252
                    ColumnWidth =3036
                    FontSize =9
                    TabIndex =7
                    Name ="txtFeature_ID"
                    ControlSource ="Feature_ID"
                    StatusBarText ="Unique identifier for each feature record"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2340
                            Width =996
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labFeature_ID"
                            Caption ="Feature_ID"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1848
                    Top =900
                    Width =2568
                    Height =252
                    ColumnWidth =1224
                    FontSize =9
                    TabIndex =4
                    Name ="txtDistance_m"
                    ControlSource ="Distance_m"
                    StatusBarText ="Distance in meters, measured from the previous point for travel features"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =900
                            Width =1116
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labDistance_m"
                            Caption ="Distance (m)"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1848
                    Top =1260
                    Width =636
                    Height =252
                    ColumnWidth =1356
                    FontSize =9
                    TabIndex =5
                    Name ="txtFeature_azimuth"
                    ControlSource ="Feature_azimuth"
                    StatusBarText ="Azimuth (degrees, declination corrected) from the sampling point to the feature"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1260
                            Width =1224
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labFeature_azimuth"
                            Caption ="Bearing (deg)"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5040
                    Left =1848
                    Top =1620
                    Width =2568
                    Height =252
                    ColumnWidth =888
                    FontSize =9
                    TabIndex =2
                    Name ="cmbFeature_status"
                    ControlSource ="Feature_status"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Marker_Status.Marker_status, tlu_Marker_Status.Marker_status_desc FRO"
                        "M tlu_Marker_Status; "
                    ColumnWidths ="1152;3888"
                    StatusBarText ="Current status of the feature"
                    DefaultValue ="\"Active\""
                    FontName ="Arial"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1620
                            Width =600
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labFeature_status"
                            Caption ="Status"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1848
                    Top =1980
                    Width =1896
                    Height =252
                    ColumnWidth =1260
                    FontSize =9
                    TabIndex =6
                    Name ="txtFeature_updated"
                    ControlSource ="Feature_updated"
                    Format ="yyyy mmm dd"
                    StatusBarText ="Date on which the feature record was last updated"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1980
                            Width =1164
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labFeature_updated"
                            Caption ="Last updated"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =120
                    Width =2112
                    Height =252
                    FontSize =9
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"32\""
                    Name ="cmbFeature_type"
                    ControlSource ="Feature_type"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Feature_Type.Feature_type FROM tlu_Feature_Type; "
                    StatusBarText ="Type of feature"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1140
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="labFeature_type"
                            Caption ="Feature type"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5580
                    Top =2160
                    Name ="txtEvent_date"
                    ControlSource ="Event_date"
                    Format ="yyyy mmm dd"
                    StatusBarText ="Start date of the sampling event"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4140
                            Top =2160
                            Width =945
                            Height =240
                            Name ="labEvent_date"
                            Caption ="Event date"
                        End
                    End
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
' FORM NAME:    fsub_Features_Browser
' Description:  Data browser subform for point travel features
' Data source:  In-line query based on tbl_Features
' Data access:  edit, add and delete
' Pages:        none
' Functions:    none
' References:   none
' Source/date:  John R. Boetsch, October 2008
' Revisions:    <name, date, desc>
' =================================

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Validate the record
    If IsNull(Me.txtFeature_desc) Then
        MsgBox "Please give a brief description of the feature", vbOKOnly, "Validation error"
        Me.txtFeature_desc.SetFocus
        DoCmd.CancelEvent
    ElseIf IsNull(Me.cmbFeature_status) Then
        MsgBox "Please indicate the feature status", vbOKOnly, "Validation error"
        Me.cmbFeature_status.SetFocus
        DoCmd.CancelEvent
    End If
    ' Upon updating the record, set the updated field
    Me.txtFeature_updated = Now()

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbFeature_status_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler

    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
    ' Drop down or pull up the combo box list
        KeyCode = 0
        SendKeys "{F4}"
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
