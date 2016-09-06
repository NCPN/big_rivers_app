Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8640
    DatasheetFontHeight =9
    ItemSuffix =14
    Left =5100
    Top =2610
    Right =13995
    Bottom =8715
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xedda7c0f2fb5e440
    End
    RecordSource ="qry_Unknown_species"
    Caption ="Unknown Species List"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =255
    AllowLayoutView =0
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
            Height =1080
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =780
                    Width =2160
                    Height =240
                    FontWeight =700
                    Name ="Unknown_Code_Label"
                    Caption ="Unknown Species Code"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3360
                    Top =780
                    Width =1560
                    Height =240
                    FontWeight =700
                    Name ="Plant_Description_Label"
                    Caption ="Plant Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =180
                    Width =3360
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label7"
                    Caption ="Unknown Species List"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7440
                    Top =180
                    Width =1020
                    Height =300
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7440
                    Top =600
                    Width =1020
                    Height =300
                    TabIndex =1
                    Name ="ButtonNew"
                    Caption ="Add New"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5880
                    Top =780
                    Width =1260
                    Height =240
                    FontWeight =700
                    Name ="Label11"
                    Caption ="Confirmed As"
                    Tag ="DetachedLabel"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1215
                    Left =5100
                    Top =240
                    Width =1320
                    TabIndex =2
                    Name ="ConfirmedFilter"
                    RowSourceType ="Value List"
                    RowSource ="\"Not Confirmed\";\"Confirmed\";\"All\""
                    ColumnWidths ="1215"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4440
                            Top =240
                            Width =600
                            Height =245
                            FontWeight =700
                            Name ="Filter_Label"
                            Caption ="Filter"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =420
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =480
                    Height =255
                    ColumnWidth =2310
                    Name ="Unknown_ID"
                    ControlSource ="Unknown_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =1920
                    Height =288
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Unknown_Code"
                    ControlSource ="Unknown_Code"
                    StatusBarText ="Temporary code for unknown species - Line point form"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2460
                    Top =60
                    Width =3300
                    Height =288
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Plant_Description"
                    ControlSource ="Plant_Description"
                    StatusBarText ="General description"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7800
                    Top =60
                    Width =720
                    Height =300
                    TabIndex =3
                    Name ="ButtonDetails"
                    Caption ="Details"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6060
                    Top =60
                    Width =840
                    Height =288
                    TabIndex =4
                    Name ="Confirmed_Code"
                    ControlSource ="Confirmed_Code"
                    StatusBarText ="Confirmed species code"

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
Option Compare Database

Private Sub Form_Load()
  DoCmd.ApplyFilter "", "IsNull(confirmed_code)"
End Sub

Private Sub ButtonNew_Click()
On Error GoTo Err_ButtonNew_Click

    Dim stDocName As String

    stDocName = "frm_Unknown_Species"
 '   DoCmd.OpenForm stDocName, , , , acFormAdd
    DoCmd.OpenForm stDocName, , , , acFormAdd, acDialog
    Me.Requery
  '  DoCmd.Close acForm, "frm_List_Unknown"  ' cannot close now because it was opened as acdialog
Exit_ButtonNew_Click:
    Exit Sub

Err_ButtonNew_Click:
    MsgBox Err.Description
    Resume Exit_ButtonNew_Click
    
End Sub

Private Sub ConfirmedFilter_AfterUpdate()
If Me!ConfirmedFilter = "Not Confirmed" Then
        DoCmd.ApplyFilter "", "IsNull(confirmed_code)"
ElseIf Me!ConfirmedFilter = "Confirmed" Then
        DoCmd.ApplyFilter "", " NOT IsNull(confirmed_code)"
Else
  Forms!frm_List_Unknown.filter = ""
End If
End Sub

Private Sub ButtonDetails_Click()
On Error GoTo Err_ButtonDetails_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Unknown_Species"
    
    stLinkCriteria = "[Unknown_ID]=" & "'" & Me![Unknown_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonDetails_Click:
    Exit Sub

Err_ButtonDetails_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDetails_Click
    
End Sub

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click

    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub
