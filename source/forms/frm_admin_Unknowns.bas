Version =20
VersionRequired =20
Begin Form
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
    Width =10080
    DatasheetFontHeight =9
    ItemSuffix =24
    Left =3810
    Top =225
    Right =14175
    Bottom =6510
    DatasheetGridlinesColor =12632256
    Filter ="IsNull(confirmed_code)"
    RecSrcDt = Begin
        0xb8cda39ccb7be340
    End
    RecordSource ="qry_admin_Unknown"
    Caption ="Unknown Species List"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
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
            Height =1680
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =60
                    Top =1380
                    Width =1740
                    Height =240
                    FontWeight =700
                    Name ="Unknown_Code_Label"
                    Caption ="Unknown Species "
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4440
                    Top =1380
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
                    Left =2685
                    Top =180
                    Width =4155
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label7"
                    Caption ="Update Unknown Species"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8820
                    Top =240
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =420
                    Top =840
                    Width =2715
                    Height =240
                    FontWeight =700
                    Name ="Label11"
                    Caption ="Replace Selected Unknown By"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2100
                    Top =1380
                    Width =1350
                    Height =240
                    FontWeight =700
                    Name ="Label15"
                    Caption ="Plant Type"
                    Tag ="DetachedLabel"
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4095
                    Left =3180
                    Top =840
                    Width =1560
                    Height =239
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"40\""
                    Name ="Replace_By"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_lu_NCPN_Plants.Master_PLANT_Code, qry_lu_NCPN_Plants.Utah_PLANT_Code,"
                        " qry_lu_NCPN_Plants.Utah_Species FROM qry_lu_NCPN_Plants; "
                    ColumnWidths ="0;645;3450"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6945
                    Top =1380
                    Width =1695
                    Height =240
                    FontWeight =700
                    Name ="Label19"
                    Caption ="Confirmed As"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5040
                    Top =840
                    Width =960
                    Height =300
                    TabIndex =2
                    ForeColor =255
                    Name ="ButtonDetails"
                    Caption ="Replace"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1215
                    Left =7800
                    Top =840
                    Width =1320
                    TabIndex =3
                    Name ="ConfirmedFilter"
                    RowSourceType ="Value List"
                    RowSource ="\"Not Confirmed\";\"Confirmed\";\"All\""
                    ColumnWidths ="1215"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7140
                            Top =840
                            Width =600
                            Height =245
                            FontWeight =700
                            Name ="Filter_Label"
                            Caption ="Filter"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =360
                    Top =240
                    Width =1380
                    Height =299
                    TabIndex =4
                    Name ="ButtonMaster"
                    Caption ="Master Species"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
                    SpecialEffect =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =1680
                    Height =239
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
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3720
                    Top =60
                    Width =3000
                    Height =239
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Plant_Description"
                    ControlSource ="Plant_Description"
                    StatusBarText ="General description"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2040
                    Top =60
                    Height =239
                    TabIndex =3
                    Name ="Plant_Type"
                    ControlSource ="Plant_Type"
                    StatusBarText ="Plant type:  herb, shrub, tree, grass, sedge, other"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7380
                    Top =60
                    Width =780
                    TabIndex =4
                    Name ="Confirmed_Code"
                    ControlSource ="Confirmed_Code"
                    StatusBarText ="Confirmed species code"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8640
                    Top =60
                    Width =720
                    Height =300
                    TabIndex =5
                    Name ="ButtonEdit"
                    Caption ="Details"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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

Private Sub ButtonDetails_Click()

  Dim db As DAO.Database
  Dim InputTable As DAO.Recordset
  Dim strSQL As String
  Dim strPrompt As String
  Dim intFieldIndex As Integer
  Dim intIndexSave As Integer
  Dim intRecordCount As Integer
  Dim SColumn As String
  Dim AColumn As String
  Dim strAlive As String
  Dim blnIsValid As Boolean
        
  On Error GoTo Err_Handler
  If IsNull(Me!Replace_By) Then
    MsgBox "Replace By species required."
    Exit Sub
  End If
  strPrompt = "You are about to replace all occurrences of " & Me!Unknown_Code & " with " & Me!Replace_By & "." & vbCrLf & "Do you want to proceed?"
  If MsgBox(strPrompt, vbYesNo, "Species Replace") = vbNo Then
    Exit Sub
  End If
  DoCmd.Hourglass True
  '  Start with tbl_LP_Intercept top canopy
  Set db = CurrentDb
  Set InputTable = db.OpenRecordset("tbl_LP_Intercept")
    Do Until InputTable.EOF
      If InputTable!Top = Me!Unknown_Code Then
        intFieldIndex = 1
        blnIsValid = True
        Do Until intFieldIndex > 10
          SpeciesColumn = "LCS" & intFieldIndex
          AliveColumn = "LCA" & intFieldIndex
          If IsNull(InputTable(SpeciesColumn)) Then
            Exit Do
          ElseIf InputTable(SpeciesColumn) = Me!Replace_By And InputTable(AliveColumn) = InputTable!Alive Then
            MsgBox "Duplicate species in top canopy - bypassed."
            blnIsValid = False
            Exit Do
          End If
          intFieldIndex = intFieldIndex + 1
        Loop
        If blnIsValid Then
          InputTable.Edit
            InputTable!Top = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
        End If
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing
  
  '  Now do lower canopy
  Set InputTable = db.OpenRecordset("tbl_LP_Intercept")
    Do Until InputTable.EOF
      intFieldIndex = 1
      Do Until intFieldIndex > 10
        SColumn = "LCS" & intFieldIndex
        If IsNull(InputTable(SColumn)) Then
          Exit Do
        ElseIf InputTable(SColumn) = Me!Unknown_Code Then
          blnIsValid = True  ' assume no duplicate
          AColumn = "LCA" & intFieldIndex
          strAlive = InputTable(AColumn)  ' Save alive/dead flag of unknown code
          intIndexSave = intFieldIndex
          intFieldIndex = 1   '  Reset index
          Do Until intFieldIndex > 10
            SColumn = "LCS" & intFieldIndex
            If IsNull(InputTable(SColumn)) Then
              Exit Do
            End If
            AColumn = "LCA" & intFieldIndex
            If InputTable(SColumn) = Me!Replace_By And InputTable(AColumn) = strAlive Then
              MsgBox "Duplicate species in lower canopy - bypassed."
              blnIsValid = False
              Exit Do
            End If
            intFieldIndex = intFieldIndex + 1
          Loop
          If InputTable!Top = Me!Replace_By And InputTable!Alive = strAlive Then
            MsgBox "Duplicate species in top canopy - bypassed."
          ElseIf blnIsValid Then
            SColumn = "LCS" & intIndexSave
            InputTable.Edit
              InputTable(SColumn) = Me!Replace_By
            InputTable.Update
            intRecordCount = intRecordCount + 1
          End If
          Exit Do
        End If   ' End if for unknown hit
        intFieldIndex = intFieldIndex + 1
      Loop
      InputTable.MoveNext
    Loop  '  Loop for input table
    InputTable.Close
    Set InputTable = Nothing
    
  ' Now do LP soil surface
  Set InputTable = db.OpenRecordset("tbl_LP_Intercept")
    Do Until InputTable.EOF
      If InputTable!Surface = Me!Unknown_Code Then
          InputTable.Edit
            InputTable!Surface = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing
    
  ' Now do belt shrubs
  Set InputTable = db.OpenRecordset("tbl_LP_Shrub")
    Do Until InputTable.EOF
      If InputTable!Species = Me!Unknown_Code Then
        If Not IsNull(DLookup("[Shrub_ID]", "tbl_LP_Shrub", "[Transect_ID] = '" & InputTable!Transect_ID & "' AND [Species] = '" & Me!Replace_By & "'")) Then
          MsgBox "Duplicate belt shrub species - bypassed."
        Else
          InputTable.Edit
            InputTable!Species = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
        End If
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing
    
  ' Now do exotics
  Set InputTable = db.OpenRecordset("tbl_LP_Exotic")
    Do Until InputTable.EOF
      If InputTable!Species = Me!Unknown_Code Then
        If Not IsNull(DLookup("[Exotic_ID]", "tbl_LP_Exotic", "[Transect_ID] = '" & InputTable!Transect_ID & "' AND [Species] = '" & Me!Replace_By & "'")) Then
          MsgBox "Duplicate exotic species - bypassed."
        Else
          InputTable.Edit
            InputTable!Species = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
        End If
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing
  
  ' Now do exotic frequency
  Set InputTable = db.OpenRecordset("tbl_LP_Exotic_freq")
    Do Until InputTable.EOF
      If InputTable!Species = Me!Unknown_Code Then
        If Not IsNull(DLookup("[Exotic_ID]", "tbl_LP_Exotic_freq", "[Transect_ID] = '" & InputTable!Transect_ID & "' AND [Species] = '" & Me!Replace_By & "'")) Then
          MsgBox "Duplicate exotic species - bypassed."
        Else
          InputTable.Edit
            InputTable!Species = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
        End If
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing
  
  ' Now do site impact exotics
  Set InputTable = db.OpenRecordset("tbl_Dist_Exotic")
    Do Until InputTable.EOF
      If InputTable!Species = Me!Unknown_Code Then
        If Not IsNull(DLookup("[Exotic_ID]", "tbl_Dist_Exotic", "[Impact_ID] = '" & InputTable!Impact_ID & "' AND [Species] = '" & Me!Replace_By & "'")) Then
          MsgBox "Duplicate exotic species - bypassed."
        Else
          InputTable.Edit
            InputTable!Species = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
        End If
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing
    
  ' Now do tree seedlings
  Set InputTable = db.OpenRecordset("tbl_LP_Seedling")
    Do Until InputTable.EOF
      If InputTable!Species = Me!Unknown_Code Then
        If Not IsNull(DLookup("[Seedling_ID]", "tbl_LP_Seedling", "[Transect_ID] = '" & InputTable!Transect_ID & "' AND [Species] = '" & Me!Replace_By & "'")) Then
          MsgBox "Duplicate tree seedling species - bypassed."
        Else
          InputTable.Edit
            InputTable!Species = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
        End If
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing
  
  ' Monument trees
  Set InputTable = db.OpenRecordset("tbl_Monument")
    Do Until InputTable.EOF
      If InputTable!Species = Me!Unknown_Code Then
          InputTable.Edit
            InputTable!Species = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing
  
  ' OT Census
  Set InputTable = db.OpenRecordset("tbl_OT_Census")
    Do Until InputTable.EOF
      If InputTable!Species = Me!Unknown_Code Then
          InputTable.Edit
            InputTable!Species = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing
  
  ' 5 meter belt
  Set InputTable = db.OpenRecordset("tbl_OT_Tree_Saplings")
    Do Until InputTable.EOF
      If InputTable!Species = Me!Unknown_Code Then
        If Not IsNull(DLookup("[TS_ID]", "tbl_OT_Tree_Saplings", "[Event_ID] = '" & InputTable!Event_ID & "' AND [Species] = '" & Me!Replace_By & "'")) Then
          MsgBox "Duplicate 5 meter belt species - bypassed."
        Else
          InputTable.Edit
            InputTable!Species = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
        End If
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing

  ' shrub line intercept
  Set InputTable = db.OpenRecordset("tbl_SLI_Gaps")
    Do Until InputTable.EOF
      If InputTable!Species = Me!Unknown_Code Then
        If Not IsNull(DLookup("[SLI_ID]", "tbl_SLI_Gaps", "[Transect_ID] = '" & InputTable!Transect_ID & "' AND [Species] = '" & Me!Replace_By & "' AND " & InputTable!Shrub_End & " BETWEEN [Shrub_Start] AND [Shrub_End]")) Then
          MsgBox "Overlapping shrub line species - bypassed."
        Else
          InputTable.Edit
            InputTable!Species = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
        End If
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing

  ' 1 meter belt additional species
  Set InputTable = db.OpenRecordset("tbl_LP_Add_Species")
    Do Until InputTable.EOF
      If InputTable!Species = Me!Unknown_Code Then
        If Not IsNull(DLookup("[Add_ID]", "tbl_LP_Add_Species", "[Transect_ID] = '" & InputTable!Transect_ID & "' AND [Species] = '" & Me!Replace_By & "'")) Then
          MsgBox "Duplicate 1 meter additional species - bypassed."
        Else
          InputTable.Edit
            InputTable!Species = Me!Replace_By
          InputTable.Update
          intRecordCount = intRecordCount + 1
        End If
      End If
    InputTable.MoveNext
    Loop
  InputTable.Close
  Set InputTable = Nothing
  
'    Me!Confirmed_Code = Me!Replace_By   ' Set confirmed code in unknowns table - deactivated 4/1/09  RD
    DoCmd.Hourglass False
    MsgBox intRecordCount & " record(s) changed."
'    Me.Requery
Exit_ButtonDetails_Click:
    Exit Sub

Err_Handler:
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

Private Sub ConfirmedFilter_AfterUpdate()
If Me!ConfirmedFilter = "Not Confirmed" Then
        DoCmd.ApplyFilter "", "IsNull(confirmed_code)"
ElseIf Me!ConfirmedFilter = "Confirmed" Then
        DoCmd.ApplyFilter "", " NOT IsNull(confirmed_code)"
Else
  Forms!frm_admin_Unknowns.filter = ""
End If
End Sub
Private Sub ButtonEdit_Click()
On Error GoTo Err_ButtonEdit_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Unknown_Species"
    
    stLinkCriteria = "[Unknown_ID]=" & "'" & Me![Unknown_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonEdit_Click:
    Exit Sub

Err_ButtonEdit_Click:
    MsgBox Err.Description
    Resume Exit_ButtonEdit_Click
    
End Sub
Private Sub ButtonMaster_Click()
On Error GoTo Err_ButtonMaster_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonMaster_Click:
    Exit Sub

Err_ButtonMaster_Click:
    MsgBox Err.Description
    Resume Exit_ButtonMaster_Click
    
End Sub
