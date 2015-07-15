Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =9
    ItemSuffix =13
    Left =4110
    Top =180
    Right =11310
    Bottom =3765
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3d34192b53bbe340
    End
    Caption ="Species Cover"
    DatasheetFontName ="Arial"
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
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =3600
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1320
                    Top =240
                    Width =4575
                    Height =420
                    FontSize =16
                    FontWeight =700
                    Name ="Label3"
                    Caption ="Species Cover by Route"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4440
                    Top =2580
                    Width =1350
                    Height =299
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =2820
                    Top =1080
                    Width =2520
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Park_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.ParkCode, tlu_Parks.ParkName FROM tlu_Parks; "
                    ColumnWidths ="0;2565"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1560
                            Top =1080
                            Width =1140
                            Height =245
                            FontWeight =700
                            Name ="Park_Code_Label"
                            Caption ="Select Park:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =720
                    Left =2820
                    Top =1680
                    Width =1200
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Visit_Year"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_sel_Infest_Year.Visit_Year FROM qry_sel_Infest_Year; "
                    ColumnWidths ="2820"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1380
                            Top =1680
                            Width =1320
                            Height =245
                            FontWeight =700
                            Name ="Plot_ID_Label"
                            Caption ="Select Year:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =2580
                    Width =1350
                    Height =299
                    TabIndex =3
                    Name ="ButtonReport"
                    Caption ="Create Table"
                    OnClick ="[Event Procedure]"

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

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub


Private Sub Park_Code_AfterUpdate()
  If Not IsNull(Me!Park_Code) Then
    Me!Visit_Year.RowSource = "SELECT Distinct Visit_Year FROM qry_sel_cover_Year WHERE [Unit_Code] = '" & Me!Park_Code & "' ORDER BY Visit_Year"
    Me.Refresh
  End If
End Sub
Private Sub ButtonReport_Click()
' Build work table for Species cover by route
' Russ DenBleyker - Northern Colorado Plateau Network - January 2010

  Dim db As DAO.Database
  Dim tdf As TableDef
  Dim WorkOutput As DAO.Recordset
  Dim SpeciesIn As DAO.Recordset
  Dim Transects As DAO.Recordset
  Dim WorkStdDev As DAO.Recordset
  Dim Routes As DAO.Recordset
  Dim strSQL As String
  Dim PlotSave As String
  Dim SpeciesSave As String
  Dim CommonSave As String
  Dim SearchChar As String
  Dim strFieldName As String
  Dim strRouteColumnName As String
  Dim strCountColumnName As String
  Dim strCoverColumnName As String
  Dim strSEColumnName As String
  Dim RouteArray(50, 1) As String  '  Array for route names
  ' column 1 is route name
  ' column 2 is total transect count
  Dim TCount As Variant
  Dim ArrayIndex As Integer
  Dim arrayend As Integer
  Dim EmptyTransects As Integer
  Dim PlotCount As Integer  ' Count of transects in which species was found
  Dim intTextLength As Integer
  Dim CoverSum As Double
  Dim CoverCalc As Double
  Dim varStandardDeviation As Variant
  
   If IsNull(Me!Park_Code) Or IsNull(Me!Visit_Year) Then
     MsgBox "You must select both park and year.", , "Species cover by route"
     Exit Sub
   End If
   
   On Error Resume Next
   DoCmd.DeleteObject acTable, "tbl_wrk_Route_Species"   ' Delete old work table if there was one
   On Error GoTo Err_ButtonReport_Click
   ' Copy template table
   DoCmd.CopyObject , "tbl_wrk_Route_Species", acTable, "tbl_Species_Cover_Template"

  ' Create necessary table fields
   strSQL = "SELECT Plot_ID FROM qry_Group_Cover_Route WHERE Unit_Code= '" & Me!Park_Code & "' AND Visit_Year= " & Me!Visit_Year
   Set db = CurrentDb
   Set Routes = db.OpenRecordset(strSQL)
   Set tdf = db.tabledefs("tbl_wrk_Route_Species")
   ArrayIndex = 0
   Do Until Routes.EOF
     strSQL = "SELECT Count(Transect) AS Transect_Count FROM qry_Group_Route_Transect GROUP BY Unit_Code, Visit_Year, Plot_ID HAVING Unit_Code= '" & Me!Park_Code & "' AND Plot_ID= '" & Routes!Plot_ID & "' AND Visit_Year= " & Me!Visit_Year
     Set Transects = db.OpenRecordset(strSQL)
     TCount = Transects!transect_count  ' save transect count
     Transects.Close
     Set Transects = Nothing
     strRouteColumnName = Left(Routes!Plot_ID, 48) & "(" & TCount & ")"
     strCountColumnName = strRouteColumnName & "PlotCount"
     strCoverColumnName = strRouteColumnName & "CoverPct"
     strSEColumnName = strRouteColumnName & " (SE)"
     With tdf
  '     .Fields.Append .CreateField(strRouteColumnName, dbText, 50)
       .Fields.Append .CreateField(strCountColumnName, dbInteger)
       .Fields.Append .CreateField(strCoverColumnName, dbDouble)
       .Fields.Append .CreateField(strSEColumnName, dbDouble)
     End With
     RouteArray(ArrayIndex, 0) = strRouteColumnName ' Save funky route name
     RouteArray(ArrayIndex, 1) = TCount
     arrayend = ArrayIndex  ' Save last entry index
     ArrayIndex = ArrayIndex + 1
     If ArrayIndex > 49 Then
       MsgBox "Route array overflow - increase array size.", , "Load Route Names"
       Exit Sub
     End If
     Routes.MoveNext
   Loop
   Routes.Close
   Set tdf = Nothing
   Set Routes = Nothing

' calculate species cover by plot
   strSQL = "SELECT * FROM qry_Select_Species_Cover WHERE Unit_Code = '" & Me!Park_Code & "' AND Visit_Year= " & Me!Visit_Year & " ORDER BY Plot_ID, Species"
   Set SpeciesIn = db.OpenRecordset(strSQL)
   SpeciesIn.MoveFirst
   PlotSave = Left(SpeciesIn!Plot_ID, 48)
   SpeciesSave = SpeciesIn!Species
   CommonSave = SpeciesIn!Master_Common_Name
   PlotCount = 0
   CoverCalc = 0
   CoverSum = 0
   SearchChar = "("
   DoCmd.SetWarnings False
   DoCmd.OpenQuery "qry_Clear_StdDev"  ' Clear Standard Deviation work table
   DoCmd.SetWarnings True
   Do Until SpeciesIn.EOF
     If PlotSave <> Left(SpeciesIn!Plot_ID, 48) Or SpeciesSave <> SpeciesIn!Species Then
       ' write output record
       strSQL = "SELECT * FROM tbl_wrk_Route_Species WHERE [Unit_Code]= '" & Me!Park_Code & "' AND [Species] = '" & SpeciesSave & "' AND [Visit_Year] = " & Me!Visit_Year
       Set WorkOutput = db.OpenRecordset(strSQL)
       If WorkOutput.EOF Then
         WorkOutput.Close
         Set WorkOutput = db.OpenRecordset("tbl_wrk_Route_Species")
         WorkOutput.AddNew
         WorkOutput!Unit_Code = Me!Park_Code
         WorkOutput!Visit_Year = Me!Visit_Year
         WorkOutput!Species = SpeciesSave
         WorkOutput!Common_Name = CommonSave
       Else
         WorkOutput.Edit
       End If
         ArrayIndex = 0
         Do Until ArrayIndex > arrayend
           intTextLength = InStr(1, RouteArray(ArrayIndex, 0), SearchChar) - 1
           If Left(RouteArray(ArrayIndex, 0), intTextLength) = PlotSave Then
             strFieldName = RouteArray(ArrayIndex, 0) & "Plotcount"
             WorkOutput(strFieldName) = PlotCount
             strFieldName = RouteArray(ArrayIndex, 0) & "CoverPct"
             WorkOutput(strFieldName) = CoverSum / RouteArray(ArrayIndex, 1)
             ' Standard deviation calculations
             If RouteArray(ArrayIndex, 1) > PlotCount Then
               EmptyTransects = RouteArray(ArrayIndex, 1) - PlotCount  ' calculate number of empty transects
               Set WorkStdDev = db.OpenRecordset("tbl_wrk_StdDev")
               Do Until EmptyTransects = 0  ' add records to StdDev work table for plots in which species was not found
                 WorkStdDev.AddNew
                 WorkStdDev!CoverPct = 0 ' zero cover for these plots
                 WorkStdDev.Update
                 EmptyTransects = EmptyTransects - 1
               Loop
               WorkStdDev.Close
               Set WorkStdDev = Nothing
             End If
             varStandardDeviation = DStDev("CoverPct", "tbl_wrk_StdDev")
             If Not IsNull(varStandardDeviation) Then
               strFieldName = RouteArray(ArrayIndex, 0) & " (SE)"
               ' WorkOutput(strFieldName) = varStandardDeviation / Sqr(PlotCount)  ' Use number of plots in which species is found
               WorkOutput(strFieldName) = varStandardDeviation / Sqr(RouteArray(ArrayIndex, 1))  ' Use total plots in route
             End If
             Exit Do
           End If
           ArrayIndex = ArrayIndex + 1
           If ArrayIndex > arrayend Then
             MsgBox "Name not found in route array", , "Set route name"
             Exit Sub
           End If
         Loop
         WorkOutput.Update
         WorkOutput.Close
         Set WorkOutput = Nothing
       ' Save necessary fields
       PlotSave = Left(SpeciesIn!Plot_ID, 48)
       SpeciesSave = SpeciesIn!Species
       CommonSave = SpeciesIn!Master_Common_Name
       PlotCount = 0
       CoverCalc = 0
       CoverSum = 0
       DoCmd.SetWarnings False
       DoCmd.OpenQuery "qry_Clear_StdDev"  ' Clear Standard Deviation work table
       DoCmd.SetWarnings True
     End If
     PlotCount = PlotCount + 1
     CoverCalc = 0
     Select Case SpeciesIn!Visit_Year  ' put transect average in covercalc
       Case 2008
         If Not IsNull(SpeciesIn!Q1) + IsNull(SpeciesIn!Q2) + IsNull(SpeciesIn!Q3) = -3 Then
           If Not IsNull(SpeciesIn!Q1) Then
             CoverCalc = SpeciesIn!Q1
           End If
           If Not IsNull(SpeciesIn!Q2) Then
             CoverCalc = CoverCalc + SpeciesIn!Q2
           End If
           If Not IsNull(SpeciesIn!Q3) Then
             CoverCalc = CoverCalc + SpeciesIn!Q3
           End If
         End If
       Case 2009
         If Not IsNull(SpeciesIn!Q1_3m) + IsNull(SpeciesIn!Q2_8m) + IsNull(SpeciesIn!Q3_13m) = -3 Then
           If Not IsNull(SpeciesIn!Q1_3m) Then
             CoverCalc = SpeciesIn!Q1_3m
           End If
           If Not IsNull(SpeciesIn!Q2_8m) Then
             CoverCalc = CoverCalc + SpeciesIn!Q2_8m
           End If
           If Not IsNull(SpeciesIn!Q3_13m) Then
             CoverCalc = CoverCalc + SpeciesIn!Q3_13m
           End If
         End If
       Case Else
         If Not IsNull(SpeciesIn!Q1_hm) + IsNull(SpeciesIn!Q2_5m) + IsNull(SpeciesIn!Q3_10m) = -3 Then
           If Not IsNull(SpeciesIn!Q1_hm) Then
             CoverCalc = SpeciesIn!Q1_hm
           End If
           If Not IsNull(SpeciesIn!Q2_5m) Then
             CoverCalc = CoverCalc + SpeciesIn!Q2_5m
           End If
           If Not IsNull(SpeciesIn!Q3_10m) Then
             CoverCalc = CoverCalc + SpeciesIn!Q3_10m
           End If
         End If
     End Select
     CoverSum = CoverSum + (CoverCalc / 3) ' accumulate averages
     Set WorkStdDev = db.OpenRecordset("tbl_wrk_StdDev")  ' save averages for standard deviation calculation
     WorkStdDev.AddNew
     WorkStdDev!CoverPct = (CoverCalc / 3) ' save average for plot in standard deviation work table
     WorkStdDev.Update
     WorkStdDev.Close
     Set WorkStdDev = Nothing
     SpeciesIn.MoveNext
   Loop
     ' write last output record
       strSQL = "SELECT * FROM tbl_wrk_Route_Species WHERE [Unit_Code]= '" & Me!Park_Code & "' AND [Species] = '" & SpeciesSave & "' AND [Visit_Year] = " & Me!Visit_Year
       Set WorkOutput = db.OpenRecordset(strSQL)
       If WorkOutput.EOF Then
         WorkOutput.Close
         Set WorkOutput = db.OpenRecordset("tbl_wrk_Route_Species")
         WorkOutput.AddNew
         WorkOutput!Unit_Code = Me!Park_Code
         WorkOutput!Visit_Year = Me!Visit_Year
         WorkOutput!Species = SpeciesSave
         WorkOutput!Common_Name = CommonSave
       Else
         WorkOutput.Edit
       End If
         ArrayIndex = 0
         Do Until ArrayIndex > arrayend
           intTextLength = InStr(1, RouteArray(ArrayIndex, 0), SearchChar) - 1
           If Left(RouteArray(ArrayIndex, 0), intTextLength) = PlotSave Then
             strFieldName = RouteArray(ArrayIndex, 0) & "Plotcount"
             WorkOutput(strFieldName) = PlotCount
             strFieldName = RouteArray(ArrayIndex, 0) & "CoverPct"
             WorkOutput(strFieldName) = CoverSum / RouteArray(ArrayIndex, 1)
             ' Standard deviation calculations
             If RouteArray(ArrayIndex, 1) > PlotCount Then
               EmptyTransects = RouteArray(ArrayIndex, 1) - PlotCount  ' calculate number of empty transects
               Set WorkStdDev = db.OpenRecordset("tbl_wrk_StdDev")
               Do Until EmptyTransects = 0  ' add records to StdDev work table for plots in which species was not found
                 WorkStdDev.AddNew
                 WorkStdDev!CoverPct = 0 ' zero cover for these plots
                 WorkStdDev.Update
                 EmptyTransects = EmptyTransects - 1
               Loop
               WorkStdDev.Close
               Set WorkStdDev = Nothing
             End If
             varStandardDeviation = DStDev("CoverPct", "tbl_wrk_StdDev")
             If Not IsNull(varStandardDeviation) Then
               strFieldName = RouteArray(ArrayIndex, 0) & " (SE)"
               WorkOutput(strFieldName) = varStandardDeviation / Sqr(RouteArray(ArrayIndex, 1))  ' Use total plots in route
             End If
             Exit Do
           End If
           ArrayIndex = ArrayIndex + 1
           If ArrayIndex > arrayend Then
             MsgBox "Name not found in route array", , "Set route name"
             Exit Sub
           End If
         Loop
         WorkOutput.Update
   SpeciesIn.Close
   Set SpeciesIn = Nothing
   WorkOutput.Close
   Set WorkOutput = Nothing
 '   MsgBox "Finished - results are in tbl_wrk_Route_Species.", , "Species Cover by Route"
   DoCmd.OpenQuery "qry_List_Route_Species"
Exit_ButtonReport_Click:
    Exit Sub

Err_ButtonReport_Click:
    MsgBox Err.Description
    Resume Exit_ButtonReport_Click
    
End Sub
