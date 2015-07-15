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
    Left =3480
    Top =2370
    Right =10425
    Bottom =5700
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3d34192b53bbe340
    End
    Caption ="Infestations by Route"
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
                    Left =1410
                    Top =240
                    Width =4395
                    Height =420
                    FontSize =16
                    FontWeight =700
                    Name ="Label3"
                    Caption ="Infestations by Route"
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
                    Caption ="Preview Report"
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
    Me!Visit_Year.RowSource = "SELECT Visit_Year FROM qry_sel_Infest_Year WHERE [Unit_Code] = '" & Me!Park_Code & "' ORDER BY Visit_Year"
    Me.Refresh
  End If
End Sub
Private Sub ButtonReport_Click()
On Error GoTo Err_Infest_Click

    Dim stDocName As String
    Dim stOpenArg As String
    Dim strSQL As String
    Dim db As DAO.Database
    Dim WorkOutput As DAO.Recordset
    Dim Infest As DAO.Recordset
    Dim RouteStats As DAO.Recordset
    Dim PlotSave As String
    Dim AreaSave As Double
    Dim InfestSum As Integer
    Dim PrioritySum As Integer
    Dim Priority As Variant

    If IsNull(Me!Park_Code) Or IsNull(Me!Visit_Year) Then
      MsgBox "You must select both park and year.", , "Infestation by Route"
      Exit Sub
    End If

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Infest_Route"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Infest_by_Route where Unit_Code = '" & Me!Park_Code & "' AND Visit_Year = " & Me!Visit_Year
  Set db = CurrentDb

  ' Get first infestation record
   Set Infest = db.OpenRecordset(strSQL)
   If Infest.EOF Then
     MsgBox "No valid infestation records found."
     Infest.Close
     Set Infest = Nothing
     GoTo Exit_Infest_Click
   End If
   InfestSum = 0
   PrioritySum = 0
   Infest.MoveFirst
   PlotSave = Infest!Plot_ID     ' Save necessary fields
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Infest_Route")
   Do Until Infest.EOF
     If PlotSave <> Infest!Plot_ID Then  ' New plot code
       WorkOutput.AddNew
       WorkOutput!UnitCode = Me!Park_Code
       WorkOutput!PlotID = PlotSave  ' Set route
       WorkOutput!VisitYear = Me!Visit_Year  ' Set visit date
       WorkOutput!InfestTot = InfestSum
       WorkOutput!PriorityTot = PrioritySum
       strSQL = "SELECT * FROM tlu_Route_Statistics where Unit_Code = '" & Me!Park_Code & "' AND Route = '" & PlotSave & "' AND Visit_Year = " & Me!Visit_Year
       Set RouteStats = db.OpenRecordset(strSQL)
       If RouteStats.EOF Then
         WorkOutput!RouteType = "Not on Route"
         AreaSave = 0
       Else
         If RouteStats!Route_Length_k > 0 Then
           WorkOutput!RouteLength = (RouteStats!Route_Length_k * 1000)
         End If
         If Not IsNull(RouteStats!Route_Type) Then
           WorkOutput!RouteType = RouteStats!Route_Type
         End If
         AreaSave = RouteStats!Route_Area_h
       End If
       RouteStats.Close
       Set RouteStats = Nothing
       If AreaSave > 0 Then
         ' AreaSave = AreaSave * 100  ' convert hectares to 100 m2
         WorkOutput!TotPct = InfestSum / AreaSave
         WorkOutput!PriorityPct = PrioritySum / AreaSave
       End If
       WorkOutput!RouteArea = AreaSave
       WorkOutput.Update  ' Write route record
       
       InfestSum = 0
       PrioritySum = 0
       PlotSave = Infest!Plot_ID     ' Save necessary fields
     End If  ' End if for new route compare
     If Not IsNull(Infest!Master_Code) Then
       InfestSum = InfestSum + 1
       Priority = DLookup("[Priority]", "tbl_Target_Plant_lists", "[Unit_Code]= '" & Me!Park_Code & "' AND [Master_Plant_Code] = '" & Infest!Master_Code & "' AND [Visit_Year] = " & Me!Visit_Year)
       '   Priority = DLookup("[Priority]", "tbl_Target_Plant_lists", "[Unit_Code]= '" & Me!Park_Code & "' AND [Master_Plant_Code] = '" & CodeSave & "' AND [Visit_Year] = " & 2008)   ' For testing purposes
       If Not IsNull(Priority) And Priority = 1 Then
         PrioritySum = PrioritySum + 1
       End If
     End If
     Infest.MoveNext
   Loop
   WorkOutput.AddNew   ' Write last record
   WorkOutput!UnitCode = Me!Park_Code
   WorkOutput!PlotID = PlotSave  ' Set route
   WorkOutput!VisitYear = Me!Visit_Year  ' Set visit date
   WorkOutput!InfestTot = InfestSum
   WorkOutput!PriorityTot = PrioritySum
       strSQL = "SELECT * FROM tlu_Route_Statistics where Unit_Code = '" & Me!Park_Code & "' AND Route = '" & PlotSave & "' AND Visit_Year = " & Me!Visit_Year
       Set RouteStats = db.OpenRecordset(strSQL)
       If RouteStats.EOF Then
         WorkOutput!RouteType = "Not on Route"
         AreaSave = 0
       Else
         If RouteStats!Route_Length_k > 0 Then
           WorkOutput!RouteLength = RouteStats!Route_Length_k * 1000
         End If
         If Not IsNull(RouteStats!Route_Type) Then
           WorkOutput!RouteType = RouteStats!Route_Type
         End If
         AreaSave = RouteStats!Route_Area_h
       End If
       RouteStats.Close
       Set RouteStats = Nothing
       If AreaSave > 0 Then
      '   AreaSave = AreaSave * 100  ' convert hectares to 100 m2
         WorkOutput!TotPct = InfestSum / AreaSave
         WorkOutput!PriorityPct = PrioritySum / AreaSave
       End If
       WorkOutput!RouteArea = AreaSave
   WorkOutput.Update  ' Write plot record
   Set WorkOutput = Nothing
   Infest.Close
   Set Infest = Nothing
    stOpenArg = Me!Park_Code & Me!Visit_Year
    stDocName = "rpt_Infest_by_Route"
    DoCmd.OpenReport stDocName, acPreview, , , , stOpenArg
Exit_Infest_Click:
   Exit Sub

Err_Infest_Click:
    MsgBox Err.Description
    Resume Exit_Infest_Click
End Sub
