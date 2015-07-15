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
    Left =4350
    Top =225
    Right =11295
    Bottom =3555
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3d34192b53bbe340
    End
    Caption ="Infestations by Size Class"
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
                    Left =1425
                    Top =240
                    Width =4365
                    Height =420
                    FontSize =16
                    FontWeight =700
                    Name ="Label3"
                    Caption ="Infestations by Size"
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
    Dim stWhere As String
    Dim stOpenArg As String
    Dim strSQL As String
    Dim strQueryName As String
    Dim db As DAO.Database
    Dim WorkOutput As DAO.Recordset
    Dim Infest As DAO.Recordset
    Dim SpeciesSave As String
    Dim ClassName As String
    Dim NameSave As String
    Dim InfestSum As Integer
    Dim PrioritySum As Integer
    Dim ArrayIndex As Integer
    Dim SizeArray(5) As Integer
  ' Array for the four size classes

    If IsNull(Me!Park_Code) Or IsNull(Me!Visit_Year) Then
      MsgBox "You must select both park and year.", , "Infestation by Route"
      Exit Sub
    End If

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Infest_Size"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_Infest_Size WHERE Unit_Code = '" & Me!Park_Code & "' AND Visit_Year = " & Me!Visit_Year
  strSQL = strSQL & " AND [Species] Is Not Null"
  strSQL = strSQL & " AND [Species] <> ''"
  strSQL = strSQL & " ORDER BY Species"
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
   SpeciesSave = Infest!Species     ' Save necessary fields
   NameSave = Infest!Master_Common_Name
   ArrayIndex = 0
   Do Until ArrayIndex > 4
     SizeArray(ArrayIndex) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Infest_Size")
   Do Until Infest.EOF
     If SpeciesSave <> Infest!Species Then  ' New plot code
       WorkOutput.AddNew
       WorkOutput!UnitCode = Me!Park_Code
       WorkOutput!Species = SpeciesSave  ' Set species
       WorkOutput!VisitYear = Me!Visit_Year  ' Set visit date
       WorkOutput!CommonName = NameSave
       WorkOutput!InfestTot = InfestSum
       WorkOutput!PriorityTot = PrioritySum
       ArrayIndex = 0
       Do Until ArrayIndex > 4
         ClassName = "Class" & (ArrayIndex + 1) ' Set the class size field name
         WorkOutput(ClassName) = SizeArray(ArrayIndex)
         SizeArray(ArrayIndex) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write species record
       InfestSum = 0
       PrioritySum = 0
       SpeciesSave = Infest!Species     ' Save necessary fields
       NameSave = Infest!Master_Common_Name
     End If  ' End if for new species compare
     If Infest!Priority = 1 Then
       PrioritySum = PrioritySum + 1
     End If
     InfestSum = InfestSum + 1
     If IsNumeric(Infest!Size_Class) Then
       SizeArray((Infest!Size_Class - 1)) = SizeArray((Infest!Size_Class - 1)) + 1
     End If
     Infest.MoveNext
   Loop
   WorkOutput.AddNew   ' Write last record
       WorkOutput!UnitCode = Me!Park_Code
       WorkOutput!Species = SpeciesSave  ' Set species
       WorkOutput!VisitYear = Me!Visit_Year  ' Set visit date
       WorkOutput!CommonName = NameSave
       WorkOutput!InfestTot = InfestSum
       WorkOutput!PriorityTot = PrioritySum
       ArrayIndex = 0
       Do Until ArrayIndex > 4
         ClassName = "Class" & (ArrayIndex + 1) ' Set the class size field name
         WorkOutput(ClassName) = SizeArray(ArrayIndex)
         ArrayIndex = ArrayIndex + 1
       Loop
   WorkOutput.Update  ' Write plot record
   Set WorkOutput = Nothing
   Infest.Close
   Set Infest = Nothing
    stOpenArg = Me!Park_Code & Me!Visit_Year
    stDocName = "rpt_Infest_by_Size"
    DoCmd.OpenReport stDocName, acPreview, , , , stOpenArg
Exit_Infest_Click:
   Exit Sub

Err_Infest_Click:
    MsgBox Err.Description
    Resume Exit_Infest_Click
   
End Sub
