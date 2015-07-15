Operation =1
Option =0
Where ="(((tbl_Quadrat_Species.Plant_Code) Is Not Null And (tbl_Quadrat_Species.Plant_Co"
    "de)<>\"none\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Transect"
    Name ="tbl_Quadrat_Species"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Quadrat_Species.Plant_Code"
    Expression ="tbl_Quadrat_Species.Q1_hm"
    Expression ="tbl_Quadrat_Species.Q2_5m"
    Expression ="tbl_Quadrat_Species.Q3_10m"
    Expression ="tbl_Quadrat_Species.Average_Cover"
    Expression ="tbl_Quadrat_Species.Q1_3m"
    Expression ="tbl_Quadrat_Species.Q2_8m"
    Expression ="tbl_Quadrat_Species.Q3_13m"
    Expression ="tbl_Quadrat_Species.Avg_Cover_2009"
    Expression ="tbl_Quadrat_Species.Q1"
    Expression ="tbl_Quadrat_Species.Q2"
    Expression ="tbl_Quadrat_Species.Q3"
    Expression ="tbl_Quadrat_Species.Avg_Cover_2008"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
    Alias ="Species"
    Expression ="IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],IIf([Uni"
        "t_Code]=\"FOBU\",[WY_Species],[Co_Species]))"
    Expression ="tbl_Quadrat_Transect.Transect"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_Quadrat_Transect"
    Expression ="tbl_Events.Event_ID=tbl_Quadrat_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_Quadrat_Species"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_Quadrat_Species.Plant_Code=tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
    LeftTable ="tbl_Quadrat_Transect"
    RightTable ="tbl_Quadrat_Species"
    Expression ="tbl_Quadrat_Transect.Transect_ID=tbl_Quadrat_Species.Transect_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="Year([Start_Date])"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],IIf([Uni"
        "t_Code]=\"FOBU\",[WY_Species],[Co_Species]))"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4c0a29541bd6cc45842afa8f071e4ae4
End
Begin
    Begin
        dbText "Name" ="Visit_Year"
        dbBinary "GUID" = Begin
            0x33327f1e0d34e349bb954ca4fff2b741
        End
    End
    Begin
        dbText "Name" ="Species"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbBinary "GUID" = Begin
            0x7170fd28524c93479120d0aa8589532a
        End
    End
End
Begin
    State =0
    Left =18
    Top =14
    Right =1002
    Bottom =338
    Left =-1
    Top =-1
    Right =969
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =109
        Top =1
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =109
        Top =4
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =109
        Top =0
        Name ="tbl_Quadrat_Transect"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =109
        Top =0
        Name ="tbl_Quadrat_Species"
        Name =""
    End
    Begin
        Left =574
        Top =6
        Right =731
        Bottom =109
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
