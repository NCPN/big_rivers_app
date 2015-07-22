Operation =1
Option =0
Where ="(((tbl_Locations.Unit_Code)=[Forms]![frm_Select_Infest_Data]![Park_Code]) AND (("
    "Year([Start_Date]))=[Forms]![frm_Select_Infest_Data]![Visit_Year]) AND ((tbl_Loc"
    "ations.Plot_ID) Is Not Null) AND ((IIf([tbl_Locations].[Unit_Code] In (\"CARE\","
    "\"DINO\",\"GOSP\"),[Utah_Species],IIf([tbl_Locations].[Unit_Code]=\"FOBU\",[WY_S"
    "pecies],[Co_Species]))) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Infestation_Events"
    Name ="tbl_Infestation"
    Name ="tlu_NCPN_Plants"
    Name ="tlu_Size_Class"
    Name ="tlu_Cover_Class"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_Locations.Plot_ID"
    Alias ="Species"
    Expression ="IIf(tbl_Locations.Unit_Code In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Speci"
        "es],IIf(tbl_Locations.Unit_Code=\"FOBU\",[WY_Species],[Co_Species]))"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
    Expression ="tbl_Infestation.Pulled"
    Expression ="tbl_Infestation.Growth_Stage"
    Expression ="tbl_Infestation.N_Coord"
    Expression ="tbl_Infestation.E_Coord"
    Expression ="tlu_Cover_Class.Cover_Class"
    Expression ="tlu_Size_Class.Size_Class"
End
Begin Joins
    LeftTable ="tbl_Infestation"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_Infestation.Master_Code = tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
    LeftTable ="tbl_Infestation"
    RightTable ="tlu_Size_Class"
    Expression ="tbl_Infestation.Size_Text = tlu_Size_Class.Size_Description"
    Flag =2
    LeftTable ="tbl_Infestation"
    RightTable ="tlu_Cover_Class"
    Expression ="tbl_Infestation.Cover_Text = tlu_Cover_Class.Cover_Description"
    Flag =2
    LeftTable ="tbl_Infestation_Events"
    RightTable ="tbl_Infestation"
    Expression ="tbl_Infestation_Events.Infest_Event_ID = tbl_Infestation.Infest_Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Infestation_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Infestation_Events.Location_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="IIf(tbl_Locations.Unit_Code In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Speci"
        "es],IIf(tbl_Locations.Unit_Code=\"FOBU\",[WY_Species],[Co_Species]))"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5fe26818f9b0134d97c17faba6d2cd65
End
Begin
End
Begin
    State =0
    Left =98
    Top =338
    Right =1045
    Bottom =662
    Left =-1
    Top =-1
    Right =928
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =26
        Top =6
        Right =134
        Bottom =94
        Top =1
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =330
        Bottom =94
        Top =1
        Name ="tbl_Infestation_Events"
        Name =""
    End
    Begin
        Left =371
        Top =6
        Right =495
        Bottom =94
        Top =3
        Name ="tbl_Infestation"
        Name =""
    End
    Begin
        Left =542
        Top =6
        Right =667
        Bottom =94
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
    Begin
        Left =807
        Top =7
        Right =903
        Bottom =95
        Top =0
        Name ="tlu_Size_Class"
        Name =""
    End
    Begin
        Left =691
        Top =8
        Right =787
        Bottom =96
        Top =0
        Name ="tlu_Cover_Class"
        Name =""
    End
End
