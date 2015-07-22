Operation =1
Option =0
Where ="(((tbl_Locations.Plot_ID) Is Not Null) And ((IIf(tbl_Locations.Unit_Code In (\"C"
    "ARE\",\"DINO\",\"GOSP\"),[Utah_Species],IIf(tbl_Locations.Unit_Code=\"FOBU\",[WY"
    "_Species],[Co_Species]))) Is Not Null And (IIf(tbl_Locations.Unit_Code In (\"CAR"
    "E\",\"DINO\",\"GOSP\"),[Utah_Species],IIf(tbl_Locations.Unit_Code=\"FOBU\",[WY_S"
    "pecies],[Co_Species]))) Is Not Null And (IIf(tbl_Locations.Unit_Code In (\"CARE\""
    ",\"DINO\",\"GOSP\"),[Utah_Species],IIf(tbl_Locations.Unit_Code=\"FOBU\",[WY_Spec"
    "ies],[Co_Species]))) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Infestation_Events"
    Name ="tbl_Infestation"
    Name ="tlu_NCPN_Plants"
    Name ="tlu_Size_Class"
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
    Expression ="tlu_Size_Class.Size_Class"
    Expression ="tbl_Infestation.Master_Code"
End
Begin Joins
    LeftTable ="tbl_Infestation"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_Infestation.Master_Code=tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
    LeftTable ="tbl_Infestation"
    RightTable ="tlu_Size_Class"
    Expression ="tbl_Infestation.Size_Text=tlu_Size_Class.Size_Description"
    Flag =2
    LeftTable ="tbl_Infestation_Events"
    RightTable ="tbl_Infestation"
    Expression ="tbl_Infestation_Events.Infest_Event_ID=tbl_Infestation.Infest_Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Infestation_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Infestation_Events.Location_ID"
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
    0x321f4c0b75f4ec4daffa9e99c41431dd
End
Begin
    Begin
        dbText "Name" ="Visit_Year"
        dbBinary "GUID" = Begin
            0x0f58ea459ba6e64393ec44ca0665ca23
        End
    End
    Begin
        dbText "Name" ="Species"
        dbBinary "GUID" = Begin
            0xe28d3a854a833d439409de82071d32e2
        End
    End
End
Begin
    State =0
    Left =19
    Top =13
    Right =1165
    Bottom =337
    Left =-1
    Top =-1
    Right =1131
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =13
        Top =5
        Right =121
        Bottom =108
        Top =1
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =151
        Top =6
        Right =309
        Bottom =109
        Top =2
        Name ="tbl_Infestation_Events"
        Name =""
    End
    Begin
        Left =339
        Top =6
        Right =463
        Bottom =109
        Top =3
        Name ="tbl_Infestation"
        Name =""
    End
    Begin
        Left =493
        Top =6
        Right =618
        Bottom =109
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
    Begin
        Left =650
        Top =9
        Right =746
        Bottom =97
        Top =0
        Name ="tlu_Size_Class"
        Name =""
    End
End
