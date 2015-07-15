Operation =1
Option =0
Where ="(((tbl_Locations.Unit_Code)=[Forms]![frm_Select_Infest_by_Growth]![Park_Code]) A"
    "ND ((Year([Start_Date]))=[Forms]![frm_Select_Infest_by_Growth]![Visit_Year]) AND"
    " ((IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],IIf(["
    "Unit_Code]=\"FOBU\",[WY_Species],[Co_Species]))) Is Not Null))"
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
    Alias ="Species"
    Expression ="IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],IIf([Uni"
        "t_Code]=\"FOBU\",[WY_Species],[Co_Species]))"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
    Expression ="tlu_Size_Class.Size_Class"
    Expression ="tlu_Cover_Class.Cover_Class"
    Expression ="tbl_Infestation.Pulled"
    Expression ="tbl_Infestation.Growth_Stage"
    Expression ="tbl_Infestation.N_Coord"
    Expression ="tbl_Infestation.E_Coord"
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
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="Year([Start_Date])"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0438478aa6d03f4da6530702cba4e1f4
End
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="tbl_Infestation.N_Coord"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Infestation.E_Coord"
        dbInteger "ColumnWidth" ="1170"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Size_Class.Size_Class"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
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
    Right =965
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =94
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =94
        Top =1
        Name ="tbl_Infestation_Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =422
        Bottom =94
        Top =4
        Name ="tbl_Infestation"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =568
        Bottom =94
        Top =12
        Name ="tlu_NCPN_Plants"
        Name =""
    End
    Begin
        Left =606
        Top =6
        Right =724
        Bottom =94
        Top =0
        Name ="tlu_Size_Class"
        Name =""
    End
    Begin
        Left =740
        Top =6
        Right =864
        Bottom =94
        Top =0
        Name ="tlu_Cover_Class"
        Name =""
    End
End
