Operation =1
Option =0
Having ="(((tbl_Locations.Unit_Code)=[Forms]![frm_Infest_by_GS_Count]![Park_Code]) AND (("
    "Year([Start_Date]))=[Forms]![frm_Infest_by_GS_Count]![Visit_Year]) AND ((tbl_Inf"
    "estation.Growth_Stage)<>\"\"))"
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
    Expression ="tbl_Infestation.Growth_Stage"
    Alias ="Infestation Count"
    Expression ="Count(tbl_Infestation.Infestation_ID)"
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
    Expression ="tbl_Infestation.Growth_Stage"
    Flag =0
End
Begin Groups
    Expression ="tbl_Locations.Unit_Code"
    GroupLevel =0
    Expression ="Year([Start_Date])"
    GroupLevel =0
    Expression ="tbl_Infestation.Growth_Stage"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBinary "GUID" = Begin
    0xd93754283ce00149907a93fda609e8eb
End
Begin
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
    Begin
        dbText "Name" ="tbl_Infestation.Growth_Stage"
        dbInteger "ColumnWidth" ="1425"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Infestation Count"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =26
    Top =245
    Right =1010
    Bottom =569
    Left =-1
    Top =-1
    Right =965
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =109
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =109
        Top =1
        Name ="tbl_Infestation_Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =422
        Bottom =109
        Top =0
        Name ="tbl_Infestation"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =568
        Bottom =109
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
