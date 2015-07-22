Operation =1
Option =0
Where ="(((Year([Start_Date])) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Infestation_Events"
    Name ="tbl_Infestation"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Infestation.Master_Code"
End
Begin Joins
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
    Expression ="tbl_Infestation.Master_Code"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc958b026daece045895af4ae8ad59335
End
Begin
    Begin
        dbText "Name" ="tbl_Infestation.Master_Code"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =15
    Top =98
    Right =922
    Bottom =422
    Left =-1
    Top =-1
    Right =888
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =109
        Top =4
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =345
        Bottom =109
        Top =0
        Name ="tbl_Infestation_Events"
        Name =""
    End
    Begin
        Left =383
        Top =6
        Right =561
        Bottom =109
        Top =0
        Name ="tbl_Infestation"
        Name =""
    End
End
