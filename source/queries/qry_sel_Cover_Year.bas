Operation =1
Option =0
Begin InputTables
    Name ="qry_Select_Species_Cover"
End
Begin OutputColumns
    Expression ="qry_Select_Species_Cover.Unit_Code"
    Expression ="qry_Select_Species_Cover.Visit_Year"
End
Begin Groups
    Expression ="qry_Select_Species_Cover.Unit_Code"
    GroupLevel =0
    Expression ="qry_Select_Species_Cover.Visit_Year"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x24b249451e16af439245930410a20a6d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qry_Select_Species_Cover.Unit_Code"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Select_Species_Cover.Visit_Year"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =18
    Top =14
    Right =1002
    Bottom =327
    Left =-1
    Top =-1
    Right =946
    Bottom =123
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="qry_Select_Species_Cover"
        Name =""
    End
End
