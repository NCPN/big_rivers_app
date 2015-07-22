Operation =1
Option =0
Begin InputTables
    Name ="qry_Select_Species_Cover"
End
Begin OutputColumns
    Expression ="qry_Select_Species_Cover.Unit_Code"
    Expression ="qry_Select_Species_Cover.Visit_Year"
    Expression ="qry_Select_Species_Cover.Plot_ID"
End
Begin Groups
    Expression ="qry_Select_Species_Cover.Unit_Code"
    GroupLevel =0
    Expression ="qry_Select_Species_Cover.Visit_Year"
    GroupLevel =0
    Expression ="qry_Select_Species_Cover.Plot_ID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe7e1cea3eab3a14490a0e980729682af
End
Begin
End
Begin
    State =0
    Left =5
    Top =97
    Right =989
    Bottom =421
    Left =-1
    Top =-1
    Right =969
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =109
        Top =15
        Name ="qry_Select_Species_Cover"
        Name =""
    End
End
