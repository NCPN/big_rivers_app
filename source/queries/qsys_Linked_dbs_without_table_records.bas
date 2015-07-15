Operation =1
Option =0
Where ="(((tsys_Link_Tables.Link_db) Is Null))"
Begin InputTables
    Name ="tsys_Link_Dbs"
    Name ="tsys_Link_Tables"
End
Begin OutputColumns
    Expression ="tsys_Link_Dbs.*"
End
Begin Joins
    LeftTable ="tsys_Link_Dbs"
    RightTable ="tsys_Link_Tables"
    Expression ="tsys_Link_Dbs.Link_db = tsys_Link_Tables.Link_db"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Records in tsys_Link_Dbs that do not have any related child records in tsys_Link"
    "_Tables"
dbBinary "GUID" = Begin
    0x17c9f3e1e976ed4a8b1214e515bbd4ab
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =62
    Top =84
    Right =1130
    Bottom =396
    Left =-1
    Top =-1
    Right =1030
    Bottom =123
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="tsys_Link_Dbs"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="tsys_Link_Tables"
        Name =""
    End
End
