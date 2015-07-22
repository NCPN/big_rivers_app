Operation =1
Option =0
Where ="(((MSysObjects.Name) Is Null))"
Begin InputTables
    Name ="MSysObjects"
    Name ="tsys_Link_Tables"
    Name ="tsys_Link_Dbs"
End
Begin OutputColumns
    Expression ="tsys_Link_Tables.Link_table"
    Expression ="tsys_Link_Tables.Link_db"
    Expression ="tsys_Link_Dbs.Server"
    Expression ="tsys_Link_Dbs.File_path"
End
Begin Joins
    LeftTable ="MSysObjects"
    RightTable ="tsys_Link_Tables"
    Expression ="MSysObjects.Name = tsys_Link_Tables.Link_table"
    Flag =3
    LeftTable ="tsys_Link_Dbs"
    RightTable ="tsys_Link_Tables"
    Expression ="tsys_Link_Dbs.Link_db = tsys_Link_Tables.Link_db"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Linked table records in tsys_Link_Tables that are not actually in the database"
dbBinary "GUID" = Begin
    0x1a459581345955409cd444a8e20a1d61
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =0
    Top =40
    Right =978
    Bottom =649
    Left =-1
    Top =-1
    Right =940
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
        Name ="MSysObjects"
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
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =113
        Top =0
        Name ="tsys_Link_Dbs"
        Name =""
    End
End
