Operation =3
Name ="tsys_Link_dbs"
Option =0
Where ="(((tsys_Link_Dbs.Link_db) Is Null))"
Begin InputTables
    Name ="qsys_Linked_tables_not_in_tsys_Link_Tables"
    Name ="tsys_Link_Dbs"
End
Begin OutputColumns
    Name ="Link_db"
    Expression ="qsys_Linked_tables_not_in_tsys_Link_Tables.CurrDb"
    Name ="Server"
    Expression ="qsys_Linked_tables_not_in_tsys_Link_Tables.CurrServer"
    Name ="File_path"
    Expression ="qsys_Linked_tables_not_in_tsys_Link_Tables.CurrPath"
    Name ="Is_ODBC"
    Expression ="qsys_Linked_tables_not_in_tsys_Link_Tables.ODBC"
    Alias ="Backup"
    Name ="Backups"
    Expression ="Not ([ODBC])"
End
Begin Joins
    LeftTable ="qsys_Linked_tables_not_in_tsys_Link_Tables"
    RightTable ="tsys_Link_Dbs"
    Expression ="qsys_Linked_tables_not_in_tsys_Link_Tables.CurrDb = tsys_Link_Dbs.Link_db"
    Flag =2
End
Begin Groups
    Expression ="qsys_Linked_tables_not_in_tsys_Link_Tables.CurrDb"
    GroupLevel =0
    Expression ="qsys_Linked_tables_not_in_tsys_Link_Tables.CurrServer"
    GroupLevel =0
    Expression ="qsys_Linked_tables_not_in_tsys_Link_Tables.CurrPath"
    GroupLevel =0
    Expression ="qsys_Linked_tables_not_in_tsys_Link_Tables.ODBC"
    GroupLevel =0
    Expression ="Not ([ODBC])"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbText "Description" ="Automatically appends back-end databases to tsys_Link_dbs if a record is missing"
dbBinary "GUID" = Begin
    0xe191a4306c92d9498a92050a65983d52
End
Begin
    Begin
        dbText "Name" ="Backup"
        dbBinary "GUID" = Begin
            0x7c8c2c7eb7a1c3439d93736ddbde056e
        End
    End
End
Begin
    State =0
    Left =18
    Top =40
    Right =1130
    Bottom =352
    Left =-1
    Top =-1
    Right =1074
    Bottom =123
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =38
        Top =6
        Right =326
        Bottom =113
        Top =0
        Name ="qsys_Linked_tables_not_in_tsys_Link_Tables"
        Name =""
    End
    Begin
        Left =364
        Top =6
        Right =460
        Bottom =113
        Top =0
        Name ="tsys_Link_Dbs"
        Name =""
    End
End
