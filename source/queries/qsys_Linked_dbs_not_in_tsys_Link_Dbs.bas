dbMemo "SQL" ="INSERT INTO tsys_Link_dbs ( LinkDb, Server, FilePath, IsODBC, Backups )\015\012S"
    "ELECT qsys_Linked_tables_not_in_tsys_Link_Tables.CurrDb, qsys_Linked_tables_not_"
    "in_tsys_Link_Tables.CurrServer, qsys_Linked_tables_not_in_tsys_Link_Tables.CurrP"
    "ath, qsys_Linked_tables_not_in_tsys_Link_Tables.ODBC, Not ([ODBC]) AS Backup\015"
    "\012FROM qsys_Linked_tables_not_in_tsys_Link_Tables LEFT JOIN tsys_Link_Dbs ON q"
    "sys_Linked_tables_not_in_tsys_Link_Tables.CurrDb=tsys_Link_Dbs.[LinkDb]\015\012W"
    "HERE (((tsys_Link_Dbs.[LinkDb]) Is Null))\015\012GROUP BY qsys_Linked_tables_not"
    "_in_tsys_Link_Tables.CurrDb, qsys_Linked_tables_not_in_tsys_Link_Tables.CurrServ"
    "er, qsys_Linked_tables_not_in_tsys_Link_Tables.CurrPath, qsys_Linked_tables_not_"
    "in_tsys_Link_Tables.ODBC, Not ([ODBC]);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbText "Description" ="Automatically appends back-end databases to tsys_Link_dbs if a record is missing"
dbBinary "GUID" = Begin
    0xe48054c9e4c2214392d20ea121a42928
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Backup"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7c8c2c7eb7a1c3439d93736ddbde056e
        End
    End
End
