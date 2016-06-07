dbMemo "SQL" ="SELECT tsys_Link_Tables.LinkTable, tsys_Link_Tables.[LinkDb], tsys_Link_Dbs.Serv"
    "er, tsys_Link_Dbs.[FilePath]\015\012FROM tsys_Link_Dbs INNER JOIN (MSysObjects R"
    "IGHT JOIN tsys_Link_Tables ON MSysObjects.Name = tsys_Link_Tables.LinkTable) ON "
    "tsys_Link_Dbs.[LinkDb]=tsys_Link_Tables.[LinkDb]\015\012WHERE (((MSysObjects.Nam"
    "e) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Linked table records in tsys_Link_Tables that are not actually in the database"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x292357a403aa1741891cfc39a08883e1
End
Begin
    Begin
        dbText "Name" ="tsys_Link_Tables.LinkTable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Tables.[LinkDb]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.Server"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.[FilePath]"
        dbLong "AggregateType" ="-1"
    End
End
