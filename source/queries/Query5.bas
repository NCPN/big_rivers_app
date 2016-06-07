dbMemo "SQL" ="SELECT DISTINCT tsys_Link_Dbs.[IsODBC], IIf([Type]=4,ParseConnectionStr([Connect"
    "]),\015\012ParseFileName([Database])) AS CurrDb, tsys_Link_Tables.[LinkDb], tsys"
    "_Link_Dbs.Server, MSysObjects.Database AS CurrPath, tsys_Link_Dbs.[FilePath]\015"
    "\012FROM tsys_Link_Dbs INNER JOIN (MSysObjects INNER JOIN tsys_Link_Tables ON MS"
    "ysObjects.Name = tsys_Link_Tables.LinkTable) ON tsys_Link_Dbs.[LinkDb]=tsys_Link"
    "_Tables.[LinkDb]\015\012WHERE (MSysObjects.Database) IS NULL;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x25811d462885bc4aad7624b515855445
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="CurrDb"
        dbBinary "GUID" = Begin
            0x1d6334f8b432aa43be5b7e55da073fd3
        End
    End
    Begin
        dbText "Name" ="CurrPath"
        dbBinary "GUID" = Begin
            0x058ef6460b0ad243b200e1295c63b09c
        End
    End
End
