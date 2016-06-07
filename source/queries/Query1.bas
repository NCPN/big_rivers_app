dbMemo "SQL" ="SELECT DISTINCT tsys_Link_Dbs.[IsODBC], IIf([Type]=4,ParseConnectionStr([Connect"
    "]),\015\012ParseFileName([Database])) AS CurrDb, tsys_Link_Tables.[LinkDb], tsys"
    "_Link_Dbs.Server, MSysObjects.Database AS CurrPath, tsys_Link_Dbs.[FilePath]\015"
    "\012FROM tsys_Link_Dbs INNER JOIN (MSysObjects INNER JOIN tsys_Link_Tables ON MS"
    "ysObjects.Name = tsys_Link_Tables.LinkTable) ON tsys_Link_Dbs.[LinkDb]=tsys_Link"
    "_Tables.[LinkDb]\015\012WHERE (((MSysObjects.Type) In (4,6)) \015\012AND \015\012"
    "((IIf([Type]=4,ParseConnectionStr([Connect]),ParseFileName([Database])))<>tsys_L"
    "ink_Tables.[LinkDb])) \015\012OR (((MSysObjects.Type) IN (4,6))\015\012AND ((IIf"
    "([Type]=4,ParseConnectionStr([Connect],'SERVER=')))<>[Server])) \015\012OR (((MS"
    "ysObjects.Type) IN (4,6)) \015\012AND ((MSysObjects.Database)<>[FilePath])) \015"
    "\012OR (((MSysObjects.Type)=4) \015\012AND ((tsys_Link_Dbs.[IsODBC])=False)) \015"
    "\012OR (((MSysObjects.Type)=6) \015\012AND ((tsys_Link_Dbs.[IsODBC])=True)) \015"
    "\012 OR (((IIf([Type]=4,ParseConnectionStr([Connect],'SERVER='))) IS NULL) \015\012"
    "AND ((tsys_Link_Dbs.Server) IS NOT NULL)) \015\012OR (((IIf([Type]=4,ParseConnec"
    "tionStr([Connect],'SERVER='))) IS NOT NULL) \015\012AND ((tsys_Link_Dbs.Server) "
    "IS NULL))\015\012OR (((MSysObjects.Database) IS NULL));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x851f3a7f29a9ac46932d313f8115a42c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tsys_Link_Dbs.[IsODBC]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrDb"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9c22b65556e8f541966ad0d47aee66bb
        End
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
        dbText "Name" ="CurrPath"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5aa6059b78b98a4a820534f97f6209e9
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.[FilePath]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.IsODBC"
        dbLong "AggregateType" ="-1"
    End
End
