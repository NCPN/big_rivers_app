dbMemo "SQL" ="SELECT DISTINCT tsys_Link_Dbs.IsODBC, IIf([Type]=4,\015\012ParseConnectionStr([C"
    "onnect]),\015\012ParseFileName([Database])) AS CurrDb, tsys_Link_Tables.LinkDb, "
    "tsys_Link_Dbs.Server, MSysObjects.Database AS CurrPath, tsys_Link_Dbs.FilePath\015"
    "\012FROM tsys_Link_Dbs INNER JOIN (MSysObjects INNER JOIN tsys_Link_Tables ON MS"
    "ysObjects.Name = tsys_Link_Tables.LinkTable) ON tsys_Link_Dbs.LinkDb = tsys_Link"
    "_Tables.LinkDb\015\012WHERE (((MSysObjects.Type) In (4,6)) \015\012AND \015\012("
    "(IIf([Type]=4,ParseConnectionStr([Connect]),\015\012ParseFileName([Database])))<"
    ">tsys_Link_Tables.LinkDb)) \015\012OR \015\012(((MSysObjects.Type) In (4,6)) \015"
    "\012AND \015\012((IIf([Type]=4,ParseConnectionStr([Connect],'SERVER=')))<>[Serve"
    "r])) \015\012OR \015\012(((MSysObjects.Type) In (4,6)) \015\012AND ((MSysObjects"
    ".Database)<>[FilePath])) \015\012OR (((MSysObjects.Type)=4) \015\012AND ((tsys_L"
    "ink_Dbs.IsODBC)=False)) \015\012OR (((MSysObjects.Type)=6) \015\012AND ((tsys_Li"
    "nk_Dbs.IsODBC)=True)) \015\012OR (((IIf([Type]=4,ParseConnectionStr([Connect],'S"
    "ERVER='))) Is Null) AND ((tsys_Link_Dbs.Server) Is Not Null)) \015\012OR (((IIf("
    "[Type]=4,ParseConnectionStr([Connect],'SERVER='))) Is Not Null) \015\012AND ((ts"
    "ys_Link_Dbs.Server) Is Null)) \015\012OR (((MSysObjects.Database) Is Null) \015\012"
    "AND ((tsys_Link_Dbs.FilePath) Is Not Null)) \015\012OR (((MSysObjects.Database) "
    "Is Not Null) \015\012AND ((tsys_Link_Dbs.FilePath) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x114bff8597c3f74eae1281e6129bb5d3
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tsys_Link_Dbs.IsODBC"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc8201d3d0fe699459159920eaf7a33bf
        End
    End
    Begin
        dbText "Name" ="CurrDb"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4011347f91ce24468591dfaac6d303bf
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Tables.LinkDb"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3e4012ae5a7a7d449a8e1ec582483e8e
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.Server"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x72ba66257ae6a14fb339221eea6e5daf
        End
    End
    Begin
        dbText "Name" ="CurrPath"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0f771060b64593428aa19c54329eea7c
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.FilePath"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeababdd0dc062045919c61033e653895
        End
    End
End
