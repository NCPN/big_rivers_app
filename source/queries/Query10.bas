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
    0xebd84079f8512b488d20cb5c9d81524d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="CurrDb"
        dbBinary "GUID" = Begin
            0x5f7a6399807cc54b87e907424017c26d
        End
    End
    Begin
        dbText "Name" ="CurrPath"
        dbBinary "GUID" = Begin
            0x92be31816e541a4faac331a191d39be6
        End
    End
End
