dbMemo "SQL" ="SELECT DISTINCT tsys_Link_Dbs.IsODBC, IIf([Type]=4,ParseConnectionStr([Connect])"
    ",ParseFileName([Database])) AS CurrDb, tsys_Link_Tables.LinkDb, tsys_Link_Dbs.Se"
    "rver, MSysObjects.Database AS CurrPath, tsys_Link_Dbs.FilePath\015\012FROM tsys_"
    "Link_Dbs INNER JOIN (MSysObjects INNER JOIN tsys_Link_Tables ON MSysObjects.Name"
    " = tsys_Link_Tables.LinkTable) ON tsys_Link_Dbs.LinkDb = tsys_Link_Tables.LinkDb"
    "\015\012WHERE (((MSysObjects.Type) In (4,6)) And ((IIf([Type]=4,ParseConnectionS"
    "tr([Connect]),ParseFileName([Database])))<>tsys_Link_Tables.LinkDb)) Or (((MSysO"
    "bjects.Type) In (4,6)) And ((IIf([Type]=4,ParseConnectionStr([Connect],'SERVER='"
    ")))<>[Server])) Or (((MSysObjects.Type) In (4,6)) And ((MSysObjects.Database)<>["
    "FilePath])) Or (((MSysObjects.Type)=4) And ((tsys_Link_Dbs.IsODBC)=False)) Or (("
    "(MSysObjects.Type)=6) And ((tsys_Link_Dbs.IsODBC)=True)) Or (((IIf([Type]=4,Pars"
    "eConnectionStr([Connect],'SERVER='))) Is Null) And ((tsys_Link_Dbs.Server) Is No"
    "t Null)) Or (((IIf([Type]=4,ParseConnectionStr([Connect],'SERVER='))) Is Not Nul"
    "l) And ((tsys_Link_Dbs.Server) Is Null)) Or (((MSysObjects.Database) Is Null) An"
    "d ((tsys_Link_Dbs.FilePath) Is Not Null)) Or (((MSysObjects.Database) Is Not Nul"
    "l) And ((tsys_Link_Dbs.FilePath) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x29d09db68b924849984af5d6ae697ef6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tsys_Link_Dbs.IsODBC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrDb"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4115b61ac006894b885554ba1f634f8f
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Tables.LinkDb"
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
            0xdd6684d7f21c7a4f9e410182682461ba
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.FilePath"
        dbLong "AggregateType" ="-1"
    End
End
