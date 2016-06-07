dbMemo "SQL" ="SELECT DISTINCT tsys_Link_Dbs.[IsODBC], IIf([Type]=4,ParseConnectionStr([Connect"
    "]),ParseFileName([Database])) AS CurrDb, tsys_Link_Tables.[LinkDb], tsys_Link_Db"
    "s.Server, MSysObjects.Database AS CurrPath, tsys_Link_Dbs.[FilePath]\015\012FROM"
    " tsys_Link_Dbs INNER JOIN (MSysObjects INNER JOIN tsys_Link_Tables ON MSysObject"
    "s.Name = tsys_Link_Tables.LinkTable) ON tsys_Link_Dbs.[LinkDb]=tsys_Link_Tables."
    "[LinkDb]\015\012WHERE (((MSysObjects.Type) In (4,6)) And ((IIf([Type]=4,ParseCon"
    "nectionStr([Connect]),ParseFileName([Database])))<>tsys_Link_Tables.[LinkDb])) O"
    "r (((MSysObjects.Type) In (4,6)) And ((IIf([Type]=4,ParseConnectionStr([Connect]"
    ",'SERVER=')))<>[Server])) Or (((MSysObjects.Type) In (4,6)) And ((MSysObjects.Da"
    "tabase)<>[FilePath])) Or (((MSysObjects.Type)=4) And ((tsys_Link_Dbs.[IsODBC])=F"
    "alse)) Or (((MSysObjects.Type)=6) And ((tsys_Link_Dbs.[IsODBC])=True)) Or (((IIf"
    "([Type]=4,ParseConnectionStr([Connect],'SERVER='))) Is Null) And ((tsys_Link_Dbs"
    ".Server) Is Not Null)) Or (((IIf([Type]=4,ParseConnectionStr([Connect],'SERVER='"
    "))) Is Not Null) And ((tsys_Link_Dbs.Server) Is Null)) Or (((MSysObjects.Databas"
    "e) Is Null) And ((tsys_Link_Dbs.[FilePath]) Is Not Null)) Or (((MSysObjects.Data"
    "base) Is Not Null) And ((tsys_Link_Dbs.[FilePath]) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBinary "GUID" = Begin
    0x51a39ad7aa1512418325da7f505f58f5
End
Begin
    Begin
        dbText "Name" ="CurrDb"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0a751d867b87964f80ff51de9783d6db
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.Server"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x093fdb5dc079cc49b5949efef5b73b13
        End
    End
    Begin
        dbText "Name" ="CurrPath"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3f33a8983034f346838ec9d9e2277031
        End
    End
End
