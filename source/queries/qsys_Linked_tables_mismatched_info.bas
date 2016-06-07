dbMemo "SQL" ="SELECT MSysObjects.Name AS CurrTable, tsys_Link_Tables.LinkTable, MSysObjects.Ty"
    "pe, tsys_Link_Dbs.[IsODBC], IIf([Type]=4,ParseConnectionStr([Connect]),ParseFile"
    "Name([Database])) AS CurrDb, tsys_Link_Tables.[LinkDb], IIf([Type]=4,ParseConnec"
    "tionStr([Connect],'SERVER=')) AS CurrServer, tsys_Link_Dbs.Server, MSysObjects.D"
    "atabase AS CurrPath, tsys_Link_Dbs.[FilePath]\015\012FROM tsys_Link_Dbs INNER JO"
    "IN (MSysObjects INNER JOIN tsys_Link_Tables ON MSysObjects.Name = tsys_Link_Tabl"
    "es.LinkTable) ON tsys_Link_Dbs.[LinkDb]=tsys_Link_Tables.[LinkDb]\015\012WHERE ("
    "((MSysObjects.Type) In (4,6)) And ((IIf([Type]=4,ParseConnectionStr([Connect]),P"
    "arseFileName([Database])))<>tsys_Link_Tables.[LinkDb])) Or (((MSysObjects.Type) "
    "In (4,6)) And ((IIf([Type]=4,ParseConnectionStr([Connect],'SERVER=')))<>[Server]"
    ")) Or (((MSysObjects.Type) In (4,6)) And ((MSysObjects.Database)<>[FilePath])) O"
    "r (((MSysObjects.Type)=4) And ((tsys_Link_Dbs.[IsODBC])=False)) Or (((MSysObject"
    "s.Type)=6) And ((tsys_Link_Dbs.[IsODBC])=True)) Or (((IIf([Type]=4,ParseConnecti"
    "onStr([Connect],'SERVER='))) Is Null) And ((tsys_Link_Dbs.Server) Is Not Null)) "
    "Or (((IIf([Type]=4,ParseConnectionStr([Connect],'SERVER='))) Is Not Null) And (("
    "tsys_Link_Dbs.Server) Is Null)) Or (((MSysObjects.Database) Is Null) And ((tsys_"
    "Link_Dbs.[FilePath]) Is Not Null)) Or (((MSysObjects.Database) Is Not Null) And "
    "((tsys_Link_Dbs.[FilePath]) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Matches MSysObjects.Name with tsys_Link_Tables.Link_table, finds mismatches on d"
    "b name, server, file path, or where ODBC doesn't match the actual table link typ"
    "e"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x671dd0c2cdb5ac48b76fd089ad369a02
End
Begin
    Begin
        dbText "Name" ="CurrDb"
        dbInteger "ColumnWidth" ="3210"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x027c55cddfea224bb621541960f3718a
        End
    End
    Begin
        dbText "Name" ="MSysObjects.Type"
        dbInteger "ColumnWidth" ="795"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrServer"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd2ffa600084b9948abbb3fefcfa53940
        End
    End
    Begin
        dbText "Name" ="CurrTable"
        dbInteger "ColumnWidth" ="3255"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe829b56b7af0dc4bacdca3fb337e7e84
        End
    End
    Begin
        dbText "Name" ="CurrPath"
        dbInteger "ColumnWidth" ="8805"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1632622226673b4686e3b38aa8033cd6
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.Server"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Tables.LinkTable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.[IsODBC]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Tables.[LinkDb]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.[FilePath]"
        dbLong "AggregateType" ="-1"
    End
End
