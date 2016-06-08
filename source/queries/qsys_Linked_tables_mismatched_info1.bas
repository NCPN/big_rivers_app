dbMemo "SQL" ="SELECT MSysObjects.Name AS CurrTable, tsys_Link_Tables.Link_table, MSysObjects.T"
    "ype, tsys_Link_Dbs.Is_ODBC, IIf([Type]=4,ParseConnectionStr([Connect]),ParseFile"
    "Name([Database])) AS CurrDb, tsys_Link_Tables.Link_db, IIf([Type]=4,ParseConnect"
    "ionStr([Connect],'SERVER=')) AS CurrServer, tsys_Link_Dbs.Server, MSysObjects.Da"
    "tabase AS CurrPath, tsys_Link_Dbs.File_path\015\012FROM tsys_Link_Dbs INNER JOIN"
    " (MSysObjects INNER JOIN tsys_Link_Tables ON MSysObjects.Name = tsys_Link_Tables"
    ".Link_table) ON tsys_Link_Dbs.Link_db = tsys_Link_Tables.Link_db\015\012WHERE (("
    "(MSysObjects.Type) In (4,6)) And ((IIf([Type]=4,ParseConnectionStr([Connect]),Pa"
    "rseFileName([Database])))<>tsys_Link_Tables.Link_db)) Or (((MSysObjects.Type) In"
    " (4,6)) And ((IIf([Type]=4,ParseConnectionStr([Connect],'SERVER=')))<>[Server]))"
    " Or (((MSysObjects.Type) In (4,6)) And ((MSysObjects.Database)<>[File_path])) Or"
    " (((MSysObjects.Type)=4) And ((tsys_Link_Dbs.Is_ODBC)=False)) Or (((MSysObjects."
    "Type)=6) And ((tsys_Link_Dbs.Is_ODBC)=True)) Or (((IIf([Type]=4,ParseConnectionS"
    "tr([Connect],'SERVER='))) Is Null) And ((tsys_Link_Dbs.Server) Is Not Null)) Or "
    "(((IIf([Type]=4,ParseConnectionStr([Connect],'SERVER='))) Is Not Null) And ((tsy"
    "s_Link_Dbs.Server) Is Null)) Or (((MSysObjects.Database) Is Null) And ((tsys_Lin"
    "k_Dbs.File_path) Is Not Null)) Or (((MSysObjects.Database) Is Not Null) And ((ts"
    "ys_Link_Dbs.File_path) Is Null));\015\012"
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
    0xba8d34b8f50d8a4bb9e7ccacd812d8d6
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
        dbText "Name" ="tsys_Link_Tables.Link_db"
        dbInteger "ColumnWidth" ="3450"
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
        dbText "Name" ="tsys_Link_Dbs.File_path"
        dbInteger "ColumnWidth" ="8145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Tables.Link_table"
        dbInteger "ColumnWidth" ="3255"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="tsys_Link_Dbs.Is_ODBC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.Server"
        dbLong "AggregateType" ="-1"
    End
End
