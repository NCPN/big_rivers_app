dbMemo "SQL" ="SELECT MSysObjects.Name AS CurrTable, tsys_Link_Tables.LinkTable, MSysObjects.Ty"
    "pe, tsys_Link_Dbs.[IsODBC], IIf([Type]=4,ParseConnectionStr([Connect]),ParseFile"
    "Name([Database])) AS CurrDb, tsys_Link_Tables.[LinkDb], IIf([Type]=4,ParseConnec"
    "tionStr([Connect],'SERVER=')) AS CurrServer, tsys_Link_Dbs.Server, MSysObjects.D"
    "atabase AS CurrPath, tsys_Link_Dbs.[FilePath]\015\012FROM (tsys_Link_Dbs INNER J"
    "OIN tsys_Link_Tables ON tsys_Link_Dbs.[LinkDb]=tsys_Link_Tables.[LinkDb]) INNER "
    "JOIN MSysObjects ON MSysObjects.Name = tsys_Link_Tables.LinkTable\015\012WHERE M"
    "SysObjects.Type In (4,6)\015\012AND\015\012(\015\012(\015\012IIf([Type]=4,ParseC"
    "onnectionStr([Connect]),ParseFileName([Database]))<>tsys_Link_Tables.[LinkDb]\015"
    "\012OR\015\012IIf([Type]=4,ParseConnectionStr([Connect],'SERVER='))<>[Server] \015"
    "\012OR\015\012MSysObjects.Database<>[FilePath]\015\012)\015\012OR\015\012(\015\012"
    "MSysObjects.Type=4\015\012AND\015\012tsys_Link_Dbs.[IsODBC]=False\015\012)\015\012"
    "OR\015\012(\015\012MSysObjects.Type=6 \015\012AND\015\012tsys_Link_Dbs.[IsODBC]="
    "True\015\012)\015\012OR\015\012(\015\012IIf([Type]=4,ParseConnectionStr([Connect"
    "],'SERVER=')) IS NULL\015\012AND \015\012tsys_Link_Dbs.Server IS NOT NULL\015\012"
    ")\015\012OR\015\012( \015\012IIf([Type]=4,ParseConnectionStr([Connect],'SERVER='"
    ")) IS NOT NULL\015\012AND\015\012tsys_Link_Dbs.Server IS NULL\015\012) \015\012O"
    "R\015\012(\015\012MSysObjects.Database IS NULL\015\012AND \015\012tsys_Link_Dbs."
    "[FilePath] IS NOT NULL\015\012)\015\012\015\012OR \015\012(\015\012MSysObjects.D"
    "atabase IS NOT NULL\015\012AND\015\012tsys_Link_Dbs.[FilePath] IS NULL\015\012)\015"
    "\012\015\012);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="CurrTable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Tables.LinkTable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MSysObjects.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.[IsODBC]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrDb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Tables.[LinkDb]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrServer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.Server"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrPath"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.[FilePath]"
        dbLong "AggregateType" ="-1"
    End
End
