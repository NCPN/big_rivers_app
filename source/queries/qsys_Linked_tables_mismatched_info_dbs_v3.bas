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
dbBinary "GUID" = Begin
    0x847662cf74d5b440a20ee268493bf3f1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="CurrTable"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x42d79bf5d8653c4899acf87c052a7a2c
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Tables.LinkTable"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe2f22821daba9c4b81915ec67c5a8a33
        End
    End
    Begin
        dbText "Name" ="MSysObjects.Type"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x59d1d358657f7c4b88262b3a6d44d60d
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.[IsODBC]"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x90d4307e0f63934da336f640c492849c
        End
    End
    Begin
        dbText "Name" ="CurrDb"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfae7428c26074347b73fd40103193279
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Tables.[LinkDb]"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x592cd6b58f4ac349a4ee12a44cc89213
        End
    End
    Begin
        dbText "Name" ="CurrServer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x984810f2090cb8438b70134419a37b00
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.Server"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdeadb4c9288f694d8b4841df6d7f0f6f
        End
    End
    Begin
        dbText "Name" ="CurrPath"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x84824942b4bf944fbc0680297e93135e
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.[FilePath]"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x60e3607f08994748859bb4d13faa8101
        End
    End
End
