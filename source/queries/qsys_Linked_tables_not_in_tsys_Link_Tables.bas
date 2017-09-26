dbMemo "SQL" ="SELECT MSysObjects.Name AS CurrTable, IIf([Type]=4,ParseConnectionStr([Connect])"
    ",ParseFileName([Database])) AS CurrDb, IIf([Type]=4,ParseConnectionStr([Connect]"
    ",'SERVER=')) AS CurrServer, IIf([Type]=6,[Database]) AS CurrPath, IIf([Type]=4,T"
    "rue,False) AS ODBC\015\012FROM MSysObjects LEFT JOIN tsys_Link_Tables ON MSysObj"
    "ects.Name = tsys_Link_Tables.LinkTable\015\012WHERE (((MSysObjects.Name) Not Lik"
    "e \"~*\") AND ((tsys_Link_Tables.LinkTable) Is Null) AND ((MSysObjects.Type) In "
    "(4,6)));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Linked tables in MSysObjects that do not have records in tsys_Link_Tables (other"
    " than recently deleted objects that start with '~')"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CurrDb"
        dbInteger "ColumnWidth" ="2430"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrServer"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrPath"
        dbInteger "ColumnWidth" ="9285"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrTable"
        dbInteger "ColumnWidth" ="2760"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ODBC"
        dbLong "AggregateType" ="-1"
    End
End
