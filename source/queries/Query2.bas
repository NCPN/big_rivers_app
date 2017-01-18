dbMemo "SQL" ="SELECT SOPNumber, FullName, Code, Version, SOPNumber&'-'&FullName AS NumName, Ef"
    "fectiveDate, RetireDate, Year(EffectiveDate) AS StartYear, Year(RetireDate) AS E"
    "ndYear\015\012FROM SOP\015\012GROUP BY SOPNumber, FullName, Code, Version, Effec"
    "tiveDate, RetireDate\015\012ORDER BY SOPNumber, EffectiveDate;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa56e814763ff6b4b8882d15bc19beba7
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "OrderBy" ="[Query2].[StartYear]"
Begin
    Begin
        dbText "Name" ="SOPNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Version"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EffectiveDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RetireDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StartYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EndYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumName"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
