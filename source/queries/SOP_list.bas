dbMemo "SQL" ="SELECT SOPNumber, FullName, Code, Version, SOPNumber&'-'&FullName AS NumName, Ef"
    "fectiveDate, RetireDate, Year(EffectiveDate) AS StartYear, Year(RetireDate) AS E"
    "ndYear\015\012FROM SOP\015\012GROUP BY SOPNumber, FullName, Code, Version, Effec"
    "tiveDate, RetireDate\015\012ORDER BY SOPNumber, EffectiveDate;\015\012"
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
        dbText "Name" ="NumName"
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
End
