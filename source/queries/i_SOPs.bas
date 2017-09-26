dbMemo "SQL" ="SELECT SOP, SOPNum, Code, Version, EffectiveDate FROM SOP2011\015\012UNION\015\012"
    "SELECT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2012\015\012UNION\015\012"
    "SELECT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2013\015\012UNION\015\012"
    "SELECT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2014\015\012UNION\015\012"
    "SELECT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2015\015\012UNION\015\012"
    "SELECT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2016\015\012UNION SELE"
    "CT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2017a\015\012ORDER BY Effe"
    "ctiveDate, SOPNum;\015\012"
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
        dbText "Name" ="SOP"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SOPNum"
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
End
