dbMemo "SQL" ="INSERT INTO SOP ( FullName, SOPNumber, Code, Version, EffectiveDate, CreatedBy_I"
    "D, LastModifiedBy_ID )\015\012SELECT SOP, SOPNum, Code, Version, EffectiveDate, "
    "1, 1\015\012FROM i_SOPs;\015\012"
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
