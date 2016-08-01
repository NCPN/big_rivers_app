dbMemo "SQL" ="SELECT DISTINCT id, label, summary, label & ' - ' & summary AS display, Sequence"
    "\015\012FROM Enum\015\012WHERE EnumType = '[etype]'\015\012ORDER BY Sequence;\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x3fdb1e0ba05ac64a9fa9228270a4b84b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="label"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="summary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="display"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1c6636cd801b0344acd278142f412e13
        End
    End
    Begin
        dbText "Name" ="Sequence"
        dbLong "AggregateType" ="-1"
    End
End
