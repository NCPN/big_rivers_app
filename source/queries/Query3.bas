dbMemo "SQL" ="PARAMETERS mi Text ( 255 );\015\012UPDATE Contact SET middleinitial = mi\015\012"
    "WHERE id = 1;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa5ea926e0cec524fb655814d93802cb5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Contact.middleinitial"
        dbLong "AggregateType" ="-1"
    End
End
