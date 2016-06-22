dbMemo "SQL" ="SELECT Username, AccessLevel\015\012FROM (contact AS c INNER JOIN Contact_Access"
    " AS ca ON ca.Contact_ID = c.ID) INNER JOIN Access AS a ON a.ID = ca.Access_ID\015"
    "\012WHERE c.ID = 1;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x49d639c16c6b8f4b8cbe4487a421387d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Username"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AccessLevel"
        dbLong "AggregateType" ="-1"
    End
End
