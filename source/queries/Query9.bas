dbMemo "SQL" ="PARAMETERS appuser Text ( 50 );\015\012SELECT c.ID, LastName +',  '+ FirstName +"
    " ' ('+ Username +')' AS AppUser, AccessLevel, IsActive\015\012FROM (Contact AS c"
    " INNER JOIN Contact_Access AS ca ON ca.Contact_ID = c.ID) INNER JOIN Access AS a"
    " ON a.ID = ca.Access_ID\015\012WHERE Username = [appuser];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe9b773b0c737d2488e14c8a49de58510
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="p.Master_Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Master_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.IsSeedling"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SEQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AppUser"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xad90ecfc3034f34dafbd0d3cc0ad184a
        End
    End
    Begin
        dbText "Name" ="c.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AccessLevel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IsActive"
        dbLong "AggregateType" ="-1"
    End
End
