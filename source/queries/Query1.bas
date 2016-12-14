dbMemo "SQL" ="SELECT t.ID, t.TaskType, t.Task, t.Status_ID, t.Priority_ID, t.RequestedBy_ID, t"
    ".RequestDate, t.CompletedBy_ID, t.CompleteDate, p.Priority, p.Icon, s.Status, s."
    "Icon, c1.FirstName + ' ' + c1.LastName AS Requestor, c1.ID, c2.FirstName + ' ' +"
    " c2.LastName AS Completor, c2.ID\015\012FROM (((Task AS t LEFT JOIN Priority AS "
    "p ON p.ID = t.Priority_ID) LEFT JOIN Status AS s ON s.ID = t.Status_ID) LEFT JOI"
    "N Contact AS c1 ON c1.ID = t.RequestedBy_ID) LEFT JOIN Contact AS c2 ON c2.ID = "
    "t.CompletedBy_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xce35d907c67e3e49b0b0abca28cdc6d3
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="p.ParkCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.SensorType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.SensorNumber"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="r.Segment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PtName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ptID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Easting_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Northing_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Elevation_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Status_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Priority_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.RequestedBy_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.RequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.CompletedBy_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.CompleteDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Priority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Icon"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Icon"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Requestor"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe92aa8c03bc0a440b9abb6410e6243a3
        End
    End
    Begin
        dbText "Name" ="Completor"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0c7bb549e85b6543a1b8319e956f3285
        End
    End
    Begin
        dbText "Name" ="c1.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="c2.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TaskType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Task"
        dbLong "AggregateType" ="-1"
    End
End
