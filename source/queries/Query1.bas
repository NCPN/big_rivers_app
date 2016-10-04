dbMemo "SQL" ="SELECT p.ParkCode, r.Segment, s.SiteName, s.SiteCode, l.SensorType, l.SensorNumb"
    "er\015\012FROM ((Logger AS l LEFT JOIN Site AS s ON s.ID = l.Site_ID) LEFT JOIN "
    "River AS r ON r.ID = s.River_ID) LEFT JOIN Park AS p ON p.ID = s.Park_ID\015\012"
    "ORDER BY p.ParkCode, r.Segment, s.SiteName;\015\012"
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
End
