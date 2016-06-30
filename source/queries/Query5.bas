dbMemo "SQL" ="PARAMETERS ParkCode Text ( 4 ), waterway Text ( 25 );\015\012SELECT l.ID, l.Loca"
    "tionName, l.CollectionSourceName, l.LocationType\015\012FROM (((Location AS l IN"
    "NER JOIN Event AS e ON e.Location_ID = l.ID) INNER JOIN Site AS s ON s.ID = e.Si"
    "te_ID) INNER JOIN Park AS p ON s.Park_ID = p.ID) INNER JOIN River AS r ON r.ID ="
    " s.River_ID\015\012WHERE p.ParkCode = [ParkCode]\015\012AND r.River = [waterway]"
    ";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x860fbb421b73d1439e28763fcfb4076d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="l.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.CollectionSourceName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.LocationType"
        dbLong "AggregateType" ="-1"
    End
End
