dbMemo "SQL" ="PARAMETERS pkcode Text ( 4 ), waterway Text ( 25 );\015\012SELECT s.ID, SiteCode"
    ", SiteName, SiteDescription, SiteDirections, s.IsActiveForProtocol\015\012FROM ("
    "Site AS s INNER JOIN Park AS p ON p.ID = s.Park_ID) INNER JOIN River AS r ON r.I"
    "D = s.River_ID\015\012WHERE p.ParkCode = [pkcode]\015\012AND r.River = [waterway"
    "];\015\012"
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
        dbText "Name" ="s.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Site"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4c5e233e1de98045a4bfd0f54e84bf02
        End
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SiteName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SiteDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SiteDirections"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.IsActiveForProtocol"
        dbLong "AggregateType" ="-1"
    End
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
