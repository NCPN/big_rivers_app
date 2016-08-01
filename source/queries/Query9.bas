dbMemo "SQL" ="PARAMETERS ParkCode Text ( 4 ), waterway Text ( 25 );\015\012SELECT e.ID, s.Site"
    "Code, e.Site_ID, s.SiteName, loc.LocationName, e.Location_ID, e.StartDate\015\012"
    "FROM (((Event AS e INNER JOIN Site AS s ON s.ID = e.Site_ID) INNER JOIN Location"
    " AS loc ON loc.ID = e.Location_ID) INNER JOIN Park AS p ON p.ID = s.Park_ID) INN"
    "ER JOIN River AS r ON r.ID = s.River_ID\015\012WHERE p.ParkCode = [ParkCode]\015"
    "\012AND\015\012r.Segment = [waterway]\015\012ORDER BY e.StartDate DESC , s.SiteN"
    "ame, loc.LocationName;\015\012"
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
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="l.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_App_Releases.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Version"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xff04a13d1336de45908510e143ae9048
        End
    End
    Begin
        dbText "Name" ="e.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Site_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="loc.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.StartDate"
        dbLong "AggregateType" ="-1"
    End
End
