dbMemo "SQL" ="PARAMETERS pkcode Text ( 4 ), scode Text ( 2 );\015\012SELECT l.ID, CollectionSo"
    "urceName, LocationType, LocationName, HeadToOrientDistance_m, HeadToOrientBearin"
    "g, LocationNotes, s.SiteCode, p.ParkCode, p.ID AS ParkID, Site_ID, SWITCH ( \015"
    "\012LocationType ='P',  CollectionSourceName, \015\012LocationType ='T',  Collec"
    "tionSourceName, \015\012LocationType ='F',   ( \015\012\011SELECT DISTINCT f.ID "
    " \015\012\011FROM (((Feature f \015\012\011INNER JOIN Site_Feature sf ON sf.Feat"
    "ure_ID = f.ID) \015\012\011INNER JOIN Site s ON s.ID = sf.Site_ID) \015\012\011I"
    "NNER JOIN Park p ON p.ID = s.Park_ID) \015\012\011WHERE  CStr(Feature) = CStr(Co"
    "llectionSourceName) \015\012\011AND p.ParkCode = [pkcode] \015\012\011AND s.Site"
    "Code = [scode] ),\015\012LocationType = '1', 99\015\012) AS LocTypeID, (SELECT C"
    "OUNT(sl.Location_ID) FROM SensitiveLocations sl WHERE sl.Location_ID = l.ID) AS "
    "IsSensitive\015\012FROM (Location AS l INNER JOIN Site AS s ON s.ID = l.Site_ID)"
    " INNER JOIN Park AS p ON p.ID = s.Park_ID\015\012WHERE p.ParkCode = [pkcode] \015"
    "\012AND s.SiteCode = [scode];\015\012"
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
        dbText "Name" ="l.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HeadToOrientDistance_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HeadToOrientBearing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CollectionSourceName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.ParkCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Site_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IsSensitive"
        dbLong "AggregateType" ="-1"
    End
End
