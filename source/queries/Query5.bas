dbMemo "SQL" ="PARAMETERS pkcode Text ( 4 ), scode Text ( 2 );\015\012SELECT l.ID, CollectionSo"
    "urceName, LocationType, LocationName, HeadToOrientDistance_m, HeadToOrientBearin"
    "g, LocationNotes, s.SiteCode, p.ParkCode, p.ID AS ParkID, Site_ID, SWITCH (\015\012"
    "LocationType ='P', \015\012(\015\012SELECT DISTINCT vp.ID \015\012FROM ((VegPlot"
    " vp\015\012INNER JOIN Site s ON s.ID = vp.Site_ID)\015\012INNER JOIN Park p ON p"
    ".ID = s.Park_ID)\015\012WHERE \015\012CStr(PlotNumber) = CStr(CollectionSourceNa"
    "me)\015\012AND p.ParkCode = [pkcode]\015\012AND s.SiteCode = [scode]\015\012),\015"
    "\012LocationType ='T',  \015\012(\015\012SELECT DISTINCT vt.ID \015\012FROM (((V"
    "egTransect vt\015\012INNER JOIN Site_VegTransect svt ON svt.VegTransect_ID = vt."
    "ID)\015\012INNER JOIN Site s ON s.ID = svt.Site_ID)\015\012INNER JOIN Park p ON "
    "p.ID = s.Park_ID)\015\012WHERE \015\012CStr(TransectNumber) = CStr(CollectionSou"
    "rceName)\015\012AND p.ParkCode = [pkcode]\015\012AND s.SiteCode = [scode]\015\012"
    "),\015\012LocationType ='F',  \015\012(\015\012SELECT DISTINCT f.ID \015\012FROM"
    " (((Feature f\015\012INNER JOIN Site_Feature sf ON sf.Feature_ID = f.ID)\015\012"
    "INNER JOIN Site s ON s.ID = sf.Site_ID)\015\012INNER JOIN Park p ON p.ID = s.Par"
    "k_ID)\015\012WHERE \015\012CStr(Feature) = CStr(CollectionSourceName)\015\012AND"
    " p.ParkCode = [pkcode]\015\012AND s.SiteCode = [scode];\015\012)\015\012) AS Loc"
    "TypeID, (SELECT COUNT(sl.Location_ID) FROM SensitiveLocations sl WHERE sl.Locati"
    "on_ID = l.ID\015\012) AS IsSensitive\015\012FROM (Location AS l INNER JOIN Site "
    "AS s ON s.ID = l.Site_ID) INNER JOIN Park AS p ON p.ID = s.Park_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd3e36b46ca5173488cf487592f7b1c5c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="vp.PlotNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CollectionSourceName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationName"
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
        dbText "Name" ="LocationNotes"
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
    Begin
        dbText "Name" ="s.SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.ParkCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Site_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Park_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkID"
        dbLong "AggregateType" ="-1"
    End
End
