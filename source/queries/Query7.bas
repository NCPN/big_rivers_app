dbMemo "SQL" ="SELECT TOP 10 UnderstoryRootedPctCover, *\015\012FROM ((((RootedSpecies AS rs LE"
    "FT JOIN VegPlot AS v ON v.ID = rs.VegPlot_ID) LEFT JOIN Event AS e ON e.ID = v.E"
    "vent_ID) LEFT JOIN Site AS s ON s.ID = e.Site_ID) LEFT JOIN River AS r ON r.ID ="
    " s.River_ID) LEFT JOIN Park AS p ON p.ID = s.Park_ID\015\012WHERE Year(StartDate"
    ") = Year(Now())-1\015\012AND\015\012ParkCode = 'BLCA'\015\012AND\015\012Segment "
    "= 'Gunnison'\015\012AND s.IsActiveForProtocol = 1\015\012AND p.IsActiveForProtoc"
    "ol = 1\015\012ORDER BY UnderstoryRootedPctCover, Master_PLANT_Code;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x204df94ccf70ab45b50e0d98255e6510
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="s.SiteDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="rs.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="rs.VegPlot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="rs.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.UnderstoryRootedPctCover"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="v.PlotDensity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.StartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.ModalSedimentSize"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.ParkCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.ParkName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.ParkState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.IsActiveForProtocol"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.HasSocialTrail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.NoIndicatorSpecies"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Protocol_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Park_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.River_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteDirections"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.VegTransect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.PercentWater"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.IsActiveForProtocol"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.NoRootedVeg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Site_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="rs.PercentCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.Site_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.Feature_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.PlotDistance_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.PercentFine"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.NoCanopyVeg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.FilamentousAlgae"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.PlotNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="r.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="r.Park_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="r.River"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="r.Segment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PercentCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnderstoryRootedPctCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RootedSpecies.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RootedSpecies.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RootedSpecies.VegPlot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RootedSpecies.PercentCover"
        dbLong "AggregateType" ="-1"
    End
End
