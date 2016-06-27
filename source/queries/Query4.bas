dbMemo "SQL" ="SELECT *\015\012FROM (((Feature AS f INNER JOIN Location AS l ON l.ID = f.Locati"
    "on_ID) INNER JOIN Site_Feature AS sf ON sf.Feature_ID = f.ID) INNER JOIN Site AS"
    " s ON s.ID = sf.Site_ID) INNER JOIN Park AS p ON p.ID = s.Park_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd519861fa67ed34c8f7899022ed55dae
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="f.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="f.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="f.Feature"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="f.FeatureDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="f.FeatureDirections"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.CollectionSourceName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sf.Feature_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.ID"
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
        dbText "Name" ="s.SiteDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.IsActiveForProtocol"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.ID"
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
        dbText "Name" ="l.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.HeadtoOrientDistance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.HeadtoOrientBearing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.LocationNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.CreateDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.CreatedBy_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.LastModified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.LastModifiedBy_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sf.Site_ID"
        dbLong "AggregateType" ="-1"
    End
End
