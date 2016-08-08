dbMemo "SQL" ="PARAMETERS ParkCode Text ( 4 );\015\012SELECT f.ID, f.Feature, loc.LocationName,"
    " f.Location_ID\015\012FROM (((Feature AS f LEFT JOIN Location AS loc ON loc.ID ="
    " f.Location_ID) LEFT JOIN Site_Feature AS sf ON sf.Feature_ID = f.ID) LEFT JOIN "
    "Site AS s ON s.ID = sf.Site_ID) LEFT JOIN Park AS p ON p.ID = s.Park_ID\015\012W"
    "HERE p.ParkCode = [ParkCode]\015\012AND s.IsActiveForProtocol = 1\015\012ORDER B"
    "Y f.Feature;\015\012"
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
        dbText "Name" ="p.Utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Master_Plant_Code"
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
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
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
    Begin
        dbText "Name" ="Master_Plant_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Utah_species"
        dbInteger "ColumnWidth" ="3705"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Presence"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Presence"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StateAbbr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Master_Plant_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.ParkCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.StateAbbr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="r.Segment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.SiteName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ddEvent"
        dbInteger "ColumnWidth" ="2295"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2f6145602a3f52458129ce5c7b0bc00a
        End
    End
    Begin
        dbText "Name" ="t.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Presence"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ddSpecies"
        dbInteger "ColumnWidth" ="3120"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2089d88e515449479a0719ad69352730
        End
    End
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
End
