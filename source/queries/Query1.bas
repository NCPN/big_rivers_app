dbMemo "SQL" ="INSERT INTO SurveyPoint ( PointName, PointType, XCoord, YCoord, ZCoord, PointDes"
    "cription, IsDeleted )\015\012SELECT PtName, ptID, Easting_m, Northing_m, Elevati"
    "on_m, Code, 0\015\012FROM usys_temp_csv;\015\012"
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
End
