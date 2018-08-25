dbMemo "SQL" ="SELECT p.ID, p.PhotoPath, p.PhotoFilename, p.PhotoDate, p.PhotoType, p.Photograp"
    "her_ID, p.DigitalFilename, p.NCPNImageID, p.PhotogFacing, p.PhotogLocation, p.Ph"
    "otogOrientation, p.SurveyPoint_ID, p.SubjectLocation, p.IsCloseup, p.IsReplaceme"
    "nt, p.IsSkipped, p.InActive, ep.Event_ID, e.StartDate AS EventDate, c.FirstName "
    "& ' ' & c.LastName AS Photographer, sp.PointName, sp.PointType, sp.XCoord, sp.YC"
    "oord, sp.ZCoord, sp.PointDescription, s.SiteCode, s.SiteName, s.SiteDirections, "
    "s.SiteDescription, s.IsActiveForProtocol\015\012FROM ((((Photo AS p LEFT JOIN Ev"
    "ent_Photo AS ep ON ep.Photo_ID = p.ID) LEFT JOIN Event AS e ON e.ID = ep.Event_I"
    "D) LEFT JOIN Contact AS c ON c.ID = p.Photographer_ID) LEFT JOIN SurveyPoint AS "
    "sp ON sp.ID = p.SurveyPoint_ID) LEFT JOIN Site AS s ON s.ID = e.Site_ID;\015\012"
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
        dbText "Name" ="p.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotoPath"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotoFilename"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotoDate"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotoType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Photographer_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.DigitalFilename"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.NCPNImageID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotogFacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotogLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotogOrientation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.SurveyPoint_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.SubjectLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.IsCloseup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.IsReplacement"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.IsSkipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.InActive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ep.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photographer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sp.PointName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sp.PointType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sp.XCoord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sp.YCoord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sp.ZCoord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sp.PointDescription"
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
End
