dbMemo "SQL" ="SELECT p.ID AS PhotoID, p.PhotoType, p.PhotoDate, p.Photographer_ID, p.PhotogFac"
    "ing, p.PhotogLocation, p.PhotogOrientation, p.SubjectLocation, p.IsCloseup, p.Is"
    "Replacement, p.IsSkipped, p.DigitalFilename AS PhotoFilename, NULL AS PhotoPath,"
    " e.StartDate, e.ID AS EventID, c.FirstName, c.LastName, c.FirstName & ' ' & c.La"
    "stName AS PhotogName, c.Email, s.SiteCode, s.ID AS SiteID, s.Park_ID, s.River_ID"
    ", pk.ParkCode, r.River, r.Segment\015\012FROM (((((Photo AS p INNER JOIN Event_P"
    "hoto AS ep ON ep.Photo_ID = p.ID) INNER JOIN Event AS e ON e.ID = ep.Event_ID) I"
    "NNER JOIN Contact AS c ON c.ID = p.Photographer_ID) INNER JOIN Site AS s ON s.ID"
    " = e.Site_ID) INNER JOIN River AS r ON r.ID = s.River_ID) INNER JOIN Park AS pk "
    "ON pk.ID = s.Park_ID\015\012ORDER BY p.PhotoType;\015\012"
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
        dbText "Name" ="PhotoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotoType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotoDate"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Photographer_ID"
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
        dbText "Name" ="PhotoFilename"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotoPath"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.StartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="c.FirstName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="c.LastName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotogName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="c.Email"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SiteID"
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
        dbText "Name" ="pk.ParkCode"
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
End
