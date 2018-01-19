dbMemo "SQL" ="SELECT p.ID AS PhotoID, p.PhotoPath, p.PhotoFilename, p.PhotoType, p.PhotoDate, "
    "p.Photographer_ID, e.StartDate, p.Event_ID, c.FirstName, c.LastName, c.FirstName"
    " & ' ' & c.LastName AS PhotogName, c.Email, s.SiteCode, s.ID AS SiteID, s.Park_I"
    "D, s.River_ID, pk.ParkCode, r.River, r.Segment\015\012FROM ((((usys_temp_photo A"
    "S p LEFT JOIN Event AS e ON e.ID = p.Event_ID) LEFT JOIN Contact AS c ON c.ID = "
    "p.Photographer_ID) LEFT JOIN Site AS s ON s.ID = e.Site_ID) LEFT JOIN River AS r"
    " ON r.ID = s.River_ID) LEFT JOIN Park AS pk ON pk.ID = s.Park_ID\015\012ORDER BY"
    " p.PhotoType;\015\012"
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
        dbText "Name" ="p.PhotoPath"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotoFilename"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotoType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotoDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Photographer_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.StartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Event_ID"
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
