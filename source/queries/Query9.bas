dbMemo "SQL" ="PARAMETERS pk Text ( 4 );\015\012SELECT p.ID AS PhotoID, p.PhotoType, p.PhotoDat"
    "e, p.Photographer_ID, p.PhotogFacing, p.PhotogLocation, p.PhotogOrientation, p.S"
    "ubjectLocation, p.IsCloseup, p.IsReplacement, p.IsSkipped, e.StartDate, e.ID AS "
    "EventID, c.FirstName, c.LastName, c.FirstName & ' ' & c.LastName AS PhotogName, "
    "c.Email, s.SiteCode, s.ID AS SiteID, s.Park_ID, s.River_ID, pk.ParkCode, r.River"
    ", r.Segment\015\012FROM (((((Photo AS p INNER JOIN Event_Photo AS ep ON ep.Photo"
    "_ID = p.ID) INNER JOIN Event AS e ON e.ID = ep.Event_ID) INNER JOIN Contact AS c"
    " ON c.ID = p.Photographer_ID) INNER JOIN Site AS s ON s.ID = e.Site_ID) INNER JO"
    "IN River AS r ON r.ID = s.River_ID) INNER JOIN Park AS pk ON pk.ID = s.Park_ID\015"
    "\012WHERE pk.ParkCode = [pk] Or [pk] IS NULL;\015\012"
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
        dbText "Name" ="p.PhotogLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.SubjectLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="c.Email"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.PhotogFacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SiteID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1de503425a769f45907d1cd82e31e461
        End
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
        dbText "Name" ="p.IsSkipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.StartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbb20bf1dbfc02a48bed2326b772b32c9
        End
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
        dbBinary "GUID" = Begin
            0xf741c30a65c5ec469d1598fb70fa921d
        End
    End
    Begin
        dbText "Name" ="p.PhotogOrientation"
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
        dbText "Name" ="r.Segment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotoID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcdab31023e6a87419f190fd7dc7df3f7
        End
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
End
