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
dbBinary "GUID" = Begin
    0xd0adfdbe4d7a184da5399778863d2cc8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
