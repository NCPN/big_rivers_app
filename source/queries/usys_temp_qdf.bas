dbMemo "SQL" ="PARAMETERS pkcode Text ( 4 ), seg Text ( 25 );\015\012SELECT DISTINCT l.ID, l.Lo"
    "cationName, l.CollectionSourceName, l.LocationType\015\012FROM (((Location AS l "
    "INNER JOIN Event AS e ON e.Location_ID = l.ID) INNER JOIN Site AS s ON s.ID = e."
    "Site_ID) INNER JOIN Park AS p ON s.Park_ID = p.ID) INNER JOIN River AS r ON r.ID"
    " = s.River_ID\015\012WHERE p.ParkCode = [pkcode]\015\012AND r.Segment = [seg];\015"
    "\012"
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
