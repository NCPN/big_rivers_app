dbMemo "SQL" ="PARAMETERS pkcode Text ( 4 ), waterway Text ( 25 );\015\012SELECT e.ID, s.SiteCo"
    "de, s.SiteName, r.Segment, CStr(e.StartDate) + \" - \" + s.SiteName AS ddEvent\015"
    "\012FROM ((Event AS e INNER JOIN Site AS s ON s.ID = e.Site_ID) INNER JOIN Park "
    "AS p ON s.Park_ID = p.ID) INNER JOIN River AS r ON r.ID = s.River_ID\015\012WHER"
    "E p.ParkCode = [pkcode]\015\012AND r.River = [waterway]\015\012ORDER BY s.SiteNa"
    "me, e.StartDate DESC;\015\012"
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
