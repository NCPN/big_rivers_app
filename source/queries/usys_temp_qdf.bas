dbMemo "SQL" ="PARAMETERS pcode Text ( 4 ), scode Text ( 2 );\015\012SELECT f.ID, f.Feature, lo"
    "c.LocationName, f.Location_ID\015\012FROM (((Feature AS f LEFT JOIN Location AS "
    "loc ON loc.ID = f.Location_ID) LEFT JOIN Site_Feature AS sf ON sf.Feature_ID = f"
    ".ID) LEFT JOIN Site AS s ON s.ID = sf.Site_ID) LEFT JOIN Park AS p ON p.ID = s.P"
    "ark_ID\015\012WHERE p.ParkCode = [pcode]\015\012AND s.SiteCode = [scode]\015\012"
    "AND s.IsActiveForProtocol = 1\015\012ORDER BY f.Feature;\015\012"
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
