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
    0xd0adfdbe4d7a184da5399778863d2cc8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
