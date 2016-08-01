dbMemo "SQL" ="PARAMETERS ParkCode Text ( 4 ), waterway Text ( 25 );\015\012SELECT l.ID, l.Loca"
    "tionName, l.CollectionSourceName, l.LocationType\015\012FROM (((Location AS l IN"
    "NER JOIN Event AS e ON e.Location_ID = l.ID) INNER JOIN Site AS s ON s.ID = e.Si"
    "te_ID) INNER JOIN Park AS p ON s.Park_ID = p.ID) INNER JOIN River AS r ON r.ID ="
    " s.River_ID\015\012WHERE p.ParkCode = [ParkCode]\015\012AND r.River = [waterway]"
    ";\015\012"
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
