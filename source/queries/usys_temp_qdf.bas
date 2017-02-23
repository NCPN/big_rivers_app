dbMemo "SQL" ="PARAMETERS pkcode Text ( 4 ), scode Text ( 2 );\015\012SELECT DISTINCT e.ID, e.S"
    "tartDate, e.StartDate  & ' - ' & s.SiteCode AS SiteEventDate, e.StartDate & ' - "
    "' & s.SiteName & ' (' & s.SiteCode  & ')' AS SiteNameEventDate, s.SiteCode, p.Pa"
    "rkCode\015\012FROM (Event AS e INNER JOIN Site AS s ON s.ID = e.Site_ID) INNER J"
    "OIN Park AS p ON p.ID = s.Park_ID\015\012WHERE s.SiteCode = [scode]\015\012AND\015"
    "\012p.ParkCode = [pkcode]\015\012ORDER BY e.StartDate DESC;\015\012"
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
