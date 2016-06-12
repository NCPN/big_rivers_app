dbMemo "SQL" ="SELECT DISTINCT TOP 15 ParkCode, Segment, Master_Species, mp.LU_Code\015\012FROM"
    " (SELECT DISTINCT ParkCode, Segment, Year(StartDate) AS LastYr,\015\012Master_Sp"
    "ecies, mp.LU_Code, sp.PercentCover \015\012FROM (((((Park p\015\012LEFT OUTER JO"
    "IN River r ON r.Park_ID = p.ID)\015\012LEFT OUTER JOIN Site s ON s.Park_ID = p.I"
    "D)\015\012LEFT OUTER JOIN Event e ON e.Site_ID = s.ID)\015\012LEFT OUTER JOIN Ve"
    "gPlot v ON v.Event_ID = e.ID)\015\012LEFT OUTER JOIN RootedSpecies sp ON sp.VegP"
    "lot_ID = v.ID)\015\012LEFT OUTER JOIN tlu_NCPN_Plants mp ON mp.Master_PLANT_Code"
    " = sp.Master_PLANT_Code\015\012WHERE\015\012Year(StartDate) = 2015\015\012AND\015"
    "\012ParkCode = 'BLCA'\015\012AND\015\012Segment = 'Gunnison'\015\012AND s.IsActi"
    "veForProtocol = 1\015\012AND p.IsActiveForProtocol = 1\015\012ORDER BY PercentCo"
    "ver DESC,Master_Species ASC)  AS [%$##@_Alias]\015\012WHERE Master_Species <> ''"
    " AND mp.LU_Code <> '';\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4647687db86e30449b162468442b2f1d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ParkCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Segment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="mp.LU_Code"
        dbLong "AggregateType" ="-1"
    End
End
