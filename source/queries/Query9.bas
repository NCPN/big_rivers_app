dbMemo "SQL" ="SELECT DISTINCT TOP 15 ParkCode, Segment, Master_Species, mp.LU_Code\015\012FROM"
    " (SELECT DISTINCT ParkCode, Segment, Year(StartDate) AS LastYr, Master_Species, "
    "mp.LU_Code, sp.PercentCover FROM (((((Park AS p LEFT JOIN River AS r ON r.Park_I"
    "D = p.ID) LEFT JOIN Site AS s ON s.Park_ID = p.ID) LEFT JOIN Event AS e ON e.Sit"
    "e_ID = s.ID) LEFT JOIN VegPlot AS v ON v.Event_ID = e.ID) LEFT JOIN RootedSpecie"
    "s AS sp ON sp.VegPlot_ID = v.ID) LEFT JOIN tlu_NCPN_Plants AS mp ON mp.Master_PL"
    "ANT_Code = sp.Master_PLANT_Code WHERE Year(StartDate) = 2015 AND ParkCode = 'BLC"
    "A' AND Segment = 'Gunnison' AND s.IsActiveForProtocol = 1 AND p.IsActiveForProto"
    "col = 1 ORDER BY PercentCover DESC , Master_Species)  AS [%$##@_Alias]\015\012WH"
    "ERE Master_Species <> '' AND mp.LU_Code <> '';\015\012"
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
