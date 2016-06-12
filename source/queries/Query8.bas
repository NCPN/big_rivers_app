dbMemo "SQL" ="SELECT DISTINCT ParkCode, Segment, Year(StartDate) AS LastYr, Master_Species, sp"
    ".Master_PLANT_Code, sp.PercentCover\015\012FROM (((((Park AS p LEFT JOIN River A"
    "S r ON r.Park_ID = p.ID) LEFT JOIN SIte AS s ON s.Park_ID = p.ID) LEFT JOIN Even"
    "t AS e ON e.Site_ID = s.ID) LEFT JOIN VegPlot AS v ON v.Event_ID = e.ID) LEFT JO"
    "IN UnderstorySpecies AS sp ON sp.VegPlot_ID = v.ID) LEFT JOIN tlu_NCPN_Plants AS"
    " mp ON mp.Master_PLANT_Code = sp.Master_PLANT_Code\015\012WHERE Year(StartDate) "
    "= Year(Now())-1\015\012AND\015\012ParkCode = 'BLCA'\015\012AND\015\012Segment = "
    "'Gunnison'\015\012AND s.IsActiveForProtocol = 1\015\012AND p.IsActiveForProtocol"
    " = 1\015\012ORDER BY PercentCover DESC , sp.Master_PLANT_Code;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x33c38c4b452fdc4898fe8ae3d67a5b65
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
        dbText "Name" ="LastYr"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5929bde89251644187a1f283ad0f8695
        End
    End
    Begin
        dbText "Name" ="sp.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sp.PercentCover"
        dbLong "AggregateType" ="-1"
    End
End
