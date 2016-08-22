dbMemo "SQL" ="PARAMETERS park Text ( 4 );\015\012SELECT p.Master_Family, p.Master_Species, p.L"
    "U_Code, vw.Master_PLANT_Code, vw.IsSeedling, 0 AS SEQ  \015\012FROM (((((VegWalk"
    "Species vw\015\012INNER JOIN tlu_NCPN_Plants p ON p.Master_PLANT_Code = vw.Maste"
    "r_PLANT_Code)\015\012INNER JOIN VegWalk v ON v.ID = vw.VegWalk_ID)\015\012INNER "
    "JOIN Event e ON e.ID = v.Event_ID)\015\012INNER JOIN Site s ON s.ID = e.Site_ID)"
    "\015\012INNER JOIN Park pk ON pk.ID = s.Park_ID)\015\012WHERE \015\012p.LU_Code "
    "IS NOT NULL \015\012AND YEAR(e.StartDate) = YEAR(Date())-1\015\012AND pk.ParkCod"
    "e = [park]\015\012UNION ALL SELECT TOP 8 NULL AS Master_Family, NULL AS Master_S"
    "pecies, NULL AS LU_Code, NULL AS Master_PLANT_Code, NULL as IsSeedling, 1 AS SEQ"
    " \015\012FROM usys_temp_table\015\012ORDER BY SEQ, LU_Code;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe9b773b0c737d2488e14c8a49de58510
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="p.Master_Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Master_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.IsSeedling"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SEQ"
        dbLong "AggregateType" ="-1"
    End
End
