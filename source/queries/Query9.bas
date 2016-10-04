dbMemo "SQL" ="PARAMETERS yractive Long, activeonly Byte;\015\012SELECT w.ID, w.Code, w.Label, "
    "w.ActiveYear, w.RetireYear, w.DiameterRange_mm, w.Label + '('+ w.Code +')' AS ca"
    "tegory\015\012FROM ModWentworthScale AS w\015\012WHERE (\015\012[activeonly] = 1"
    "\015\012AND\015\012w.ActiveYear <= [yractive]\015\012AND\015\012w.RetireYear IS "
    "NULL\015\012)\015\012OR\015\012(\015\012[activeonly] = 0\015\012AND\015\012([yra"
    "ctive]  BETWEEN w.ActiveYear AND w.RetireYear\015\012OR\015\012w.ActiveYear <=[y"
    "ractive]\015\012)\015\012)\015\012ORDER BY w.RetireYear, w.CategoryOrder;\015\012"
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
        dbText "Name" ="s.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.IsActiveForProtocol"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteDirections"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1006"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Site"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4c5e233e1de98045a4bfd0f54e84bf02
        End
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SiteName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SiteDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SiteDirections"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.CollectionSourceName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.River_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Park_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="w.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="w.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="w.Label"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="w.StartYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="w.RetireYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="w.DiameterRange_mm"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="w.ActiveYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="category"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcc901dae9bb169408afb344a68510b73
        End
    End
End
