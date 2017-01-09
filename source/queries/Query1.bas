dbMemo "SQL" ="PARAMETERS eventyr Long;\015\012SELECT w.ID, w.Label, w.Code, w.ActiveYear, w.Di"
    "ameterRange_mm, w.RetireYear, w.Label + ' ('+ w.Code +')' AS category\015\012FRO"
    "M ModWentworthScale AS w\015\012WHERE (\015\012(w.ActiveYear = [eventyr]) \015\012"
    "OR\015\012(w.RetireYear = [eventyr])\015\012OR\015\012(w.ActiveYear <[eventyr]) "
    "\015\012AND \015\012((w.RetireYear IS NULL) OR ([eventyr] < w.RetireYear))\015\012"
    ")\015\012ORDER BY w.CategoryOrder;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xff7ad948611d5b49b1b51c1013b470b2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
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
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="w.ActiveYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="w.RetireYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="w.DiameterRange_mm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="category"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="'Size Class'"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
