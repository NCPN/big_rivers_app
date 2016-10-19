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
