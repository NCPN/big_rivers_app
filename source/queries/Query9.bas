dbMemo "SQL" ="PARAMETERS tbl Text ( 255 ), flds Text ( 255 ), ident Long;\015\012SELECT [flds]"
    "\015\012FROM tbl\015\012WHERE ID = [ident];\015\012"
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
        dbText "Name" ="s.SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.StartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.ParkCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="f.Feature"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.FilamentousAlgae"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.PlotNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.PlotDistance_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.ModalSedimentSize"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.PercentFine"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.PlotDensity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.NoCanopyVeg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.NoRootedVeg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.HasSocialTrail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.NoIndicatorSpecies"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vt.TransectNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vt.SampleDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.PercentWater"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.UnderstoryRootedPctCover"
        dbLong "AggregateType" ="-1"
    End
End
