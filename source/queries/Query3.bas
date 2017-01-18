dbMemo "SQL" ="TRANSFORM FIRST(Version)\015\012SELECT SOPNumber, FullName, Code, Version, Effec"
    "tiveDate, RetireDate, Year(EffectiveDate) AS StartYear, Year(RetireDate) AS EndY"
    "ear\015\012FROM SOP\015\012GROUP BY SOPNumber, FullName, Code, Version, Effectiv"
    "eDate, RetireDate\015\012PIVOT FullName;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6678ccd6451dcf4eaef01819f0b74801
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Version"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EffectiveDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SOP.SOPNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="After each field season"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="After each field visit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLCA field methods"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CANY & DINO equip_ lists"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CANY field methods"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DINO measuring vegetation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Facies mapping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GPS methods"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Prior to field season/equip_ lists"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sentinel site set up"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLCA field methods, DINO sentinel site set up"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CURE field methods"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Data analysis & reporting"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Data management"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rapid assessment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Revising the protocol"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Surveying"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Total station surveying"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Training observers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StartYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EndYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CURE methods"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Facies mapping & grain size dist_"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RTK surveying part 1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vegetation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SOPNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RetireDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DINO facies mapping & grain size dist_"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DINO field methods"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hydrologic measurements"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Measuring vegetation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remote sensing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RTK surveying part 2"
        dbLong "AggregateType" ="-1"
    End
End
