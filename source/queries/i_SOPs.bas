dbMemo "SQL" ="SELECT SOP, SOPNum, Code, Version, EffectiveDate FROM SOP2011\015\012UNION\015\012"
    "SELECT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2012\015\012UNION\015\012"
    "SELECT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2013\015\012UNION\015\012"
    "SELECT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2014\015\012UNION\015\012"
    "SELECT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2015\015\012UNION\015\012"
    "SELECT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2016\015\012UNION SELE"
    "CT SOP, SOPNum, Code, Version, EffectiveDate  FROM SOP2017a\015\012ORDER BY Effe"
    "ctiveDate, SOPNum;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd19876b21bbcf74c967a724261f1aba6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="SOP"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x33a7d71b7440b94686883230b43bb6c1
        End
    End
    Begin
        dbText "Name" ="SOPNum"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6e2ee7411b039e48a126be952e3f6b90
        End
    End
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8423d9918184a542a4014766609d6da6
        End
    End
    Begin
        dbText "Name" ="Version"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2eee22af20fa9248ad7d857277991f0b
        End
    End
    Begin
        dbText "Name" ="EffectiveDate"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9f1a423710874049b570335092418e4d
        End
    End
End
