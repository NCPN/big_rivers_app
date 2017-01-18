dbMemo "SQL" ="INSERT INTO SOP ( FullName, SOPNumber, Code, Version, EffectiveDate, CreatedBy_I"
    "D, LastModifiedBy_ID )\015\012SELECT SOP, SOPNum, Code, Version, EffectiveDate, "
    "1, 1\015\012FROM i_SOPs;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x49e8f2e5edf3ca41a0a2b047e01740b5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="SOP"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9501805a6bbfa545804a8da62a851922
        End
    End
    Begin
        dbText "Name" ="SOPNum"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb20522f440cd0b49a38f31d96f4fd4c4
        End
    End
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2accd088f5325d408f0b27f17c9b1d44
        End
    End
    Begin
        dbText "Name" ="Version"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbd8fcc9d5679e5498ed5460b2f63a7e7
        End
    End
    Begin
        dbText "Name" ="EffectiveDate"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3b8c47ac7d9c2540896358c8dc6376fe
        End
    End
End
