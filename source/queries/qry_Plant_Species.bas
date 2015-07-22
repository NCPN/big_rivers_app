dbMemo "SQL" ="SELECT tlu_NCPN_Plants.Master_PLANT_Code AS Code, tlu_NCPN_Plants.Master_Species"
    " AS Species, Switch(tlu_NCPN_Plants.LU_Code Is Null,\" \",tlu_NCPN_Plants.LU_Cod"
    "e<>\"\",tlu_NCPN_Plants.LU_Code) AS LUCode\015\012FROM tlu_NCPN_Plants\015\012OR"
    "DER BY tlu_NCPN_Plants.Master_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0xdb6256e17322bf4a8f3ee2ec1a17baea
End
Begin
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbd93ec5fffe281499ceb78908087bc23
        End
        dbInteger "ColumnWidth" ="1704"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xabe9b57fcb7ef64598b255922b8dac2f
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LUCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x89f8cd1b07547b45892fc1da691ff909
        End
    End
End
