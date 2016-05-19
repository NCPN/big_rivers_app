dbMemo "SQL" ="SELECT Utah_Species AS Species, ListedSpecies.LU_Code\015\012FROM ListedSpecies "
    "LEFT JOIN tlu_NCPN_Plants ON tlu_NCPN_Plants.LU_Code = ListedSpecies.LU_Code\015"
    "\012WHERE RiverSegment_ID = 1\015\012AND\015\012FieldSeason = Year(Now)\015\012A"
    "ND\015\012CoverType = \"WCC\"\015\012ORDER BY Utah_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4aa33f857fae6d41984af2f99e65c481
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ListedSpecies.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf9e182b4baf8a04a9ec13a6a5b8916d4
        End
    End
End
