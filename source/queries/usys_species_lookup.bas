dbMemo "SQL" ="SELECT Master_Species, LU_Code, Utah_Species\015\012FROM tlu_NCPN_Plants\015\012"
    "WHERE LU_Code IS NOT NULL OR LU_Code IS NULL\015\012ORDER BY Master_Species;\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "OrderBy" ="[usys_species_lookup].[LU_Code], [usys_species_lookup].[Master_Species]"
Begin
    Begin
        dbText "Name" ="LU_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Utah_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master_Species"
        dbLong "AggregateType" ="-1"
    End
End
