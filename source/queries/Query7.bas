dbMemo "SQL" ="SELECT tlu_NCPN_Plants.Master_Species, tlu_NCPN_Plants.Utah_species, tlu_NCPN_Pl"
    "ants.LU_Code, tlu_NCPN_Plants.Master_family, tlu_NCPN_Plants.UT_family, tlu_NCPN"
    "_Plants.Master_PLANT_Code\015\012FROM tlu_NCPN_Plants\015\012ORDER BY tlu_NCPN_P"
    "lants.LU_Code;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "OrderBy" ="[Query7].[UT_family], [Query7].[Master_family], [Query7].[LU_Code], [Query7].[Ma"
    "ster_Species]"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe60f75e4308b6a43ae590d15d0ffbe04
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.UT_family"
        dbLong "AggregateType" ="-1"
    End
End
