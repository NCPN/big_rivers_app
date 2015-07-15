dbMemo "SQL" ="PARAMETERS park Text ( 255 ), tgtYear Short;\015\012SELECT tbl_Target_Species.*,"
    " tbl_Target_Species.Park_Code, tbl_Target_Species.Target_Year, *\015\012FROM tbl"
    "_Target_Species\015\012WHERE (((tbl_Target_Species.Target_Year)=CInt(tgtYear)) A"
    "nd ((LCase(tbl_Target_Species.Park_Code))=LCase(park)))\015\012ORDER BY tbl_Targ"
    "et_Species.Species_Name;\015\012"
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
    0x0ba4512c4d2ef849a44e1d7cc2d6bc08
End
dbText "Description" =" Target species list for a park & year\015\012(Target List Tool update)"
Begin
    Begin
        dbText "Name" ="tbl_Target_Species.Tgt_Species_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1724a816b11f30498639fd773f52483c
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3feb71a96e5fff41a1355b8adaa86516
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Park_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Target_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Species_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x37e209f0931b624e9e27c4382eaa68c5
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Priority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8197c802f29b3a4a8ea780fab958b4db
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Transect_Only"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xafa335647e1cb74da609c29c11406aa8
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Target_Area_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2a29cb3f9b5f8a4496ec6fd254749fcd
        End
    End
End
