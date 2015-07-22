dbMemo "SQL" ="SELECT tbl_Target_List.Park_Code AS Park, tbl_Target_List.Target_Year AS TgtYear"
    ", tbl_Target_Species.LU_Code, tbl_Target_Species.Master_Plant_Code_FK, tbl_Targe"
    "t_Species.Species_Name, tbl_Target_Species.Priority, tbl_Target_Species.Transect"
    "_Only, tbl_Target_Species.Target_Area_ID, tbl_Target_Areas.Target_Area AS Tgt_Ar"
    "ea, tlu_NCPN_Plants.Master_Family AS Family, tlu_NCPN_Plants.Master_Common_Name,"
    " tlu_NCPN_Plants.Utah_Species, tlu_NCPN_Plants.Co_Species, tlu_NCPN_Plants.Wy_Sp"
    "ecies, tbl_Target_List.Park_Code & \"-\" & tbl_Target_List.Target_Year AS TgtLis"
    "t, tbl_Target_List.Created, tbl_Target_List.Last_Modified\015\012FROM ((tbl_Targ"
    "et_Species LEFT JOIN tbl_Target_Areas ON tbl_Target_Species.Target_Area_ID = tbl"
    "_Target_Areas.Target_Area_ID) LEFT JOIN tbl_Target_List ON tbl_Target_Species.Tg"
    "t_List_ID_FK = tbl_Target_List.Tgt_List_ID) LEFT JOIN tlu_NCPN_Plants ON tbl_Tar"
    "get_Species.Master_Plant_Code_FK = tlu_NCPN_Plants.Master_PLANT_Code;\015\012"
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
    0x130bab6b315e3049a7418bad3e6bf946
End
dbText "Description" ="Park target species listings including priority, transect_only, and target_area "
    "(Target List Tool update)"
Begin
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Co_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x640ed8196c68c54f9a7c743f22ce8780
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Wy_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x329ae115eb7c774d9a76093144a0fff0
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.utah_species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x111986e05779da419826e014adaa7dcb
        End
    End
    Begin
        dbText "Name" ="Park"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x708a1aedb2467c42b189b5fb31b2f2db
        End
    End
    Begin
        dbText "Name" ="TgtYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc2858dba0326f14385c4bef8b3cae300
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Target_Area_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9edc1f27fa94a44b86fbf4808ba68c0c
        End
    End
    Begin
        dbText "Name" ="Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe043677f6a6fda4cae0c1ca189be4dba
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x430b7b36751bf946aea7cb3ebf49f154
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Species_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa5b83c2eec2e9a43ac5c9a48b23ad750
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Priority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x22730abafa393842b97152c594b52b90
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Transect_Only"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x625706e520ed394c9c40ff23350280c0
        End
    End
    Begin
        dbText "Name" ="Tgt_Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x96a55420d4a01344ad51cd89f5fc5426
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8cec9c3064f2544d87b6cf301d0487c7
        End
    End
    Begin
        dbText "Name" ="TgtList"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ccfa02359522645aa4237d398990945
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.LU_code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_List.Created"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_List.Last_Modified"
        dbLong "AggregateType" ="-1"
    End
End
