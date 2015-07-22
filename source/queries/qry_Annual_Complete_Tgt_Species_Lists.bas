dbMemo "SQL" ="SELECT DISTINCT tbl_Target_List.Park_Code AS Park, tbl_Target_List.Target_Year A"
    "S TgtYear, tbl_Target_Species.Master_Plant_Code_FK, tlu_NCPN_Plants.LU_Code, tbl"
    "_Target_Species.Species_Name, tbl_Target_Species.Priority, tbl_Target_Species.Tr"
    "ansect_Only, tbl_Target_Species.Target_Area_ID, tbl_Target_Areas.Target_Area AS "
    "Tgt_Area, tlu_NCPN_Plants.Master_Family AS Family, tlu_NCPN_Plants.Master_Common"
    "_Name, tlu_NCPN_Plants.utah_species, tlu_NCPN_Plants.Co_Species, tlu_NCPN_Plants"
    ".Wy_Species, IIf(tbl_Target_Species.Target_Area_ID>0,tbl_Target_Areas.Target_Are"
    "a,IIf\015\015\012(tbl_Target_Species.Transect_Only>0,\"Transect\",tbl_Target_Spe"
    "cies.Priority)) AS PriorityTarget, (tbl_Target_List.Park_Code+\"-\"+PriorityTarg"
    "et) AS ParkPriority, (tbl_Target_Species.Species_Name+\"-\"+CStr(tbl_Target_List"
    ".Target_Year)) AS SpeciesYear\015\012FROM ((tbl_Target_Species LEFT JOIN tbl_Tar"
    "get_Areas ON tbl_Target_Species.Target_Area_ID = tbl_Target_Areas.Target_Area_ID"
    ") LEFT JOIN tbl_Target_List ON tbl_Target_Species.Tgt_List_ID_FK = tbl_Target_Li"
    "st.Tgt_List_ID) LEFT JOIN tlu_NCPN_Plants ON tbl_Target_Species.Master_Plant_Cod"
    "e_FK = tlu_NCPN_Plants.Master_Plant_Code\015\012ORDER BY tbl_Target_Species.Spec"
    "ies_Name;\015\012"
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
    0x1b80d8ab936a2945a2bd72782b6a1882
End
dbText "Description" ="complete list of target species for a given year\015\012 (Target List Tool updat"
    "e)"
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
        dbText "Name" ="PriorityTarget"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc72926b730f46b408df0e1f6bf6e67f5
        End
    End
    Begin
        dbText "Name" ="ParkPriority"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf38e7f541e6a234daa9759d57010db2f
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpeciesYear"
        dbInteger "ColumnWidth" ="2208"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7597b113575789469ca53d2a7dc20e54
        End
    End
End
