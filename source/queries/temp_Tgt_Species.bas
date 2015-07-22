dbMemo "SQL" ="SELECT tbl_Target_List.Park_Code AS Park, tbl_Target_List.Target_Year AS TgtYear"
    ", Master_Plant_Code_FK, Species_Name, LU_Code, Priority, Transect_Only, Target_A"
    "rea_ID\015\012FROM tbl_Target_Species INNER JOIN tbl_Target_List ON tbl_Target_S"
    "pecies.Tgt_List_ID_FK = tbl_Target_List.Tgt_List_ID\015\012WHERE (((tbl_Target_L"
    "ist.Target_Year) = CInt(2016)) And ((LCase([tbl_Target_List].[Park_Code])) = LCa"
    "se('CURE')))\015\012ORDER BY tbl_Target_Species.Species_Name;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
