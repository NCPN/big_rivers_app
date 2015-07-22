dbMemo "SQL" ="SELECT qry_Infest_Size_Select.*, tbl_Target_Species.Priority\015\012FROM qry_Inf"
    "est_Size_Select LEFT JOIN tbl_Target_Species ON (qry_Infest_Size_Select.Master_C"
    "ode = tbl_Target_Species.Master_Plant_Code_FK) AND (qry_Infest_Size_Select.Unit_"
    "Code = tbl_Target_Species.Park_Code) AND (qry_Infest_Size_Select.Visit_Year = tb"
    "l_Target_Species.Target_Year)\015\012WHERE (((qry_Infest_Size_Select.Size_Class)"
    " Is Not Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x29e685a9421a354390ecb3bad41b7881
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Infestation.Pulled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Infestation.Growth_Stage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Infestation.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Infestation.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tlu_Size_Class.Size_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Infestation.Master_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Priority"
        dbLong "AggregateType" ="-1"
    End
End
