dbMemo "SQL" ="SELECT DISTINCT qry_Annual_Complete_Tgt_Species_Lists.TgtYear, qry_Annual_Comple"
    "te_Tgt_Species_Lists.Master_Plant_Code_FK, qry_Annual_Complete_Tgt_Species_Lists"
    ".LU_Code, qry_Annual_Complete_Tgt_Species_Lists.Family, qry_Annual_Complete_Tgt_"
    "Species_Lists.Species_Name, qry_Annual_Complete_Tgt_Species_Lists.utah_species, "
    "qry_Annual_Complete_Tgt_Species_Lists.Co_Species, qry_Annual_Complete_Tgt_Specie"
    "s_Lists.Wy_Species, qry_Annual_Complete_Tgt_Species_Lists.Master_Common_Name, Co"
    "ncatRelated(\"ParkPriority\",\"qry_Annual_Complete_Tgt_Species_Lists\",\"Species"
    "Year='\"+SpeciesYear+\"'\",'',\"|\") AS ParkPriorities\015\012FROM qry_Annual_Co"
    "mplete_Tgt_Species_Lists;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x66d43e01165ab44589186a3091b85e33
End
dbText "Description" ="Target species summary for all parks for a given year  (Target List Tool update)"
Begin
    Begin
        dbText "Name" ="ParkPriorities"
        dbInteger "ColumnWidth" ="3744"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd9889c2e6375c048961504abd97141f5
        End
    End
    Begin
        dbText "Name" ="qry_Annual_Complete_Tgt_Species_Lists.TgtYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Annual_Complete_Tgt_Species_Lists.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Annual_Complete_Tgt_Species_Lists.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Annual_Complete_Tgt_Species_Lists.Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Annual_Complete_Tgt_Species_Lists.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Annual_Complete_Tgt_Species_Lists.utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Annual_Complete_Tgt_Species_Lists.Co_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Annual_Complete_Tgt_Species_Lists.Wy_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Annual_Complete_Tgt_Species_Lists.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
End
