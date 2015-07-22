Operation =1
Option =0
Where ="(((IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],IIf(["
    "Unit_Code]=\"FOBU\",[WY_Species],[Co_Species]))) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Transect"
    Name ="tbl_Quadrat_Species"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Quadrat_Transect.Transect"
    Expression ="tbl_Locations.Area"
    Alias ="Species"
    Expression ="IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],IIf([Uni"
        "t_Code]=\"FOBU\",[WY_Species],[Co_Species]))"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
    Expression ="tbl_Quadrat_Transect.E_Coord"
    Expression ="tbl_Quadrat_Transect.N_Coord"
    Alias ="Q1_hm"
    Expression ="IIf(IsNull(tbl_Quadrat_Species.Q1_hm),0,tbl_Quadrat_Species.Q1_hm)"
    Alias ="Q2_5m"
    Expression ="IIf(IsNull(tbl_Quadrat_Species.Q2_5m),0,tbl_Quadrat_Species.Q2_5m)"
    Alias ="Q3_10m"
    Expression ="IIf(IsNull(tbl_Quadrat_Species.Q3_10m),0,tbl_Quadrat_Species.Q3_10m)"
    Alias ="Q1_3m"
    Expression ="IIf(IsNull(tbl_Quadrat_Species.Q1_3m),0,tbl_Quadrat_Species.Q1_3m)"
    Alias ="Q2_8m"
    Expression ="IIf(IsNull(tbl_Quadrat_Species.Q2_8m),0,tbl_Quadrat_Species.Q2_8m)"
    Alias ="Q3_13m"
    Expression ="IIf(IsNull(tbl_Quadrat_Species.Q3_13m),0,tbl_Quadrat_Species.Q3_13m)"
    Alias ="Q1"
    Expression ="IIf(IsNull(tbl_Quadrat_Species.Q1),0,tbl_Quadrat_Species.Q1)"
    Alias ="Q2"
    Expression ="IIf(IsNull(tbl_Quadrat_Species.Q2),0,tbl_Quadrat_Species.Q2)"
    Alias ="Q3"
    Expression ="IIf(IsNull(tbl_Quadrat_Species.Q3),0,tbl_Quadrat_Species.Q3)"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_Quadrat_Transect"
    Expression ="tbl_Events.Event_ID=tbl_Quadrat_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_Quadrat_Species"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_Quadrat_Species.Plant_Code=tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
    LeftTable ="tbl_Quadrat_Transect"
    RightTable ="tbl_Quadrat_Species"
    Expression ="tbl_Quadrat_Transect.Transect_ID=tbl_Quadrat_Species.Transect_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="tbl_Quadrat_Transect.Transect"
    Flag =0
    Expression ="IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],IIf([Uni"
        "t_Code]=\"FOBU\",[WY_Species],[Co_Species]))"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x133c6bcdab25ec4eb1d1c02d769aa2d3
End
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbInteger "ColumnWidth" ="1050"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
        dbBinary "GUID" = Begin
            0x62c44a96f682974eb1392781ba0c2b05
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbInteger "ColumnWidth" ="2520"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Transect.Transect"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Area"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Species"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbBinary "GUID" = Begin
            0x1a35b91843572b4585fc45dc92c4987a
        End
    End
    Begin
        dbText "Name" ="Q1_hm"
        dbBinary "GUID" = Begin
            0x962410f264a2cc459bcf8a5a892e9341
        End
    End
    Begin
        dbText "Name" ="Q2_5m"
        dbBinary "GUID" = Begin
            0xc2ae12c155aee74c9231a4be3b322bcb
        End
    End
    Begin
        dbText "Name" ="Q3_10m"
        dbBinary "GUID" = Begin
            0xeee4dc51e4c1c44d8ac509ae4eed3ee4
        End
    End
    Begin
        dbText "Name" ="Q1_3m"
        dbBinary "GUID" = Begin
            0x8b1bdeddf2dad148847762810ab6b16c
        End
    End
    Begin
        dbText "Name" ="Q2_8m"
        dbBinary "GUID" = Begin
            0x9d80194b9caab245a4cc2b9fb3271ed2
        End
    End
    Begin
        dbText "Name" ="Q3_13m"
        dbBinary "GUID" = Begin
            0xe99ddd91776828448345580fb98e003f
        End
    End
    Begin
        dbText "Name" ="Q1"
        dbBinary "GUID" = Begin
            0x497373d379d92148ab1f891d39e2ae31
        End
    End
    Begin
        dbText "Name" ="Q2"
        dbBinary "GUID" = Begin
            0x47d39a19113139459d97a6c119f47ea8
        End
    End
    Begin
        dbText "Name" ="Q3"
        dbBinary "GUID" = Begin
            0xcec45e49da1f45459d14814c86c679b3
        End
    End
End
Begin
    State =0
    Left =54
    Top =98
    Right =1124
    Bottom =422
    Left =-1
    Top =-1
    Right =1055
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =109
        Top =1
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =109
        Top =1
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =109
        Top =0
        Name ="tbl_Quadrat_Transect"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =109
        Top =12
        Name ="tbl_Quadrat_Species"
        Name =""
    End
    Begin
        Left =574
        Top =6
        Right =670
        Bottom =109
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
