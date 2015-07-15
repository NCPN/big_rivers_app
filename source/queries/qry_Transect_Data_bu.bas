Operation =1
Option =0
Where ="(((tbl_Locations.Unit_Code)=Forms!frm_Monitoring_Transect!Park_Code) And ((Year("
    "[Start_Date]))=Forms!frm_Monitoring_Transect!Visit_Year) And ((IIf([Unit_Code] I"
    "n (\"CARE\",\"DINO\",\"GOSP\"),[Utah_Species],IIf([Unit_Code]=\"FOBU\",[WY_Speci"
    "es],[Co_Species]))) Is Not Null))"
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
    Expression ="IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\"),[Utah_Species],IIf([Unit_Code]=\""
        "FOBU\",[WY_Species],[Co_Species]))"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
    Alias ="Cover_Average"
    Expression ="IIf([Visit_Year]=2008,([Q1]+[Q2]+[Q3])/3,IIf([Visit_Year]=2009,([Q1_3m]+[Q2_8m]+"
        "[Q3_13m])/3,([Q1_hm]+[Q2_5m]+[Q3_10m])/3))"
    Expression ="tbl_Quadrat_Transect.E_Coord"
    Expression ="tbl_Quadrat_Transect.N_Coord"
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
    Expression ="IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\"),[Utah_Species],IIf([Unit_Code]=\""
        "FOBU\",[WY_Species],[Co_Species]))"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x489334f9bad7fb40a441d8d7710d72a1
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
    End
End
Begin
    State =0
    Left =13
    Top =150
    Right =999
    Bottom =474
    Left =-1
    Top =-1
    Right =971
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
        Top =3
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
