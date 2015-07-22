Operation =1
Option =0
Where ="(((qry_Transect_Select.Unit_Code)=[Forms]![frm_Monitoring_Transect]![Park_Code])"
    " AND ((qry_Transect_Select.Visit_Year)=[Forms]![frm_Monitoring_Transect]![Visit_"
    "Year]) AND ((qry_Transect_Select.Species) Is Not Null))"
Begin InputTables
    Name ="qry_Transect_Select"
End
Begin OutputColumns
    Expression ="qry_Transect_Select.Unit_Code"
    Expression ="qry_Transect_Select.Visit_Year"
    Expression ="qry_Transect_Select.Plot_ID"
    Expression ="qry_Transect_Select.Transect"
    Expression ="qry_Transect_Select.Area"
    Expression ="qry_Transect_Select.Species"
    Expression ="qry_Transect_Select.Master_Common_Name"
    Alias ="Cover_Average"
    Expression ="IIf([Visit_Year]=2008,([Q1]+[Q2]+[Q3])/3,IIf([Visit_Year]=2009,([Q1_3m]+[Q2_8m]+"
        "[Q3_13m])/3,([Q1_hm]+[Q2_5m]+[Q3_10m])/3))"
    Expression ="qry_Transect_Select.E_Coord"
    Expression ="qry_Transect_Select.N_Coord"
End
Begin OrderBy
    Expression ="qry_Transect_Select.Plot_ID"
    Flag =0
    Expression ="qry_Transect_Select.Transect"
    Flag =0
    Expression ="qry_Transect_Select.Species"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xbf25c4aaf0564d4c993b43e86c3de99d
End
Begin
    Begin
        dbText "Name" ="qry_Transect_Select.Unit_Code"
        dbInteger "ColumnWidth" ="1050"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Transect_Select.Visit_Year"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Transect_Select.Plot_ID"
        dbInteger "ColumnWidth" ="2520"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Transect_Select.Transect"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Transect_Select.Area"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Transect_Select.Species"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =13
    Top =124
    Right =999
    Bottom =448
    Left =-1
    Top =-1
    Right =967
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =103
        Top =5
        Right =297
        Bottom =123
        Top =0
        Name ="qry_Transect_Select"
        Name =""
    End
End
