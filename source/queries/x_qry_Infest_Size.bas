Operation =1
Option =0
Where ="(((qry_Infest_Size_Select.Size_Class) Is Not Null))"
Begin InputTables
    Name ="qry_Infest_Size_Select"
    Name ="tbl_Target_Plant_Lists"
End
Begin OutputColumns
    Expression ="qry_Infest_Size_Select.*"
    Expression ="tbl_Target_Plant_Lists.Priority"
End
Begin Joins
    LeftTable ="qry_Infest_Size_Select"
    RightTable ="tbl_Target_Plant_Lists"
    Expression ="qry_Infest_Size_Select.Master_Code = tbl_Target_Plant_Lists.Master_Plant_Code"
    Flag =2
    LeftTable ="qry_Infest_Size_Select"
    RightTable ="tbl_Target_Plant_Lists"
    Expression ="qry_Infest_Size_Select.Unit_Code = tbl_Target_Plant_Lists.Unit_Code"
    Flag =2
    LeftTable ="qry_Infest_Size_Select"
    RightTable ="tbl_Target_Plant_Lists"
    Expression ="qry_Infest_Size_Select.Visit_Year = tbl_Target_Plant_Lists.Visit_Year"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf5e8c4a2253ea049b24db8b8833f7d52
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =18
    Top =14
    Right =1318
    Bottom =338
    Left =-1
    Top =-1
    Right =1268
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =200
        Bottom =109
        Top =0
        Name ="qry_Infest_Size_Select"
        Name =""
    End
    Begin
        Left =260
        Top =3
        Right =446
        Bottom =106
        Top =0
        Name ="tbl_Target_Plant_Lists"
        Name =""
    End
End
