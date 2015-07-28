Operation =1
Option =0
Begin InputTables
    Name ="tlu_Parks"
    Name ="Site"
    Name ="xref_Site_Feature"
    Name ="Features"
End
Begin OutputColumns
    Expression ="tlu_Parks.ParkCode"
    Expression ="tlu_Parks.ParkName"
    Expression ="Site.Site_ID"
    Expression ="Site.Site_code"
    Expression ="Site.Site_name"
    Expression ="Features.Feature_ID"
    Expression ="Features.Feature"
    Expression ="Features.Feature_description"
End
Begin Joins
    LeftTable ="tlu_Parks"
    RightTable ="Site"
    Expression ="tlu_Parks.ParkCode = Site.Unit_code"
    Flag =1
    LeftTable ="Site"
    RightTable ="xref_Site_Feature"
    Expression ="Site.Site_ID = xref_Site_Feature.Site_FK"
    Flag =1
    LeftTable ="Features"
    RightTable ="xref_Site_Feature"
    Expression ="Features.Feature_ID = xref_Site_Feature.Feature_FK"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc26502348e872a479d422115e5fa6eb5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tlu_Parks.ParkCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc1c73150e3c57445b5b358601eff5182
        End
    End
    Begin
        dbText "Name" ="tlu_Parks.ParkName"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdbc301166002dc49a3483698ca528b0e
        End
    End
    Begin
        dbText "Name" ="Site.Site_code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xef6a9f354dca084ab4e9cef356d5ea34
        End
    End
    Begin
        dbText "Name" ="Site.Site_ID"
        dbInteger "ColumnWidth" ="3330"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x288b859f6cafd94d90b9d1848545a23f
        End
    End
    Begin
        dbText "Name" ="Site.Site_name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcd97e0c800f9664ca73b90262b80b09e
        End
    End
    Begin
        dbText "Name" ="Features.Feature_ID"
        dbInteger "ColumnWidth" ="3330"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x80767fab41c6954c8e9e950ac9fbeba2
        End
    End
    Begin
        dbText "Name" ="Features.Feature"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd5bb4189a89ae64c958cf6a99b85f970
        End
    End
    Begin
        dbText "Name" ="Features.Feature_description"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7c36458645d98446906983514334cd4d
        End
    End
End
Begin
    State =0
    Left =8
    Top =65
    Right =956
    Bottom =783
    Left =-1
    Top =-1
    Right =916
    Bottom =418
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =28
        Top =6
        Right =172
        Bottom =150
        Top =0
        Name ="tlu_Parks"
        Name =""
    End
    Begin
        Left =220
        Top =12
        Right =364
        Bottom =156
        Top =0
        Name ="Site"
        Name =""
    End
    Begin
        Left =412
        Top =12
        Right =556
        Bottom =156
        Top =0
        Name ="xref_Site_Feature"
        Name =""
    End
    Begin
        Left =604
        Top =12
        Right =748
        Bottom =156
        Top =0
        Name ="Features"
        Name =""
    End
End
