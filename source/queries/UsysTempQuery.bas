Operation =1
Option =0
Having ="Count(TemplateName) > 1"
Begin InputTables
    Name ="tsys_Db_Templates"
End
Begin OutputColumns
    Expression ="TemplateName"
    Alias ="NumberOfDupes"
    Expression ="Count(TemplateName)"
End
Begin Groups
    Expression ="TemplateName"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xdfa89f91ad4e0840acbd7c389f87a880
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TemplateName"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumberOfDupes"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x21a050c0b9ffe845b485f1f448661b92
        End
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1179
    Bottom =750
    Left =-1
    Top =-1
    Right =1163
    Bottom =203
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tsys_Db_Templates"
        Name =""
    End
End
