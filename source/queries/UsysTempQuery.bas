Operation =1
Option =0
Having ="(((Count([tsys_Db_Templates].[TemplateName]))>1))"
Begin InputTables
    Name ="tsys_Db_Templates"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tsys_Db_Templates.TemplateName"
    Alias ="NumberOfDupes"
    Expression ="Count(tsys_Db_Templates.TemplateName)"
End
Begin Groups
    Expression ="tsys_Db_Templates.TemplateName"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
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
    End
    Begin
        dbText "Name" ="Expr1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =867
    Bottom =776
    Left =-1
    Top =-1
    Right =851
    Bottom =152
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
