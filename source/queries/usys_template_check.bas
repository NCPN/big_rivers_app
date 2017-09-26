dbMemo "SQL" ="SELECT Count(*) AS cnt, tsys_Db_Templates.TemplateName, Template, Params, Remark"
    "s, EffectiveDate\015\012FROM tsys_Db_Templates\015\012GROUP BY TemplateName, Tem"
    "plate, Params, Remarks, EffectiveDate;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbInteger "RowHeight" ="270"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="cnt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.TemplateName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Template"
        dbInteger "ColumnWidth" ="8940"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EffectiveDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Params"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remarks"
        dbLong "AggregateType" ="-1"
    End
End
