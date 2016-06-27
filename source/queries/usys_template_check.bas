dbMemo "SQL" ="SELECT Count(*) AS cnt, tsys_Db_Templates.TemplateName, Template, EffectiveDate\015"
    "\012FROM tsys_Db_Templates\015\012GROUP BY TemplateName, Template, EffectiveDate"
    ";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbInteger "RowHeight" ="270"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x285bcd0b45514445aa5049d69341e0e0
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TemplateName"
        dbInteger "ColumnWidth" ="3870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cnt"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x95e2b9916d0efe49978831528b5ae2f1
        End
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.TemplateName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x04f166bd97c90a4dbed7e5594e533315
        End
    End
    Begin
        dbText "Name" ="Template"
        dbInteger "ColumnWidth" ="8940"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x861aa2ad1c86444891a059dbdd19adbf
        End
    End
    Begin
        dbText "Name" ="EffectiveDate"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x67584be957b8ef478077e95482442903
        End
    End
End
