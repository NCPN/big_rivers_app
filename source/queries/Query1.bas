dbMemo "SQL" ="SELECT ID, Context, Template, Remarks\015\012FROM tsys_Db_Templates\015\012WHERE"
    " Template LIKE '*vegplot*';\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbInteger "RowHeight" ="4230"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tsys_Db_Templates.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Version"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.IsSupported"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.FieldCheck"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.FieldOK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Context"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Syntax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.TemplateName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Context"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Template"
        dbInteger "ColumnWidth" ="7215"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Remarks"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.DataScope"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Params"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Template"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Remarks"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Dependencies"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.EffectiveDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.RetireDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.CreateDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.CreatedBy_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.LastModified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.LastModifiedBy_ID"
        dbLong "AggregateType" ="-1"
    End
End
