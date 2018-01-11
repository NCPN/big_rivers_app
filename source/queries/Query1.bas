dbMemo "SQL" ="SELECT t.*, ae.Label AS Priority, aes.Label AS Status\015\012FROM (Task AS t INN"
    "ER JOIN AppEnum AS ae ON t.Priority_ID = ae.ID) INNER JOIN AppEnum AS aes ON t.S"
    "tatus_ID = aes.ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="t.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TaskType"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TaskType_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Task"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Status_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Priority_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.RequestedBy_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.LastModifiedBy_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.LastModified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ae.Label"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Priority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.RequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.CompletedBy_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.CompleteDate"
        dbLong "AggregateType" ="-1"
    End
End
